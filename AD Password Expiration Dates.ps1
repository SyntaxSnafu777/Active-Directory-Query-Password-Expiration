# Import the Active Directory module with error handling
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Host "Failed to import ActiveDirectory module. Ensure RSAT is installed." -ForegroundColor Red
    exit
}

# Retrieve the domain's default password policy to get MaxPasswordAge
try {
    $PasswordPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
    $MaxPasswordAge = $PasswordPolicy.MaxPasswordAge
}
catch {
    Write-Host "Failed to retrieve password policy. Ensure you have the necessary permissions." -ForegroundColor Red
    exit
}

# Function to extract the Top-Level OU from DistinguishedName
function Get-TopLevelOU {
    param (
        [string]$DistinguishedName
    )
    # Split the DistinguishedName into its components
    $components = $DistinguishedName -split ","
    # Find all OU components
    $ous = $components | Where-Object { $_ -like "OU=*" }
    if ($ous.Count -gt 0) {
        # Extract the first OU (Top-Level OU)
        return ($ous[0] -replace "^OU=", "")
    }
    else {
        return "No OU Found"
    }
}

# Initialize an empty array for users
$UserList = @()

# Inform the user about the script's options
Write-Host ""
Write-Host "This script will check password expiration for Active Directory users." -ForegroundColor Cyan
Write-Host ""
Write-Host "You have the following options to filter users:" -ForegroundColor Cyan
Write-Host ""
Write-Host "Check a specific Organizational Unit (OU)" -ForegroundColor Cyan
Write-Host "Check a specific Active Directory (AD) Group" -ForegroundColor Cyan
Write-Host "Check all active accounts in AD" -ForegroundColor Cyan
Write-Host ""
Write-Host "Please make your selections below:" -ForegroundColor Yellow
Write-Host ""

# Prompt the user to decide whether to check a specific OU
$checkOU = Read-Host "Would you like to check a specific Organizational Unit (OU)? (Y/N)"
Write-Host ""

if ($checkOU.Trim().ToUpper() -eq 'Y') {
    # Prompt for the OU name
    $ouInput = Read-Host "Enter the distinguished name of the OU (e.g., OU=Marketing,DC=contoso,DC=com)"

    # Attempt to retrieve the OU
    try {
        $OU = Get-ADOrganizationalUnit -Filter { DistinguishedName -eq $ouInput } -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to find the OU with distinguished name '$ouInput'. Please ensure the format is correct." -ForegroundColor Red
        exit
    }

    # Retrieve all enabled users within the specified OU
    try {
        $UserList = Get-ADUser -Filter { Enabled -eq $True } -SearchBase $OU.DistinguishedName -Properties PasswordLastSet, PasswordNeverExpires, DistinguishedName -ErrorAction Stop | 
            Select-Object Name, SamAccountName, DistinguishedName,
                @{Name='TopLevelOU'; Expression={ Get-TopLevelOU $_.DistinguishedName }},
                @{Name='PasswordLastSet'; Expression={$_.PasswordLastSet}},
                @{Name='PasswordExpiryDate'; Expression={
                    if ($_.PasswordNeverExpires) {
                        $null  # Use $null for 'Never' to facilitate sorting
                    }
                    else {
                        $_.PasswordLastSet + $MaxPasswordAge
                    }
                }}
        
        if ($UserList.Count -eq 0) {
            Write-Host "No enabled users found in the OU '$ouInput'." -ForegroundColor Yellow
            exit
        }

        Write-Host "Retrieved users from OU '$ouInput'." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to retrieve users from the OU '$ouInput'. Ensure you have the necessary permissions." -ForegroundColor Red
        exit
    }
}
else {
    # Prompt the user to decide whether to check a specific AD group
    $checkGroup = Read-Host "Would you like to check a specific AD group? (Y/N)"

    if ($checkGroup.Trim().ToUpper() -eq 'Y') {
        # Prompt for the group name
        $groupInput = Read-Host "Enter the AD group name in 'DOMAIN\GroupName' format (e.g., CONTOSO\Marketing Users)"

        # Attempt to parse the input into Domain and Name
        if ($groupInput -match "^(?<Domain>[^\\]+)\\(?<Name>.+)$") {
            $Domain = $matches['Domain']
            $GroupName = $matches['Name']
        }
        else {
            Write-Host "Invalid format. Please enter in 'DOMAIN\GroupName' format." -ForegroundColor Red
            exit
        }

        # Try to get the group first
        try {
            $Group = Get-ADGroup -Identity $GroupName -Server $Domain -ErrorAction Stop
            # If it's a group, get all enabled members
            try {
                $GroupMembers = Get-ADGroupMember -Identity $Group -Recursive -ErrorAction Stop | 
                    Where-Object { $_.objectClass -eq 'user' } | 
                    Get-ADUser -Properties PasswordLastSet, PasswordNeverExpires, DistinguishedName -ErrorAction Stop |
                    Where-Object { $_.Enabled -eq $True }
                
                if ($GroupMembers.Count -eq 0) {
                    Write-Host "The group '$groupInput' has no enabled user members." -ForegroundColor Yellow
                    exit
                }

                # Assign to UserList
                $UserList = $GroupMembers | Select-Object Name, SamAccountName, DistinguishedName,
                    @{Name='TopLevelOU'; Expression={ Get-TopLevelOU $_.DistinguishedName }},
                    @{Name='PasswordLastSet'; Expression={$_.PasswordLastSet}},
                    @{Name='PasswordExpiryDate'; Expression={
                        if ($_.PasswordNeverExpires) {
                            $null  # Use $null for 'Never' to facilitate sorting
                        }
                        else {
                            $_.PasswordLastSet + $MaxPasswordAge
                        }
                    }}

                Write-Host "Retrieved users from group '$groupInput'." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to retrieve members of group '$groupInput'. Ensure you have the necessary permissions." -ForegroundColor Red
                exit
            }
        }
        catch {
            Write-Host "The group '$groupInput' does not exist in the domain '$Domain'." -ForegroundColor Red
            exit
        }
    }
    elseif ($checkGroup.Trim().ToUpper() -eq 'N') {
        # Proceed to retrieve all enabled users as in the original script
        try {
            $UserList = Get-ADUser -Filter {Enabled -eq $True} -Properties PasswordLastSet, PasswordNeverExpires, DistinguishedName -ErrorAction Stop | 
                Select-Object Name, SamAccountName, DistinguishedName,
                    @{Name='TopLevelOU'; Expression={ Get-TopLevelOU $_.DistinguishedName }},
                    @{Name='PasswordLastSet'; Expression={$_.PasswordLastSet}},
                    @{Name='PasswordExpiryDate'; Expression={
                        if ($_.PasswordNeverExpires) {
                            $null  # Use $null for 'Never' to facilitate sorting
                        }
                        else {
                            $_.PasswordLastSet + $MaxPasswordAge
                        }
                    }}
        }
        catch {
            Write-Host "Failed to retrieve user list. Check your Active Directory connectivity and permissions." -ForegroundColor Red
            exit
        }
    }
    else {
        Write-Host "Invalid input. Please enter 'Y' or 'N'." -ForegroundColor Red
        exit
    }
}

# Sort the list:
# 1. Users with PasswordNeverExpires = $true (PasswordExpiryDate = $null) appear first
# 2. Then by PasswordExpiryDate ascending (soonest to latest)
$SortedUserList = $UserList | Sort-Object `
    @{Expression = { $_.PasswordExpiryDate -eq $null }; Descending = $true}, `
    @{Expression = { $_.PasswordExpiryDate }; Ascending = $true}

# Add a DisplayPasswordExpiryDate property for user-friendly display
$SortedUserList = $SortedUserList | Select-Object *, @{
    Name = 'PasswordExpiryDateDisplay';
    Expression = {
        if ($_.PasswordExpiryDate -eq $null) {
            'Never'
        }
        else {
            $_.PasswordExpiryDate
        }
    }
} | Select-Object Name, SamAccountName, TopLevelOU, PasswordLastSet, PasswordExpiryDateDisplay

# Display the sorted list in the console
$SortedUserList | Format-Table -AutoSize

# Prompt the user to decide whether to export the list
$exportChoice = Read-Host "Would you like to export the list to a CSV file? (Y/N)"

if ($exportChoice.Trim().ToUpper() -eq 'Y') {
    # Define the export directory and ensure it exists
    $exportDirectory = "C:\Reports\"
    if (-not (Test-Path $exportDirectory)) {
        try {
            New-Item -Path $exportDirectory -ItemType Directory -Force -ErrorAction Stop
            Write-Host "Created directory $exportDirectory" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to create directory $exportDirectory. Check your permissions." -ForegroundColor Red
            exit
        }
    }
    
    # Define the export path with current date
    $date = Get-Date -Format "yyyyMMdd"
    $exportPath = "${exportDirectory}ActiveUsersPasswordExpiry_$date.csv"
    
    # Check if the file already exists
    if (Test-Path $exportPath) {
        $overwrite = Read-Host "File '$exportPath' already exists. Do you want to overwrite it? (Y/N)"
        if ($overwrite.Trim().ToUpper() -ne 'Y') {
            Write-Host "Export canceled to avoid overwriting the existing file." -ForegroundColor Red
            exit
        }
    }
    
    # Export the data to CSV, using the display property for expiry date
    try {
        $SortedUserList | Export-Csv -Path $exportPath -NoTypeInformation -ErrorAction Stop
        Write-Host "The list has been exported to $exportPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to export the list. Check your write permissions to the directory." -ForegroundColor Red
    }
}
else {
    Write-Host "Export skipped. The list is displayed above." -ForegroundColor Yellow
}