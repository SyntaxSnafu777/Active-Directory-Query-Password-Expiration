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

# Retrieve all enabled users with necessary properties
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
