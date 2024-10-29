# Active Directory Password Expiration Report

## Overview

This PowerShell script generates a report of active users in Active Directory, displaying their password expiration statuses along with their top-level Organizational Unit (OU). It sorts users whose passwords never expire at the top, followed by those with upcoming password expirations. Additionally, it offers the option to export the report to a CSV file for further analysis.

## Features

- **Active Users Retrieval:** Lists all enabled users in Active Directory.
- **Password Expiry Calculation:** Shows when each user's password is set to expire based on the domain's password policy.
- **Top-Level OU Display:** Indicates the primary OU each user belongs to.
- **Sorting Mechanism:** Prioritizes users with non-expiring passwords, then sorts by nearest expiration dates.
- **Export Capability:** Option to export the report to a CSV file.

## Prerequisites

- **Operating System:** Windows 10 or later, or Windows Server editions.
- **PowerShell Version:** 5.1 or higher.
- **Active Directory Module:** Ensure the **ActiveDirectory** module is installed (part of RSAT).

## Installation

1. **Install RSAT (If Not Installed):**
   - **Windows 10 and Later:**
     ```powershell
     Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
     ```
   - **Windows Server:**
     ```powershell
     Install-WindowsFeature -Name "RSAT-AD-PowerShell" -IncludeAllSubFeature -IncludeManagementTools
     ```

2. **Download the Script:**
   - Save the script to a directory, e.g., `C:\Scripts\ActiveDirectoryPasswordExpirationReport.ps1`.

3. **Set Execution Policy (If Necessary):**
   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser