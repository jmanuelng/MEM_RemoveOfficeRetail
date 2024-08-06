<#
.SYNOPSIS
    Remediates retail versions of Office by uninstalling them.

.DESCRIPTION
    This script performs the following tasks:
    1. Verifies administrative privileges.
    2. Checks if Intune is installing any applications.
    3. Identifies installed retail versions of Office.
    4. Uninstalls detected retail versions of Office by executing the uninstall strings from the registry with "DisplayLevel=False" appended.
    5. Writes a summary of the script's execution.

.EXAMPLE
    This example runs the remediation script:
    .\Fix_OfficeRetail.ps1

.NOTES

    Functions:
    WriteAndExitWithSummary:
        Writes a summary of the script's execution to the console and exits the script with a specified status code.
        Parameters: StatusCode, Summary

    Get-OfficeInstallation:
        Retrieves a list of installed retail versions of Office.
        No parameters.

    Test-IntuneInstallation:
        Checks if Intune is currently installing any applications.
        No parameters.

    Test-AdminPrivileges:
        Verifies if the script is run with administrative privileges.
        No parameters.

    Uninstall-Office:
        Uninstalls detected retail versions of Office.
        Parameters: installedLanguages

.SUMMARY
    This script is designed to remediate retail versions of Office by uninstalling them.
    It uses registry queries to determine the presence of Office installations and uninstalls them using the uninstall strings from the registry.
    The results are summarized and output to the console, and the script exits with an appropriate status code.
#>

# Function to write a summary of execution and exit with the specified status code
function WriteAndExitWithSummary {
    param (
        [int]$StatusCode,
        [string]$Summary
    )
    
    $finalSummary = "$([datetime]::Now) = $Summary"
    $prefix = switch ($StatusCode) {
        0 { "OK" }
        1 { "FAIL" }
        default { "WARNING" }
    }
    
    Write-Host "`n`n"
    Write-Host "$prefix $finalSummary"
    Write-Host "`n`n"

    if ($StatusCode -lt 0) {$StatusCode = 0}
    Exit $StatusCode
}

# Function to check if Intune is currently installing any application
function Test-IntuneInstallation {
    <#
    .SYNOPSIS
    Checks if any application installations are currently in progress by Intune.

    .DESCRIPTION
    This function checks specific registry paths that Intune uses to track the installation status of Win32 applications. 
    It scans through all users on the device and looks for any active installation status codes that indicate an ongoing process.
    
    It primarily looks at the "EnforcementState" registry value for each application under the user's SID path. 
    If the value is 1003 (Received command to install) or 2000 (Enforcement action is in progress), it indicates an active installation.

    .EXAMPLE
    $isInstalling = Test-IntuneInstallation
    if ($isInstalling) { "Intune is currently installing an application." } else { "No active Intune installations detected." }

    .NOTES
    The function checks the "HKLM\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps" registry path for ongoing installations.
    The presence of certain status codes (e.g., 1003, 2000) indicates an active installation.

    References:
    - [Jannik Reinhard's blog on Intune Management Extension](https://jannikreinhard.com/blog/2020/11/09/intune-management-extension/)
    - [Eswar Koneti's blog on Win32 app deployment status codes in Intune](https://eskonr.com/2021/06/win32-app-deployment-status-codes-in-intune/)

    #>

    # Get the registry path for Intune Management Extension
    $baseRegistryPath = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps"
    
    # Get all user SIDs under the Intune registry path
    $userSIDs = Get-ChildItem -Path $baseRegistryPath | Select-Object -ExpandProperty PSChildName

    # Iterate through each user SID and check the installation status of each app
    foreach ($sid in $userSIDs) {
        # Get all app GUIDs under each user's SID
        $appGuids = Get-ChildItem -Path "$baseRegistryPath\$sid" | Where-Object { $_.PSChildName -ne "GRS" } | Select-Object -ExpandProperty PSChildName

        foreach ($appGuid in $appGuids) {
            # Get the enforcement state for each app
            $enforcementState = Get-ItemProperty -Path "$baseRegistryPath\$sid\$appGuid" -Name "EnforcementState" -ErrorAction SilentlyContinue

            # Check if the enforcement state indicates an active installation
            if ($enforcementState.EnforcementState -eq 1003 -or $enforcementState.EnforcementState -eq 2000) {
                return $true
            }
        }
    }

    # If no active installations are detected, return false
    return $false
}

# Function to verify if the script is run with administrative privileges
function Test-AdminPrivileges {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Script must run as admin."
    }
}

# Function to retrieve installed retail versions of Office
function Get-OfficeInstallation {
    try {
        $languages = @("cs-CZ", "da-DK", "de-DE", "en-US", "en-GB", "es-ES", "es-MX", "et-EE", "fi-FI", "fr-FR", 
                       "hr-HR", "hu-HU", "it-IT", "lt-LT", "nb-NO", "nl-NL", "pl-PL", "pt-BR", "pt-PT", "ro-RO", 
                       "sk-SK", "sl-SI", "sr-Latn-RS", "sv-SE", "tr-TR")
        $registryBasePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        $installedLanguages = @()

        foreach ($lang in $languages) {
            $registryPath = "$registryBasePath\O365HomePremRetail - $lang"
            if (Test-Path -Path $registryPath) {
                $installedLanguages += $lang
            }
        }

        return $installedLanguages
    } catch {
        throw "Error occurred while checking Office installations: $_"
    }
}

# Function to uninstall Office retail versions
function Uninstall-Office {
    <#
    .SYNOPSIS
    Uninstalls Office retail versions from the device.

    .DESCRIPTION
    Function retrieves the uninstall command for each detected Office installation from the registry and executes it to remove the application. It ensures that the uninstall process runs silently by adding "DisplayLevel=False" to the uninstall string.

    The function uses `cmd.exe` to execute the uninstall commands for several reasons:
        Compatibility: Ensures shell-specific syntax or commands are interpreted correctly.
        Error Handling and Execution Context: Provides better error handling and ensures the correct execution context.
        Batch Processing: Allows for sequential execution of commands in the correct order.

    To prevent the Command Prompt window from popping up during execution, the function uses the `-WindowStyle Hidden` parameter with `Start-Process`.

    .PARAMETER installedLanguages
    An array of languages for which Office retail versions are installed, based on registry entries.

    .EXAMPLE
    Uninstall-Office -installedLanguages @("en-US", "fr-FR")

    #>

    param (
        [string[]]$installedLanguages
    )
    
    $registryBasePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

    foreach ($lang in $installedLanguages) {
        $registryPath = "$registryBasePath\O365HomePremRetail - $lang"
        try {
            # Retrieve the uninstall string from the registry
            $uninstallString = (Get-ItemProperty -Path $registryPath -Name UninstallString).UninstallString
            
            if ($uninstallString) {
                # Add 'DisplayLevel=False' to the uninstall string for silent uninstallation
                $uninstallString += " DisplayLevel=False"
                
                Write-Host "Uninstalling Office for language $lang"
                
                # Execute the uninstall command using cmd.exe, with no window popping up
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString" -WindowStyle Hidden -Wait
                
                # Wait for 5 seconds to ensure the uninstall process completes
                Start-Sleep -Seconds 5
            }
        } catch {
            throw "Error uninstalling Office for language $($lang): $_"
        }
    }
}


# Main script execution
try {
    $summary = ""               # Summary of script execution.
    $imeInstallation = $true    # Detects if Intune Management Engine (IME) is busy executing App installation.
        
    # Verify administrative privileges
    Test-AdminPrivileges

    # Check if Intune is currently installing applications
    $imeInstallation = Test-IntuneInstallation
    if ($imeInstallation) {
        $summary = "Intune busy installing applications."
        WriteAndExitWithSummary -StatusCode 1 -Summary $summary
    }

    # Check for retail versions of Office installations
    $officeInstallations = Get-OfficeInstallation

    if ($officeInstallations.Count -eq 0) {
        $summary = "Office Retail NOT detected."
        WriteAndExitWithSummary -StatusCode 0 -Summary $summary
    } else {
        # Uninstall detected retail versions of Office
        Uninstall-Office -installedLanguages $officeInstallations
        $summary = "Office Retail uninstalled for: $($officeInstallations -join ', ')."
        WriteAndExitWithSummary -StatusCode 0 -Summary $summary
    }
} catch {
    WriteAndExitWithSummary -StatusCode 1 -Summary "Error: $_"
}
