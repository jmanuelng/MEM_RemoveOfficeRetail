<#
.SYNOPSIS
    Checks if the OS installation date is within a specified number of hours and identifies retail versions of Office installed on the device.

.DESCRIPTION
    Script performs two main tasks:
    1. It checks if the OS installation date is within a specified threshold of hours from the current time.
    2. It searches the registry to find if any retail versions of Office are installed, specifically looking for multiple language versions.
    The script returns:
    - 0 if no retail versions of Office are found.
    - 1 if retail versions of Office are found and the OS installation date is within the threshold.
    - -2 if retail versions of Office are found but the OS installation date is not within the threshold.
    - -3 if an error occurs during the script execution.
    It also writes a summary of the script's execution.

.PARAMETER HoursThreshold
    The number of hours to check against the OS installation date. Default is 24 hours.

.EXAMPLE
    This example runs the script with a default threshold of 24 hours:
    .\Script.ps1

.NOTES
    Functions:
    1. WriteAndExitWithSummary:
        Writes a summary of the script's execution to the console and exits the script with a specified status code.
        Parameters: StatusCode, Summary

    2. Get-OSInstallationDate:
        Checks if the OS installation date is within a specified number of hours.
        Parameters: HoursThreshold

    3. Get-OfficeInstallation:
        Retrieves a list of installed retail versions of Office.
        No parameters.

.SUMMARY
    This script is designed to check the OS installation date and identify installed retail versions of Office.
    It uses registry queries to determine the installation date and the presence of Office installations.
    The results are summarized and output to the console, and the script exits with an appropriate status code.

#>

# Function to check OS installation date
function Get-OSInstallationDate {
    <#
    .SYNOPSIS
        Checks if the OS installation date is within a specified number of hours.

    .DESCRIPTION
        This function retrieves the OS installation date from the registry and checks if it is within a specified threshold of hours from the current time.
        The default threshold is set to 24 hours. It converts the Unix epoch time (seconds since 1970-01-01) to a readable date format and calculates the time difference.
        Returns a hashtable containing the result and the installation date.

    .PARAMETER HoursThreshold
        The number of hours to check against the OS installation date. Default is 24 hours.

    .EXAMPLE
        $result = Get-OSInstallationDate -HoursThreshold 48
        This example checks if the OS was installed within the last 48 hours and returns a hashtable with the result and the installation date.

    #>
    param (
        [int]$HoursThreshold = 24  # Parameter to specify the number of hours threshold
    )

    try {
        # Retrieve the InstallDate property from the registry
        $osInstallDate = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'InstallDate'
        
        # Check if the InstallDate property exists and is not null
        if ($osInstallDate -and $osInstallDate.InstallDate) {
            # Convert the InstallDate from Unix epoch time to a readable DateTime format
            $installDateTime = (Get-Date "1970-01-01 00:00:00").AddSeconds($osInstallDate.InstallDate)
            
            # Get the current date and time
            $currentTime = Get-Date
            
            # Calculate the time difference in hours between the current time and the installation date
            $timeDifference = ($currentTime - $installDateTime).TotalHours

            # Check if the time difference is less than or equal to the specified threshold
            if ($timeDifference -le $HoursThreshold) {
                return @{ IsWithinThreshold = $true; InstallationDate = $installDateTime }  # Return true and the installation date if within the threshold
            } else {
                return @{ IsWithinThreshold = $false; InstallationDate = $installDateTime }  # Return false and the installation date if outside the threshold
            }
        } else {
            throw "The InstallDate property was not found."  # Throw an error if the InstallDate property is missing
        }
    } catch {
        throw "An error occurred while retrieving the OS installation date: $_"  # Throw an error if an exception occurs
    }
}

# Function to check if retail versions of Office are installed
function Get-OfficeInstallation {
    <#
    .SYNOPSIS
        Retrieves a list of installed retail versions of Office.

    .DESCRIPTION
        Function checks registry to identify any retail versions of Office on the device. 
        Loops through a predefined list of language codes, constructs the corresponding registry path, and verifies its existence.
        If the registry key for a language is found, it adds the language code to a list of detected installations.
        Function returns a list of languages for which retail versions of Office are installed.

    .EXAMPLE
        $installedLanguages = Get-OfficeInstallation
        This example retrieves a list of languages for which retail versions of Office are installed on the device.
    
    #>
    try {
        # List of language codes to check for Office installations
        $languages = @("cs-CZ", "da-DK", "de-DE", "en-US", "en-GB", "es-ES", "es-MX", "et-EE", "fi-FI", "fr-FR", 
                       "hr-HR", "hu-HU", "it-IT", "lt-LT", "nb-NO", "nl-NL", "pl-PL", "pt-BR", "pt-PT", "ro-RO", 
                       "sk-SK", "sl-SI", "sr-Latn-RS", "sv-SE", "tr-TR")
        
        # Base registry path for Office installations
        $registryBasePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        
        # List to store detected languages with installed Office versions
        $installedLanguages = @()

        # Loop through each language code
        foreach ($lang in $languages) {
            # Construct the registry path for the current language
            $registryPath = "$registryBasePath\O365HomePremRetail - $lang"
            
            # Check if the registry path exists
            if (Test-Path -Path $registryPath) {
                # Add the language code to the list if the registry path exists
                $installedLanguages += $lang
            }
        }

        # Return the list of detected languages
        return $installedLanguages
    } catch {
        # Throw an error if an exception occurs
        throw "An error occurred while checking Office installations: $_"
    }
}



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

# Main script execution
try {
    $summary = ""                                                               # Summary of script execution
    $osInstallResult = Get-OSInstallationDate -HoursThreshold 24                # Check OS installation date and capture the result
    $osInstallationWithinThreshold = $osInstallResult.IsWithinThreshold             # True or False if it is within treshold
    $installationDate = $osInstallResult.InstallationDate                           # The installation date of OS
    $officeInstallations = Get-OfficeInstallation                               # Check for retail versions of Office installations

    # Determine the summary and status code based on the results
    if ($officeInstallations.Count -eq 0) {
        $summary = "Office Retail NOT detected. OS installation date: $installationDate"
        WriteAndExitWithSummary -StatusCode 0 -Summary $summary
    } elseif ($officeInstallations.Count -gt 0 -and $osInstallationWithinThreshold) {
        $summary = "Office Retail detected: $($officeInstallations -join ', '). OS installation within threshold, date: $installationDate."
        WriteAndExitWithSummary -StatusCode 1 -Summary $summary
    } else {
        $summary = "Office Retail detected: $($officeInstallations -join ', '). OS installation date: $installationDate. OS installation NOT within threshold."
        WriteAndExitWithSummary -StatusCode -2 -Summary $summary
    }
} catch {
    # Handle any errors that occur during the script execution
    WriteAndExitWithSummary -StatusCode -3 -Summary "Error: $_"
}
