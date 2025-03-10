<#
.SYNOPSIS
    Deletes Internet Explorer cache data based on environment flag settings from RMM (NinjaOne) chosen via checkbox selection.

.DESCRIPTION
    This script reads environment flags (for example, passed from an external system)
    to determine which IE cache deletion actions to perform. For each flagged action,
    it executes the appropriate RunDll32.exe command with a specific numeric code.
    Logs are saved to user's TEMP folder.
    This script needs to be executed as user.

.PARAMETER
todo: Add parameters for custom log file path, etc.

.NOTES
Author: Hart Hoppe
#>

$logFile = Join-Path $env:TEMP "Delete-IECache_Data.log"

function Write-Log {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp - $Message"
}

# Define deletion actions based on environment flags (using string comparison to "True")
$actions = @(
    @{ Flag = ($env:everythingIncludingSavedPasswords -eq "True");       Code = 4351;  Description = "Delete Everything Including Saved Passwords" },
    @{ Flag = ($env:everythingExceptSavedPasswords -eq "True");          Code = 255;   Description = "Delete Everything Except Saved Passwords" },
    @{ Flag = ($env:temporaryInternetFiles -eq "True");                  Code = 1;     Description = "Delete Temporary Internet Files (Cache)" },
    @{ Flag = ($env:cookies -eq "True");                                 Code = 2;     Description = "Delete Cookies" },
    @{ Flag = ($env:browsingHistory -eq "True");                         Code = 8;     Description = "Delete Browsing History" },
    @{ Flag = ($env:formData -eq "True");                                Code = 16;    Description = "Delete Form Data" },
    @{ Flag = ($env:savedPasswords -eq "True");                          Code = 32;    Description = "Delete Saved Passwords" },
    @{ Flag = ($env:indexDatFiles -eq "True");                           Code = 64;    Description = "Delete Index.dat Files (Legacy)" },
    @{ Flag = ($env:autoCompleteData -eq "True");                        Code = 128;   Description = "Delete AutoComplete Data" },
    @{ Flag = ($env:feeds -eq "True");                                   Code = 256;   Description = "Delete Feeds (RSS Subscription Cache)" },
    @{ Flag = ($env:downloadHistory -eq "True");                         Code = 512;   Description = "Delete Download History" },
    @{ Flag = ($env:activeXFiltering -eq "True");                        Code = 1024;  Description = "Delete ActiveX Filtering & Tracking Protection Data" },
    @{ Flag = ($env:doNotTrackExceptions -eq "True");                    Code = 2048;  Description = "Delete Do Not Track Exceptions" },
    @{ Flag = ($env:passwordProtectedWebsitesData -eq "True");           Code = 4096;  Description = "Delete Password Protected Websites Data" },
    @{ Flag = ($env:webData -eq "True");                                 Code = 8192;  Description = "Delete Web Data (Adobe Flash Storage, Deprecated)" },
    @{ Flag = ($env:downloadHistory16384 -eq "True");                    Code = 16384; Description = "Delete Download History (Alternate)" }
)

# Check if no actions are flagged; if true, log and exit.
if (-not ($actions | Where-Object { $_.Flag })) {
    Write-Log "Nothing selected, no script run."
    return
}

# Add check: if deleteEverythingIncludingSavedPasswords is true, run only that deletion and exit.
if ($env:EverythingIncludingSavedPasswords -eq "True") {
    try {
        Write-Log "Executing: Delete Everything Except Saved Passwords."
        Start-Process -FilePath "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess", "4351" -NoNewWindow -Wait
        Write-Log "Completed: Delete Everything Except Saved Passwords."
    }
    catch {
        Write-Log "Error Executing (Delete Everything Except Saved Passwords): $($_.Exception.Message)"
    }
    return
}

# Add check: if deleteEverythingExceptSavedPasswords is true, run only that deletion and exit.
if ($env:EverythingExceptSavedPasswords -eq "True") {
    try {
        Write-Log "Executing: Delete Everything Including Saved Passwords."
        Start-Process -FilePath "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess", "4351" -NoNewWindow -Wait
        Write-Log "Completed: Delete Everything Including Saved Passwords."
    }
    catch {
        Write-Log "Error Executing (Delete Everything Including Saved Passwords): $($_.Exception.Message)"
    }
    return
}

foreach ($action in $actions) {
    if ($action.Flag) {
        try {
            Write-Log "Executing: $($action.Description)."
            Start-Process -FilePath "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess", "$($action.Code)" -NoNewWindow -Wait
            Write-Log "Completed deletion: $($action.Description)."
        }
        catch {
            Write-Log "Error Executing ($($action.Description)): $($_.Exception.Message)"
        }
    }
}