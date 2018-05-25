######################################################################
## (C) 2018 Michael Miklis (michaelmiklis.de)
##
##
## Filename:      ConvertTo-ApplicationErrorObject.ps1
##
## Version:       1.0
##
## Release:       Final
##
## Requirements:  -none-
##
## Description:   The ConvertTo-ApplicationErrorObject CMDlet parses a given Application
##                Error Event (Event-ID 1000) from the Windows Eventlog into a full
##                PowerShell object.
##                Currently only german localization is supported. To support other languages
##                modify the $EventTemplate object to match the Eventlog language
##
## This script is provided 'AS-IS'.  The author does not provide
## any guarantee or warranty, stated or implied.  Use at your own
## risk. You are free to reproduce, copy & modify the code, but
## please give the author credit.
##
####################################################################


Set-PSDebug -Strict
Set-StrictMode -Version latest
 
function ConvertTo-ApplicationErrorObject
{
    <#
        .SYNOPSIS
        Converts Event-ID 1000 messages (Application Error) into 
        a PowerShell object containing all properties

        .DESCRIPTION
        The ConvertTo-ApplicationErrorObject CMDlet parses a given Application
        Error Event (Event-ID 1000) from the Windows Eventlog into a full
        PowerShell object.
        Currently only german localization is supported. To support other languages
        modify the $EventTemplate object to match the Eventlog language
  
        .PARAMETER ApplicationErrorEvent
        A EventLogRecord containing the Event-ID 1000 Error Message
  
        .EXAMPLE
        Get-WinEvent -LogName Application | Where-Object { ($_.ID -eq 1000) -and ($_.ProviderName -eq "Application Error") } | ConvertTo-ApplicationErrorObject | ft *
 
    #>

    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()][System.Diagnostics.Eventing.Reader.EventLogRecord]$ApplicationErrorEvent
    )


    # executes once before first item in pipeline is processed
    Begin 
    {
        # intialize empty MessageObject
        $MessageObject = [PSCustomObject]@{
            MachineName = $null
            UserId = $null
            TimeCreated = $null
            ProcessName = $null
            ProcessVersion = $null
            ProcessTimestamp = $null
            ModuleName = $null
            ModuleVersion = $null
            ModuleTimestamp = $null
            Exception = $null
            FailureOffset = $null
            ProcessID = $null
            ProcessStartTime = $null
            ProcessPath = $null
            ModulePath = $null
            ReportID = $null
            FullPackageName = $null
            AppID = $null
        }

        # Template object containing localized strings - german language
        $EventTemplate_DE = [PSCustomObject]@{
            ProcessName = "Name der fehlerhaften Anwendung: "
            ProcessVersion = ", Version: " 
            ProcessTimestamp = ", Zeitstempel: "
            ModuleName = "Name des fehlerhaften Moduls: "
            ModuleVersion = ", Version: "
            ModuleTimestamp = ", Zeitstempel: "
            Exception = "Ausnahmecode: "
            FailureOffset = "Fehleroffset: "
            ProcessID = "ID des fehlerhaften Prozesses: "
            ProcessStartTime = "Startzeit der fehlerhaften Anwendung: "
            ProcessPath = "Pfad der fehlerhaften Anwendung: "
            ModulePath = "Pfad des fehlerhaften Moduls: "
            ReportID = "Berichtskennung: "
            FullPackageName = "Vollständiger Name des fehlerhaften Pakets: "
            AppID = "Anwendungs-ID, die relativ zum fehlerhaften Paket ist: "
        }

        # Template object containing localized strings - english language
        $EventTemplate_EN = [PSCustomObject]@{
            ProcessName = "Faulting application name: "
            ProcessVersion = ", version: " 
            ProcessTimestamp = ", time stamp: "
            ModuleName = "Faulting module name: "
            ModuleVersion = ", version: "
            ModuleTimestamp = ", time stamp: "
            Exception = "Exception code: "
            FailureOffset = "Fault offset: "
            ProcessID = "Faulting process id: "
            ProcessStartTime = "Faulting application start time: "
            ProcessPath = "Faulting application path: "
            ModulePath = "Faulting module path: "
            ReportID = "Report Id: "
            FullPackageName = "Faulting package full name: "
            AppID = "Faulting package-relative application ID: "
        }
    }

    # executes once for each pipeline object
    Process 
    {
   
        # try to detect language of eventlog entry
        if ($ApplicationErrorEvent.Message.Contains($EventTemplate_EN.ProcessName))
        {
            $EventTemplate = $EventTemplate_EN
        }
        elseif ($ApplicationErrorEvent.Message.Contains($EventTemplate_DE.ProcessName))
        {
            $EventTemplate = $EventTemplate_DE
        }

        # add properties from event
        $MessageObject.MachineName = $ApplicationErrorEvent.MachineName
        $MessageObject.UserId = $ApplicationErrorEvent.UserId
        $MessageObject.TimeCreated = $ApplicationErrorEvent.TimeCreated

        # split Message into line array
        $Message = ($ApplicationErrorEvent.Message -split "`r`n")

        # 1st line - extract ProcessName
        $MessageObject.ProcessName = $Message[0].Replace($EventTemplate.ProcessName, "").split(",")[0].Trim()

        # 1st line - extract ProcessVersion
        $MessageObject.ProcessVersion = $Message[0].Replace($($EventTemplate.ProcessName + $MessageObject.ProcessName + $EventTemplate.ProcessVersion), "").split(",")[0].Trim()

        # 1st line - extract ProcessTimestamp
        $MessageObject.ProcessTimestamp = $Message[0].Replace($($EventTemplate.ProcessName + $MessageObject.ProcessName + $EventTemplate.ProcessVersion + $MessageObject.ProcessVersion + $EventTemplate.ProcessTimestamp), "").split(",")[0].Trim()

        # 2nd line - extract ModuleName
        $MessageObject.ModuleName = $Message[1].Replace($EventTemplate.ModuleName, "").split(",")[0].Trim()

        # 2nd line - extract ModuleVersion
        $MessageObject.ModuleVersion = $Message[1].Replace($($EventTemplate.ModuleName + $MessageObject.ModuleName + $EventTemplate.ModuleVersion), "").split(",")[0].Trim()

        # 2nd line - extract ModuleTimestamp
        $MessageObject.ModuleTimestamp = $Message[1].Replace($($EventTemplate.ModuleName + $MessageObject.ModuleName + $EventTemplate.ModuleVersion + $MessageObject.ModuleVersion + $EventTemplate.ModuleTimestamp), "").split(",")[0].Trim()

        # 3nd line - extract Exception
        $MessageObject.Exception = $Message[2].Replace($EventTemplate.Exception, "").Trim()

        # 4th line - extract FailureOffset
        $MessageObject.FailureOffset = $Message[3].Replace($EventTemplate.FailureOffset, "").Trim()

        # 5th line - extract ProcessID
        $MessageObject.ProcessID = $Message[4].Replace($EventTemplate.ProcessID, "").Trim()

        # 6th line - extract ProcessStartTime
        $MessageObject.ProcessStartTime = $Message[5].Replace($EventTemplate.ProcessStartTime, "").Trim()

        # 7th line - extract ProcessPath
        $MessageObject.ProcessPath = $Message[6].Replace($EventTemplate.ProcessPath, "").Trim()

        # 8th line - extract ModulePath
        $MessageObject.ModulePath = $Message[7].Replace($EventTemplate.ModulePath, "").Trim()

        # 9th line - extract ReportID
        $MessageObject.ReportID = $Message[8].Replace($EventTemplate.ReportID, "").Trim()

        # 10th line - extract FullPackageName
        $MessageObject.FullPackageName = $Message[9].Replace($EventTemplate.FullPackageName, "").Trim()

        # 11th line - extract AppID
        $MessageObject.AppID = $Message[10].Replace($EventTemplate.AppID, "").Trim()

        Write-Output $MessageObject
    }
}

Get-WinEvent -LogName Application | Where-Object { ($_.ID -eq 1000) -and ($_.ProviderName -eq "Application Error") } | Select -first 1 | ConvertTo-ApplicationErrorObject | fl *