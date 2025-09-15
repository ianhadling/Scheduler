<#
Name:			Scheduler_v1.ps1
Version:		1.1
Description:	Trigger scheduled tasks from the database between the start and end times at given intervals
Request Source:	Powershell scheduler project
Author:			Ian Hadlington
Created on:		22-Mar-2025
Called From:	Windows Task Scheduler
Parameters:		@StartTime, @EndTime, @IntervalMinutes


Change Control
==============

When		Who		Version		What
====		===		=======		====

#>


param(
    [string]$StartTime  = "09:30" # Default to 12 noon"
    ,[string]$EndTime = "23:00" # Default to 23.5 hours
    ,[int32]$IntervalMinutes = 15 # Default to 15 minutes
)


# Constants
$PowerShellExe = "C:\Windows\System32\WindowsPowershell\v1.0\Powershell.exe"
$Server = "itsazsql050.itservices.local"
$RepDatabase = "EngageReporting"
$GetTasksSP = "EXEC dbo.uspTaskResultsGetForSchedulerByDT @DateFor = '<rundate>'"
$CleanupSP = "EXEC dbo.uspTaskResultsCleanup @ScheduleRunTime = '<rundate>'"
$UpdateTasksResultsSP = "EXEC dbo.uspTaskResultsUpdate @TaskResultId = '<TaskResultid>', @TaskStatus = '<status>', @Message = '<message>'"
$LogMessageSP = "EXEC dbo.uspGeneralLogAdd @LogSource = '<LogSource>', @LogStatus = '<LogStatus>', @LogMessage = '<LogMessage>', @AddnlInfo = '<AddnlInfo>'"
$LogSource = "Scheduler - Live"

# Import the SQL Server module
# Import-Module SqlServer -Force
 Import-Module "\\itservices.local\mgl$\Darwen\System Development\Powershell\CreateSpreadsheetFromSP.psm1" -Force


# Function to Run Excel VBA
function Run-ExcelMacro {
    param (
        [object]$ExcelApp,
        [string]$ExcelFilePath,
        [string]$MacroName
    )
    $ExcelMacroSuccess = $false
    # Check if the Excel application is already running
    if (-not $ExcelApp) {
        Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Creating new Excel application instance." -WriteHostOutput $true
        $ExcelApp = New-Object -ComObject Excel.Application
    } 
    # Check if the file exists
    if (-not (Test-Path -Path $ExcelFilePath)) {
        Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Excel file not found: $ExcelFilePath" -WriteHostOutput $true
        return $false
    }
    try {
        # Make Excel visible (optional)
        $ExcelApp.Visible = $true

        # Open the macro-enabled workbook
        $Workbook = $ExcelApp.Workbooks.Open($ExcelFilePath, [ref]$null, $true)

        # Run the macro
        if ($Task.SendTaskID){
            $TaskResultID = [ref]$Task.TaskResultID
            $ExcelMacroSuccess = $ExcelApp.Application.Run("'$($ExcelFilePath)'!$($MacroName)", "Auto", $TaskResultID)
        } else{
            $ExcelMacroSuccess = $ExcelApp.Application.Run("'$($ExcelFilePath)'!$($MacroName)", "Auto")
        }

    } catch {
        Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "An error occurred while running Excel macro: $($_)" -WriteHostOutput $true
        $ExcelMacroSuccess = $false
    } finally {
        # Clean up
        $Workbook.Close($false)
    }
    Return $ExcelMacroSuccess
}

# Function to log information
function Log-Information {
    param (
        [string]$LogSource,
        [string]$LogStatus = 'Info',
        [string]$LogMessage,
        [string]$AddnlInfo = '',
        [boolean]$WriteHostOutput = $true
    )
    $Query = $LogMessageSP -replace '<LogSource>', $LogSource -replace '<LogStatus>', $LogStatus -replace '<LogMessage>', $LogMessage -replace '<AddnlInfo>', $AddnlInfo

    Try{
        $null = Invoke-Sqlcmd -ServerInstance $Server -Database $RepDatabase -Query $Query -ErrorAction Stop -TrustServerCertificate
    } catch {
        Write-Host "Error logging information: $($_)"
    }
    # If the logging fails, we can still log to the console
    if ($WriteHostOutput) {
        Write-Host "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) Log Entry: $($LogStatus), $($LogSource), $($LogMessage), $($AddnlInfo)"
    }
}


function Update-TaskResults {
    param (
        [int32]$TaskResultId,
        [string]$Status,
        [string]$Message = 'NULL'
    )

    $Query = $UpdateTasksResultsSP -replace '<TaskResultid>', $TaskResultId -replace '<status>', $Status -replace '<message>', $Message

    Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Update Task Results for Task Result ID $($TaskResultId) with Status $($Status)" -WriteHostOutput $true
    Try {
        $null = Invoke-Sqlcmd -ServerInstance $Server -Database $RepDatabase -Query $Query -TrustServerCertificate -ErrorAction Stop
    } catch {
        Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Error updating task results: $($_)" -WriteHostOutput $true
    }
}

function Cleanup-TaskResults {
    param (
        [string]$Database,
        [datetime]$ScheduleRunTime
    )

    $Query = $CleanupSP -replace '<rundate>', $ScheduleRunTime.ToString("yyyy-MM-dd HH:mm")

    Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Cleanup Task Results for $($Database) at $($ScheduleRunTime.ToString("yyyy-MM-dd HH:mm"))" -WriteHostOutput $true
    Try {
        # Execute the cleanup stored procedure
        $null = Invoke-Sqlcmd -ServerInstance $Server -Database $Database -Query $Query -TrustServerCertificate -ErrorAction Stop
    } catch {
        Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Error cleaning up task results: $($_)" -WriteHostOutput $true
    }
}

# Main function
function Run-AccessTasks {
    param (
        [object]$TasksToRun
        ,[datetime]$Timeslot
    )
    try {
        # Create a new instance of the Access application
        $AccessApp = New-Object -ComObject Access.Application
        $AccessApp.OpenCurrentDatabase("$($TasksToRun[0].SourceFolder)\$($TasksToRun[0].SourceFileName)")  

        $AccessTaskTimeout = 600 # Timeout in seconds (10 minutes)

        foreach ($Task in $TasksToRun) {
            # Check if the current date is still the date of the timeslot
            if ($Timeslot.Date -ne (Get-Date).Date) {
                Log-Information -LogSource $LogSource -LogStatus "Warning" -LogMessage "The timeslot has run into another day. Aborting schedule: $($Task.TaskName)" -WriteHostOutput $true
                break
            }
            Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Processing $($Task.SourceFileName.Replace('.accdb', '')) Task: $($Task.SourceFunction)" -WriteHostOutput $true
            $TaskStatus = 'Start'
            Update-TaskResults -TaskResultId $Task.TaskResultID -Status $TaskStatus
            # Run the Access function

            $TaskSuccess = $false
            try {
                if ($Task.SendTaskID){
                    $TaskResultID = [ref]$Task.TaskResultID
                    $TaskSuccess = $AccessApp.Run($Task.SourceFunction, $TaskResultID)
                } else{
                    $TaskSuccess = $AccessApp.Run($Task.SourceFunction )
                }
            } catch {
                    Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Access function $($Task.SourceFunction) failed. Unhandled Exception: $($_.Exception.Message)" -WriteHostOutput $true
                    $TaskSuccess = $false
            }

                if ($TaskSuccess) {
                    $TaskStatus = "Success"
                } else {
                    $TaskStatus = "Fail"
                    Log-Information -LogSource $LogSource -LogStatus "Warning" -LogMessage "Access function $($Task.SourceFunction) returned a failure status." -WriteHostOutput $true
            }
            Update-TaskResults -TaskResultId $Task.TaskResultID -Status $TaskStatus
        }

        if ($AccessApp) {
            $AccessApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($AccessApp) | Out-Null
        }

    } catch {
        Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Error running Access tasks(3): $($_)" -WriteHostOutput $true
    }
}

function Run-ExcelTasks {
    param (
        [object]$TasksToRun
        ,[datetime]$Timeslot
    )
    try {
        $ExcelApp = New-Object -ComObject Excel.Application

        foreach ($Task in $TasksToRun) {
            # Check if the current date is still the date of the timeslot
            if ($Timeslot.Date -ne (Get-Date).Date) {
                Log-Information -LogSource $LogSource -LogStatus "Warning" -LogMessage "The timeslot has run into another day. Aborting schedule: $($Task.TaskName)" -WriteHostOutput $true
                break
            }

            Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Processing Excel Task: $($Task.SourceFileName)" -WriteHostOutput $true
            $TaskStatus = 'Start'
            Update-TaskResults -TaskResultId $Task.TaskResultID -Status $TaskStatus
            # Run the Excel macro
            $Result = Run-ExcelMacro -ExcelApp $ExcelApp -ExcelFilePath "$($Task.SourceFolder)\$($Task.SourceFileName)" -MacroName "RunReport"
            if ($Result) {
                $TaskStatus = "Success"
            } else {
                $TaskStatus = "Fail"
            }
            Update-TaskResults -TaskResultId $Task.TaskResultID -Status $TaskStatus
        }
        $ExcelApp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp) | Out-Null
        Remove-Variable ExcelApp
    } catch {
        Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Error encountered running Excel tasks: $($_)" -WriteHostOutput $true
    }
}

function Create-Spreadsheet {
    param (
        [object]$Task
    )
    $smtpServer = "10.225.20.8"
    $ForColour = "Black"
    $BkColour = "Yellow"

    $Config = [xml]$Task.Config

    $ccAddressList = $Task.EmailccList -split ';'
    $MailToAddressList = $Task.EmailToList -split ';'
    If ($Config.Values.EmailFrom -and  $Config.Values.EmailFrom.Contains("@")) {
        $EmailFrom = $Config.Values.EmailFrom
    } else {
        # Default email address if not specified in the task
        $EmailFrom = "en.reporting@engage-services.co.uk"
    }

    $CallParams = @{
        Server           = $Server    
        Database         = $Config.Values.Database
        StoredProc       = $Config.Values.StoredProc
        WkSheet          = $Config.Values.WkSheet
        ReportName       = $Task.TaskName
        ForColour        = $ForColour
        BkColour         = $BkColour
        SaveFolder       = $Task.SavePath
        EmailSendTo      = $MailToAddressList
        EmailCCTo        = $ccAddressList
        EmailBody        = $Task.EmailBody
        EmailSmtpServer  = $smtpServer
        EmailSmtpPort    = 25
        EmailFrom        = $EmailFrom
    }

    try{
        CreateSpreadsheetFromSP @CallParams
        return 0
    } catch {
        Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Error creating spreadsheet: $($_)" -WriteHostOutput $true
        Return -1
    }
}

function Run-PowerShellTasks {
    param (
        [object]$TasksToRun
        ,[datetime]$Timeslot
    )
    try {
        foreach ($Task in $TasksToRun) {
            # Check if the current date is still the date of the timeslot
            if ($Timeslot.Date -ne (Get-Date).Date) {
                Log-Information -LogSource $LogSource -LogStatus "Warning" -LogMessage "The timeslot has run into another day. Aborting schedule: $($Task.TaskName)" -WriteHostOutput $true
                break
            }
            Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Processing PowerShell Task: $($Task.TaskName)" -WriteHostOutput $true
            $TaskStatus = 'Start'
            Update-TaskResults -TaskResultId $Task.TaskResultID -Status $TaskStatus

            If ($Task.SourceFunction -eq "CreateSpreadsheet") {
                # If the task is to create a spreadsheet, we need to pass the task object to the script
                # Create the command to run the PowerShell script with the task object
                # Note: The CreateSpreadsheet function should be defined in the same script or imported from a module
                $PS_Output = Create-Spreadsheet -Task $Task
            } else {
#                $PSCmd = "$PowerShellExe -ExecutionPolicy Bypass -File `'$($Task.SourceFolder)\$($Task.SourceFileName)`'"
#                $PSCmd = "'$($Task.SourceFolder)\$($Task.SourceFileName)' -ExecutionPolicy Bypass -NoProfile"
                try{
                    & "$($Task.SourceFolder)\$($Task.SourceFileName)" -ExecutionPolicy Bypass -NoProfile -TaskResultId $Task.TaskResultID -ErrorAction Stop 
                    $PS_Output = 0
                } catch {
                    Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Error encountered running PowerShell task: $($_)" -WriteHostOutput $true
                    $PS_Output = 1
                }
#                finally {$PS_Output = $LASTEXITCODE}
            }

            if ($PS_Output -eq 0) {
                $TaskStatus = "Success"
            } else {
                $TaskStatus = "Fail"
            }
            Update-TaskResults -TaskResultId $Task.TaskResultID -Status $TaskStatus
        }
    } catch {
        Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Error encountered running Powershell tasks: $($_)" -WriteHostOutput $true
    }
}


# Calculate actual start time and end time for the scheduler
$StartDateTime = [datetime]::ParseExact("$((Get-Date).ToString("yyyy-MM-dd")) $($StartTime)", "yyyy-MM-dd HH:mm", $null)
$EndDateTime = [datetime]::ParseExact("$((Get-Date).ToString("yyyy-MM-dd")) $($EndTime)", "yyyy-MM-dd HH:mm", $null)

# $EndDateTime = $StartDateTime.AddMinutes($DurationMinutes - $IntervalMinutes) # Miss out the last interval in case a task runs over the end time
# Log the start and end times
Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Scheduler Start Time: $($StartDateTime.ToString("yyyy-MM-dd HH:mm"))" -WriteHostOutput $true
Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Scheduler End Time: $($EndDateTime.ToString("yyyy-MM-dd HH:mm"))" -WriteHostOutput $true

if ($StartDateTime -gt $EndDateTime) {
    Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Start time $($StartDateTime.ToString("yyyy-MM-dd HH:mm")) is after end time $($EndDateTime.ToString("yyyy-MM-dd HH:mm")). Exiting script." -WriteHostOutput $true
    Exit 1
}
# Pause until the start time if the current time is before the start time
if ((Get-Date) -lt $StartDateTime) {
    $TimeToNextStart = [Int32](New-TimeSpan -Start (Get-Date) -End $StartDateTime).TotalSeconds
    Start-Sleep -Seconds $TimeToNextStart
} else {
    # Stop the script displaying that the start time is in the past
    Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "The start time $($StartDateTime.ToString("yyyy-MM-dd HH:mm")) is in the past. Exiting script." -WriteHostOutput $true
    Exit 1
}  

Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Scheduler started at $((Get-Date).ToString("yyyy-MM-dd HH:mm"))" -WriteHostOutput $true
# Main loop to run tasks every [$IntervalMinutes] minutes

$SchedulerRunTime = Get-Date

# Run the scheduler until but not beyond the end time
while ( (Get-Date) -lt $EndDateTime.AddMinutes(-$IntervalMinutes) ) {
    $WithinTimeSlot =((Get-Date).AddMinutes(-1) -lt $SchedulerRunTime )    
    
#    $SchedulerRunTime = Get-Date

    If ($WithinTimeSlot) {
        $RunSQL = $GetTasksSP -replace '<rundate>', $SchedulerRunTime.ToString("yyyy-MM-dd HH:mm")
        Try{
            $TasksToRun = @(Invoke-Sqlcmd -ServerInstance $Server -Database $RepDatabase -Query $RunSQL -TrustServerCertificate -ErrorAction Stop)
        } Catch {
            Log-Information -LogSource $LogSource -LogStatus "Error" -LogMessage "Error retrieving tasks: $($_)" -WriteHostOutput $true
            Break # Exit the loop if there is an error retrieving tasks
        }
        Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "$($TasksToRun.Count) Tasks to run at $($SchedulerRunTime.ToString("yyyy-MM-dd HH:mm"))" -WriteHostOutput $true
    } 
    
    If(-not $WithinTimeSlot){
        Log-Information -LogSource $LogSource -LogStatus "Warning" -LogMessage "timeslot overrun for scheduler at $($SchedulerRunTime.ToString("yyyy-MM-dd HH:mm"))" -WriteHostOutput $true
    #} elseif ($TasksToRun.Count -eq 0) {
    #    Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "No tasks to run for scheduler at $($SchedulerRunTime.ToString("yyyy-MM-dd HH:mm"))" -WriteHostOutput $true
    } else {

        #Get and run all Exodus tasks
        $ExodusWarrants = @($TasksToRun | Where-Object { $_.SourceFileName -eq "Exodus Warrants.accdb" } )
        If(($ExodusWarrants).Count -gt 0) {
            Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Running Exodus Warrants tasks. No Of tasks is: $($ExodusWarrants.Count)" -WriteHostOutput $true
            Run-AccessTasks -TasksToRun $ExodusWarrants -TimeSlot $SchedulerRunTime
        }

        #Get and run all Genesys tasks
        $GenesysTasks = @($TasksToRun | Where-Object { $_.SourceFileName -eq "Genesys.accdb" } )
        If(($GenesysTasks).Count -gt 0) {
            Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Running Genesys tasks. No Of tasks is: $($GenesysTasks.Count)" -WriteHostOutput $true
            Run-AccessTasks -TasksToRun $GenesysTasks -TimeSlot $SchedulerRunTime
        }

        #Get and run all Excel tasks
        $ExcelTasks = @($TasksToRun | Where-Object { $_.SourceFileName -like "*.xlsb" })
        If(($ExcelTasks).Count -gt 0) {
            Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Running Excel tasks. No Of tasks is: $($ExcelTasks.Count)" -WriteHostOutput $true
            Run-ExcelTasks -TasksToRun $ExcelTasks -TimeSlot $SchedulerRunTime
        }
        
        #Get and run all PowerShell tasks
        $PowerShellTasks = @($TasksToRun | Where-Object { $_.SourceFileName -like "*.ps1" })
        If(($PowerShellTasks).Count -gt 0) {
            Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Running PowerShell tasks. No Of tasks is: $($PowerShellTasks.Count)" -WriteHostOutput $true
            Run-PowerShellTasks -TasksToRun $PowerShellTasks -TimeSlot $SchedulerRunTime
        }
    }

    # Give any task that has a start date/time an end date/time and set the completion message accordingly
    Cleanup-TaskResults -Database $RepDatabase -ScheduleRunTime $SchedulerRunTime

    # Get the difference in minutes between the current time and the next scheduled run time

    $SchedulerRunTime = $SchedulerRunTime.AddMinutes($IntervalMinutes)
    if ($SchedulerRunTime -ge $EndDateTime) {
        Break # Exit the loop if the next start time is beyond the end time
    }

    $TimeToNextStart = [Int32](New-TimeSpan -Start (Get-Date) -End $SchedulerRunTime).TotalSeconds

    if ($TimeToNextStart -gt 0) {
        Start-Sleep -Seconds $TimeToNextStart # Sleep until [$IntervalMinutes] minutes after the last run
    }
} 

Log-Information -LogSource $LogSource -LogStatus "Info" -LogMessage "Scheduler Ended at $((Get-Date).ToString("yyyy-MM-dd HH:mm"))" -WriteHostOutput $true
