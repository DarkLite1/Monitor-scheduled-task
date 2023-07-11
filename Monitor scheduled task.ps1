<# 
    .SYNOPSIS   
        Restart scheduled tasks that failed to run.

    .DESCRIPTION
        All scheduled tasks that failed to run will be restarted when running 
        this script. Scheduled tasks that are Disabled, Running or have never 
        ran before will be ignored.

        When a task has been restarted an e-mail will be sent to the script 
        admin.

    .PARAMETER AlwaysRunningTaskName
        Array of task names that that always need to be in state 'Running'.

    .NOTES
        0 - The operation completed successfully.
        1 - Incorrect function called or unknown function called. 
        2 - File not  found.
        10 - The environment is incorrect. 
        267008 - Task is ready to run at its next scheduled time. 
        267009 - Task is currently running. 
        267010 - The task will not run at the scheduled times because it has 
                 been disabled. 
        267011 - Task has not yet run. 
        267012 - There are no more runs scheduled for this task. 
        267013 - One or more of the properties that are needed to run this task 
                 on a schedule have not been set. 
        267014 - The last run of the task was terminated by the user. 
        267015 - Either the task has no triggers or the existing triggers are  
                 disabled or not set. 
        2147750671 - Credentials became corrupted. 
        2147750687 - An instance of this task is already running. 
        2147943645 - The service is not available (is "Run only when an user is 
                      logged on" checked?). 
        3221225786 - The application terminated as a result of a CTRL+C. 
        3228369022 - Unknown software exception.
#>

[CmdLetBinding()]
Param (
    [String]$ScriptName = 'Monitor scheduled task (ALL)',
    [String]$TaskPath = '\HCScripts',
    [String[]]$AlwaysRunningTaskName = ('Monitor mailbox', 'Monitor folder'),
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        $startTask = {
            Param (
                [Parameter(Mandatory)]
                $Task
            )
            Write-Warning "Start task '$($Task.TaskName)'"
            Write-EventLog @EventWarnParams -Message "Task '$($Task.TaskName)' not running"
            Start-ScheduledTask -InputObject $Task
            Write-EventLog @EventVerboseParams -Message "Task '$($Task.TaskName)' started"
        }

        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $tasks = Get-ScheduledTask -TaskPath "$TaskPath\*" | 
        Where-Object { ($_.State -ne 'Disabled') }

        $alwaysRunningTasksNotRunning = @()
        $failedTasks = @()

        foreach ($task in $tasks) {
            if ($task.State -eq 'Running') {
                Write-Verbose "Task is running '$($task.TaskName)'"
                Continue
            }

            if ($AlwaysRunningTaskName.where( 
                    { $task.TaskName -like "$($_)*" }
                )
            ) {
                Write-Verbose "Task should be running but is not '$($task.TaskName)'"
                $alwaysRunningTasksNotRunning += $task.TaskName
                & $startTask -Task $task
                Continue
            }

            $taskInfo = Get-ScheduledTaskInfo $task

            if ($taskInfo.LastTaskResult -eq '0') {
                Write-Verbose "Task ran successful '$($task.TaskName)' "
                Continue
            }

            if ($taskInfo.LastTaskResult -eq '267011') {
                Write-Verbose "Task never ran '$($task.TaskName)'"
                Continue
            }

            $failedTasks += $task.TaskName
            & $startTask -Task $task
        }
        
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}
End {
    Try {
        $mailParams = @{
            To       = $ScriptAdmin
            Priority = 'High'
            Header   = $ScriptName
        }

        if ($alwaysRunningTasksNotRunning) {
            $mailParams.Message += "Started the following tasks because they were no longer in state 'Running':$($alwaysRunningTasksNotRunning  | ConvertTo-HtmlListHC)"
        }
        
        if ($failedTasks) {
            $mailParams.Message += "<p>Started the following tasks because they failed their last run:$($failedTasks  | ConvertTo-HtmlListHC)</p>"
        }
        
        if ($mailParams.Message) {
            $StartedTasks = @($failedTasks + $alwaysRunningTasksNotRunning).count

            $mailParams.Subject = "$StartedTasks tasks started"
            if ($StartedTasks -eq 1) { $mailParams.Subject = "1 task started" }

            Send-MailHC @mailParams
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "- $_"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}