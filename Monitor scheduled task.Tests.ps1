#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    Function New-TaskObjectsHC {
        Param (
            [Parameter(Mandatory, ValueFromPipeline)]
            [HashTable[]]$Hash
        )
    
        Process {
            foreach ($H in $Hash) {
                $Obj = New-Object -TypeName 'Microsoft.Management.Infrastructure.CimInstance' -ArgumentList @('MSFT_ScheduledTask')                                                                                                        
                $H.GetEnumerator() | ForEach-Object {
                    $Obj.CimInstanceProperties.Add([Microsoft.Management.Infrastructure.CimProperty]::Create($_.Key, $_.Value, [Microsoft.Management.Infrastructure.CimFlags]::None))  
                }
                $Obj
            }
        }
    }
    
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptAdmin = 'amdin@contoso.com'
    }

    Mock Get-ScheduledTask
    Mock Get-ScheduledTaskInfo
    Mock Send-MailHC
    Mock Start-ScheduledTask {
        $PSBoundParameters
    }
    Mock Write-EventLog
}
Describe 'do not start tasks' {
    It 'that are disabled' {
        Mock Get-ScheduledTask {
            @(
                @{
                    TaskName = 'Pester test task'
                    State    = 'Disabled'
                }
            ) | New-TaskObjectsHC
        }
            
        & $testScript @testParams
            
        Should -Not -Invoke Send-MailHC 
        Should -Not -Invoke Start-ScheduledTask 
    } 
    It 'that are in state running' {
        Mock Get-ScheduledTask {
            @(
                @{
                    TaskName = 'Pester test task'
                    State    = 'Running'
                }
            ) | New-TaskObjectsHC
        }
            
        & $testScript @testParams
            
        Should -Not -Invoke Send-MailHC
        Should -Not -Invoke Start-ScheduledTask
    } 
    It 'that never ran and are not in the AlwaysRunningTaskName array' {
        Mock Get-ScheduledTask {
            @(
                @{
                    TaskName = 'Pester test task'
                    State    = 'Ready'
                }
            ) | New-TaskObjectsHC
        }
        Mock Get-ScheduledTaskInfo {
            @{LastTaskResult = '267011' }
        }
            
        & $testScript @testParams
            
        Should -Not -Invoke Send-MailHC
        Should -Not -Invoke Start-ScheduledTask
    } 
}
Describe 'start tasks that are not running when' {
    It 'they are in the array AlwaysRunningTaskName' {
        Mock Get-ScheduledTask {
            @(
                @{
                    TaskName = 'Pester test task'
                    State    = 'Running'
                },
                @{
                    TaskName = 'Pester START ME task'
                    State    = 'Ready'
                }
            ) | New-TaskObjectsHC
        }
            
        . $testScript @testParams -AlwaysRunningTaskName 'Pester START ME'
            
        Should -Invoke Start-ScheduledTask -Exactly -Times 1 -ParameterFilter {
            ($InputObject.TaskName -eq 'Pester START ME task')
        }
        Should -Invoke Send-MailHC -Exactly -Times 1 -ParameterFilter {
            ($To -eq $testParams.ScriptAdmin) -and 
            ($Priority -eq 'High') -and
            ($Message -like '*Pester START ME task*') -and
            ($Subject -eq '1 task started')
        }
    } 
    It 'they are not in the array AlwaysRunningTaskName and failed their last run' {
        Mock Get-ScheduledTask {
            @(
                @{
                    TaskName = 'Pester test task'
                    State    = 'Running'
                },
                @{
                    TaskName = 'Pester START ME task'
                    State    = 'Ready'
                }
            ) | New-TaskObjectsHC
        }
        Mock Get-ScheduledTaskInfo {
            @{LastTaskResult = '1' }
        }
            
        . $testScript @testParams
            
        Should -Invoke Start-ScheduledTask -Exactly -Times 1 -ParameterFilter {
            ($InputObject.TaskName -eq 'Pester START ME task')
        }
        Should -Invoke Send-MailHC -Exactly -Times 1 -ParameterFilter {
            ($To -eq $testParams.ScriptAdmin) -and 
            ($Priority -eq 'High') -and
            ($Message -like '*Pester START ME task*') -and
            ($Subject -eq '1 task started')
        }
    } 
}
