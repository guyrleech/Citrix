<#
    Common functions used by multiple scripts

    Guy Leech, 2018
#>


Function Get-RemoteInfo
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$computer , 
        [int]$jobTimeout = 60 , 
        [Parameter(Mandatory=$true)]
        [scriptblock]$work
    )
    $results = $null

    [scriptblock]$code = `
    {
        Param([string]$computer,[scriptblock]$work)
        Invoke-Command -ComputerName $computer -ScriptBlock $work
    }

    try
    {
        ## use a runspace so we can have a timeout  
        $runspace = [RunspaceFactory]::CreateRunspace()
        $runspace.Open()
        $command = [PowerShell]::Create().AddScript($code)
        $command.Runspace = $runspace
        $null = $command.AddParameters( @( $computer , $work ) )
        $job = $command.BeginInvoke()

        ## wait for command to finish
        $wait = $job.AsyncWaitHandle.WaitOne( $jobTimeout * 1000 , $false )

        if( $wait -or $job.IsCompleted )
        {
            if( $command.HadErrors )
            {
                Write-Warning "Errors occurred in remote command on $computer :`n$($command.Streams.Error)"
            }
            else
            {
                $results = $command.EndInvoke($job)
                if( ! $results )
                {
                    Write-Warning "No data returned from remote command on $computer"
                }
            }   
            ## if we do these after timeout too then takes an age to return which defeats the point of running via a runspace
            $command.Dispose() 
            $runSpace.Dispose()
        }
        else
        {
            Write-Warning "Job to retrieve info from $computer is still running after $jobTimeout seconds so aborting"
            $null = $command.BeginStop($null,$null)
            ## leaking command and runspace but if we dispose it hangs
        }
    }   
    catch
    {
        Write-Error "Failed to get remote info from $computer : $($_.ToString())"
    }
    $results
}

Export-ModuleMember -Function Get-RemoteInfo
