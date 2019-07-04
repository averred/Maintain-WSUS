#Requires -Version 3
<#
    .SYNOPSIS
        Script designed for daily maintenance of WSUS.
    .DESCRIPTION
        Synchronises with upstream WSUS, declines all unwanted updates, accepts
        license agreements and cleans up WSUS files/SQL Server.

        Must be run on the WSUS server locally.
    .INPUTS
        None
    .OUTPUTS
        Progress
        Log file stored in .\Logs\<Date>_<Time>_Maintain-WSUS.ps1.log
    .NOTES
        Version:        1.1
        Author:         Talha Khan <talha@averred.net>
        Creation Date:  01/05/2019
        Last modified:  04/07/2019
    .LINK
        https://github.com/averred/Maintain-WSUS
    .EXAMPLE
        .\Maintain-WSUS.ps1
#>

# Helper functions
Function Invoke-Exe{
    [CmdletBinding(SupportsShouldProcess=$true)]

    param(
        [parameter(mandatory=$true,position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Executable,

        [parameter(mandatory=$false,position=1)]
        [string]
        $Arguments
    )

    If ($Arguments -eq '') {
        Write-Verbose "Running $ReturnFromEXE = Start-Process -FilePath $Executable -ArgumentList $Arguments -NoNewWindow -Wait -Passthru"
        $ReturnFromEXE = Start-Process -FilePath $Executable -NoNewWindow -Wait -Passthru
    }
    Else {
        Write-Verbose "Running $ReturnFromEXE = Start-Process -FilePath $Executable -ArgumentList $Arguments -NoNewWindow -Wait -Passthru"
        $ReturnFromEXE = Start-Process -FilePath $Executable -ArgumentList $Arguments -NoNewWindow -Wait -Passthru
    }
    Write-Verbose "Returncode is $($ReturnFromEXE.ExitCode)"
    Return $ReturnFromEXE.ExitCode
}

# Setup
$ScriptFolder = Split-Path $MyInvocation.MyCommand.Path -Parent
$ScriptName = Split-Path $MyInvocation.MyCommand.Path -Leaf
$WsusDBMaintenanceFile = '{0}\WsusDBMaintenance.sql' -f $ScriptFolder

# Logging
$LogFolder = '{0}\Logs' -f $ScriptFolder
$LogFile = '{0}\{1}_{2}.log' -f $LogFolder, (Get-Date -Format yyyyMMdd_HHmmss), $ScriptName
$LogRetentionDays = 7

Start-Transcript -Path $LogFile
Write-Output ('Cleaning up logs older than {0} days' -f $LogRetentionDays)
Get-ChildItem -Path $LogFolder | Where-Object {($_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays))} | Remove-Item -Verbose

# Connect to DB
$WSUSDB = '\\.\pipe\Microsoft##WID\tsql\query'

If (!(Test-Path $WSUSDB) -eq $True) {
    Write-Warning ('Could not access the SUSDB: {0}' -f $WSUSDB)
}

ElseIf (!(Test-Path $WsusDBMaintenanceFile) -eq $True) {
    Write-Warning ('Could not access {0}' -f $WsusDBMaintenanceFile)
    Write-Warning "Make sure you have downloaded the file from https://gallery.technet.microsoft.com/scriptcenter/6f8cde49-5c52-4abd-9820-f1d270ddea61#content"
}
Else {
    Write-Output ('Running from: {0}' -f $RunningFromFolder)
    Write-Output ('Using SQL FIle: {0}' -f $WsusDBMaintenanceFile)
    Write-Output ('Using DB: {0}' -f $WSUSDB)

    # Get and Set the WSUS Server target
    $WSUSSrv = Get-WsusServer -Name $env:COMPUTERNAME -PortNumber 8530
    Write-Output ('Working on {0}' -f $WSUSSrv.name)

    # Choose Languages
    $WSUSSrvCFG = $WSUSSrv.GetConfiguration()

    # Synchronization
    $WSUSSrvSubscrip = $WSUSSrv.GetSubscription()
    Write-Output 'Begin WSUS Synchronization'
    $SyncTimer = [system.diagnostics.stopwatch]::StartNew()
    $WSUSSrvSubscrip.StartSynchronization()
    While ($WSUSSrvSubscrip.GetSynchronizationStatus() -ne 'NotProcessing') {
        If ($WSUSSrvSubscrip.GetSynchronizationProgress().TotalItems -ne 0) {
            Write-Progress -PercentComplete (100*$WSUSSrvSubscrip.GetSynchronizationProgress().ProcessedItems/$WSUSSrvSubscrip.GetSynchronizationProgress().TotalItems) -Activity "WSUS Synchronization Progress"
        }
        Start-Sleep 1
    }

    Write-Progress -Activity "WSUS Synchronization Progress" -Status "Ready" -Completed

    $SyncTotal = [math]::Round($SyncTimer.Elapsed.TotalSeconds,0)
    Write-Output ('WSUS Synchronization completed after {0} seconds' -f $SyncTotal)

    #Begin cleanup of WSUS
    Write-Output 'Decline Superseded Updates'
    Get-WsusUpdate -Approval AnyExceptDeclined -Classification All -Status Any | Where-Object -Property UpdatesSupersedingThisUpdate -NE -Value 'None' -Verbose | Deny-WsusUpdate -Verbose

    Write-Output 'Get All Updates except Declined'
    $AllUpdates = Get-WsusUpdate -Approval AnyExceptDeclined -Classification All -Status Any

    Write-Output 'Decline x86 Updates'
    $AllUpdates | Where-Object { $_.Update.Title -like "*X86*" } | Deny-WsusUpdate -Verbose

    Write-Output 'Decline ARM64 Updates'
    $AllUpdates | Where-Object { $_.Update.Title -like "*ARM64*" } | Deny-WsusUpdate -Verbose

    Write-Output 'Decline Preview Updates'
    $AllUpdates | Where-Object { $_.Update.Title -like "*Preview*" } | Deny-WsusUpdate -Verbose

    Write-Output 'Decline Beta Updates'
    $AllUpdates | Where-Object { $_.Update.Title -like "*Beta*" } | Deny-WsusUpdate -Verbose

    Write-Output 'Decline Other Updates'
    $AllUpdates | Where-Object { $_.Update.PublicationState -like "Expired" `
                      -or $_.Update.Title -like "*Windows 10*N,*" `
                      -or $_.Update.Title -like "*Windows 10*N version*" `
                      -or $_.Update.Title -like "*Windows 10 Education,*" `
                      -or $_.Update.Title -like "*consumer editions*" `
                      -or $_.Update.Title -like "*1507*" `
                      -or $_.Update.Title -like "*1511*" `
                      -or $_.Update.Title -like "*1607*" `
    				  -or $_.Update.Title -like "*1703*" `
    				  -or $_.Update.Title -like "*1709*" `
                      -or $_.Update.Title -like "*April 2018 Update*" `
                      -or $_.Update.Title -like "*1803*" } | Deny-WsusUpdate -Verbose

    Write-Output 'Decline Language Pack not en-GB'
    $AllUpdates | Where-Object { $_.Update.Title -like "*Lang*" -and $_.Update.Title -notlike "*en-GB*" } | Deny-WsusUpdate -Verbose

    Write-Output 'Decline Feature update not en-GB'
    $AllUpdates | Where-Object { $_.Update.Title -like "*Feature update to Windows 10*" -and $_.Update.Title -notlike "*en-GB*" } | Deny-WsusUpdate -Verbose

    Write-Output 'Cleanup Obsolete Computers'
    $CleanupObsoleteComputers = Invoke-WsusServerCleanup -UpdateServer $WSUSSrv -CleanupObsoleteComputers
    Write-Output $CleanupObsoleteComputers

    Write-Output 'Decline Expired Updates'
    $DeclineExpiredUpdates = Invoke-WsusServerCleanup -UpdateServer $WSUSSrv -DeclineExpiredUpdates
    Write-Output $DeclineExpiredUpdates

    Write-Output 'Decline Superseded Updates'
    $DeclineSupersededUpdates = Invoke-WsusServerCleanup -UpdateServer $WSUSSrv -DeclineSupersededUpdates
    Write-Output $DeclineSupersededUpdates

    Write-Output 'Cleanup Obsolete Updates'
    $CleanupObsoleteUpdates = Invoke-WsusServerCleanup -UpdateServer $WSUSSrv -CleanupObsoleteUpdates
    Write-Output $CleanupObsoleteUpdates

    Write-Output 'Cleanup Unneeded Content Files'
    $CleanupUnneededContentFiles = Invoke-WsusServerCleanup -UpdateServer $WSUSSrv -CleanupUnneededContentFiles
    Write-Output "Diskspace Freed: $([Math]::Round($(($CleanupUnneededContentFiles).Split(":")[1]/1GB),2)) GB"

    Write-Output 'Compress Updates'
    $CompressUpdates = Invoke-WsusServerCleanup -UpdateServer $WSUSSrv -CompressUpdates
    Write-Output $CompressUpdates

    # Accept license agreements
    [void][reflection.assembly]::LoadWithPartialName('Microsoft.UpdateServices.Administration')
    $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($ENV:ComputerName, $False, 8530)
    $Updates = $WSUS.GetUpdates()
    $License = $Updates | Where-Object { $_.RequiresLicenseAgreementAcceptance }
    $License | Select-Object Title
    $License | ForEach-Object { $_.AcceptLicenseAgreement() }


    # Cleanup the SUDB
    Write-Output 'Defrag and Cleanup DB'
    $Command = 'SQLCMD.EXE'
    $Arguments = '-E -I -S {0} -i {1}' -f $WSUSDB, $WsusDBMaintenanceFile
    $ReturnCode = Invoke-Exe -Executable $Command -Arguments $Arguments
    If ($ReturnCode -ne 0) {
        Write-Warning ('{0} Return Code: {1}' -f $Command, $ReturnCode)
    }
    Else {
        Write-Output ('{0} Return Code: {1}' -f $Command, $ReturnCode)
    }

    Write-Output '*** WSUS Maintenance Complete ***'
}

# End log
Stop-Transcript
