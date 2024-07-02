# Script for archiving a user's mailbox to a pst file and subsequent deletion
#v.1.4b
#Added connection error handling
#Added mailbox delegation check
#Changes to Exchange 2016, removed running as administrator

Function Check-RunAsAdministrator {
    #Get current user context
    $CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())

    #Check user is running the script is member of Administrator Group
    if ($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        Write-host "Script is running with Administrator privileges!"
    } else {
        #Create a new Elevated process to Start PowerShell
        $ElevatedProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell";

        # Specify the current script path and name as a parameter
        $ElevatedProcess.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"

        #Set the Process to elevated
        $ElevatedProcess.Verb = "runas"

        #Start the new elevated process
        [System.Diagnostics.Process]::Start($ElevatedProcess)

        #Exit from the current, unelevated, process
        Exit
    }
}
#Check Script is running with Elevated Privileges
#Check-RunAsAdministrator
Import-Module ActiveDirectory

$filials = @(
    "Branch1", "Branch2", "Branch3", "Branch4", "Branch5",
    "Branch6", "Branch7", "Branch8", "Branch9", "Branch10",
    "Branch11", "Branch12", "Branch13", "Branch14", "Branch15",
    "Branch16", "Branch17", "Branch18"
)
$DC = "dc1.contoso.com" #Domain controller where everything will happen except group membership

Function Check-PstDirectory {
    While (-not $User.st) {
        # Check for the presence of a branch in the account to determine or create a subfolder
        Write-Host "User branch not specified!"
        [int]$n = 1
        # Display a list of branches and select a branch from the list
        for ($i = 0; $i -lt $filials.Count; $i += 1) {
            Write-Host "$n $filials[$i]"
            $n = $n + 1
        }
        do {
            Write-Host "Select branch (1-$($filials.count.ToString())):" -Separator ""
            [int]$userCity = Read-Host
        } until (1 -le $userCity -and $userCity -le $filials.Count)
        
        $filialId = ($userCity - 1)
        $filialChoice = ""
        While ($filialChoice -notlike "[y|n]") {
            # Dialog for saving the branch
            Write-Host "User branch - " -NoNewline
            Write-Host $filials[$filialId] -NoNewline -ForegroundColor Yellow
            Write-Host ". Save? (Y/N):" -NoNewline
            $filialChoice = Read-Host
        }
        if ($filialChoice -eq "y") {
            $User.st = $filials[$filialId]
            Set-ADUser $User -State $User.st -Server $DC
        } else {
            Write-Host "You MUST select a branch office" # Cannot NOT specify a branch
        }
    }

    $pstDirectory = ("\\fileserver.contoso.com\P$\PST MAILBOXES\" + $User.st + "\")
    if (Test-Path $pstDirectory) {
        # Check for the existence of the path
    } else {
        mkdir $pstDirectory
    }
}

Function ArchiveAndDelete-Mailbox {
    trap [System.Management.Automation.Remoting.PSRemotingTransportException] {
        Write-Host "Connection lost..."
        $Error
        Get-PSSession
        Write-Host "Trying to reconnect"
        Remove-PSSession -Name exchange # Disconnecting exchange session
        $session_exchange = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.contoso.com/PowerShell –Authentication  Kerberos -Name exchange
        Import-PSSession $session_exchange -CommandName New-MailboxExportRequest, Get-MailboxExportRequest, Get-MailboxExportRequestStatistics, Remove-MailboxExportRequest, Disable-Mailbox
        Continue
    }
    trap {
        Write-Host "Something went wrong..."
        $Error
    }

    $session_exchange = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.contoso.com/PowerShell –Authentication  Kerberos -Name exchange
    Import-PSSession $session_exchange -CommandName New-MailboxExportRequest, Get-MailboxExportRequest, Get-MailboxExportRequestStatistics, Remove-MailboxExportRequest, Disable-Mailbox
    Check-PstDirectory # Since there may be no branch, you need to re-request data from AD

    $User = Get-ADUser -Filter "samAccountName -eq '$UserName'" -Properties memberof, msRTCSIP-PrimaryUserAddress, name, st -Server $DC
    $pstPath = "\\fileserver.contoso.com\P$\PST MAILBOXES\" + $User.st + "\" + $User.Name + ".pst"
    Get-MailboxExportRequest | where {$_.Name -eq $User.Name} | Remove-MailboxExportRequest -Confirm:$false -WarningAction SilentlyContinue # Delete request if exists

    New-MailboxExportRequest -Mailbox $User.SamAccountName -Name $User.Name -FilePath $pstPath -confirm:$false -WarningAction SilentlyContinue -BadItemLimit 50
    $Stat = Get-MailboxExportRequest | where {$_.Name -eq $User.Name} | Get-MailboxExportRequestStatistics
    $Status = $Stat.Status.Value

    while (($Status -eq "Queued") -or ($Status -eq "InProgress")) {
        # Tracking progress
        $Stat = Get-MailboxExportRequest | where {$_.Name -eq $User.Name} | Get-MailboxExportRequestStatistics
        $Status = $Stat.Status.Value
        $PercentComplete = $Stat.PercentComplete
        $BytesTransferred = $Stat.BytesTransferred
        if ($PercentComplete -eq "") {
            $PercentComplete = 0
        } else {
            $PercentComplete = $Stat.PercentComplete
        }
        Write-Host "Archiving status is:" $Status
        Write-Host "Completed, %:" $PercentComplete
        Write-Host "Bytes transferred:" $BytesTransferred
        Write-Progress -Activity "Archiving... $PercentComplete % complete" -PercentComplete $PercentComplete -Status "Archiving user's mailbox"
        Start-Sleep -s 10
    }

    $Stat = Get-MailboxExportRequest | where {$_.Name -eq $User.Name} | Get-MailboxExportRequestStatistics -IncludeReport
    $ExportLog = $Stat.Report
    $logsDir = "//fileserver.contoso.com/P$/PST Mailboxes/logs/"
    if (Test-Path $logsDir) {
        # Check for the existence of the path
    } else {
        mkdir $logsDir
    }
    $logFilePath = "$logsDir" + $User.Name + ".txt"
    $ExportLog | Out-File -FilePath "$logFilePath" -Append

    if ($Stat.Status.Value -eq "Completed") {
        Write-Host "User mailbox saved to $pstPath" -ForegroundColor Green
        Write-Output "User mailbox saved to $pstPath" | Out-File -FilePath $logFilePath -Append
        # Add condition handling!!!
        if ($toDelete -eq "y") {
            Write-Host "Deleting user's mailbox..."
            Disable-Mailbox $user.SamAccountName -Confirm:$false
            Write-Host "Mailbox deleted" -ForegroundColor Green
        } else {
            Write-Host "User mailbox will not be deleted" -ForegroundColor Yellow
        }
    } else {
        Write-Host "User mailbox not saved! Exchange account will not be disabled!" -ForegroundColor Red
        Write-Host "USER MAILBOX WILL NOT BE DELETED!" -ForegroundColor Red
        Write-Output "User mailbox not saved! Exchange account will not be disabled!" | Out-File -FilePath $logFilePath -Append
    }

    Get-MailboxExportRequest | where {$_.Name -eq $User.Name} | Remove-MailboxExportRequest -Confirm:$false -WarningAction SilentlyContinue
    Remove-PSSession -Name exchange # Disconnecting exchange session
}

Write-Host "Enter the user's login whose mailbox you want to save (e.g., j.smith): " -NoNewline
$UserName = Read-Host ""
$User = ""
$User = Get-ADUser -Filter "samAccountName -eq '$UserName'" -Properties memberof, msRTCSIP-PrimaryUserAddress, name, st, msExchDelegateListLink -Server $DC

if ($User) {
    Write-Host "User found:" $User
    $isDelegated = $User.msExchDelegateListLink
    if ($isDelegated) {
        Write-Host "Mailbox connected to users:" -ForegroundColor Yellow
        $isDelegated
    }
    Write-Host "User's mailbox will be moved to a pst file"
    $toDelete = "" # Condition to delete the mailbox or only archive it
    While ($toDelete -notlike "[y|n]") {
        Write-Host "Do you need to DELETE the mailbox?" -ForegroundColor Red
        Write-Host "Choose N (do not delete) only if you need to receive and/or send from this mailbox! (Y/N): " -ForegroundColor Red -NoNewline
        $toDelete = Read-Host
    }
    $Choice = ""
    While ($Choice -notlike "[y|n]") {
        $Choice = Read-Host "Do you REALLY want to continue?(y/n)"
    }
    if ($Choice -eq "y") {
        Write-Host "Trying to transfer the mailbox..."
        ArchiveAndDelete-Mailbox
        Write-Host "Press any key to exit..."
        $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") # press any key
    } else {
        Write-Host "Houston, abort launch!"
        Write-Host "Press any key to exit..."
        $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
} else {
    Write-Host "User not found, check the login"
    Write-Host "Press any key to exit..."
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") # press any key
}
