#Requires -Version 5.1
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0" }

<#
    .SYNOPSIS
        This PowerShell script automates mailbox delegation settings based on data from a CSV file.
        This script requires the module ExchangeOnlineManagement version 3.0 or higher

    .DESCRIPTION

    .NOTES

    .COMPONENT

    .EXAMPLE
        PS> Set-MailboxDelegationFromCSV -MailboxID 'John' -CsvPath C:\temp\Data.csv

    .LINK

    .Parameter MailboxId
        Specifies the mailbox to which the delegations will be applied.
        This should contain the unique identifier of the target mailbox. For example:
            Name
            Alias
            Distinguished name (DN)
            Canonical DN
            Domain\Username
            Email address
            GUID
            LegacyExchangeDN
            SamAccountName
            User ID or user principal name (UPN)

    .Parameter CsvPath

        Specifies the file path to the CSV file containing the mailbox delegation details.
        The CSV file requires following format, where mail@example.com represents the delegate's email address: 

        User;FullAccess;SendAs;SendOnBehalf
        mail@example.com;1;0;0

#>

[CmdletBinding()]

param(
    [Parameter(Mandatory = $true)]
    [string]$MailboxId,
    [Parameter(Mandatory = $true)]
    [string]$CsvPath
)

if (!(Test-Path $csvPath -PathType Leaf)) {
    throw "CSV-File not found at $csvPath"
}

[string]$csvFile = Get-Content $csvPath -Raw
$expectedFormat = '(^User;FullAccess;SendAs;SendOnBehalf\r?\n)([^;\r\n\s]{1,}(;(\b1\b|\b0\b)){3}(\r?\n){0,}){1,}$'
<#
(^User;FullAccess;SendAs;SendOnBehalf\r?\n) - headers with a new line, \r?\n for UNIX-like new line without \r
[^;\r\n\s] - any symbol except semicolon, new line and space
(;(\b1\b|\b0\b)){3} - semicolon followed by either 1 or 0, exactly 3 times
(\r?\n){0,} - can be a new line, but also EOF
{1,}$ - the whole sentence min once after headers
#>
if ($csvFile.trim() -notmatch $expectedFormat){
    throw "Error parsing CSV-File - invalid format!"
}

[array]$csvData = $csvFile | ConvertFrom-Csv -Delimiter ";"

#Converting CSV string values to their appropriate data types
$resultSettings = @()
foreach ($item in $csvData) {
    $settings = [PSCustomObject]@{
        User = [string]$item.User
        FullAccess = [bool]([int]($item.FullAccess))
        SendAs = [bool]([int]($item.SendAs))
        SendOnBehalf = [bool]([int]($item.SendOnBehalf))
        AssignedPermissions = [System.Collections.ArrayList]::new() #Flags after assignment
    }
    $resultSettings += $settings
}

Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline -ShowBanner:$false

try{
    $mailbox = Get-Mailbox -Identity $MailboxId -ErrorAction Stop
}
catch{
    throw
}

$warnings = [System.Collections.ArrayList]::new()
Function Catch-Error {

<#
Handles processing errors based on the stage at which the script fails.
Failing to process a delegate or a permission does not abort the execution of the remaining components.
Warnings will be displayed in a batch at the end of the execution.
#>

    $errorArgs = @{
        ErrorAction = 'Continue'
        Exception = $Error[0].Exception
    }
    if ($assigningPermission) {
        $warnings.Add("Processing of $($item.User) failed. Permission $assigningPermission not set.") | Out-Null
    }
    else {
        $warnings.Add("Processing of $($item.User) failed. Permissions not set.") | Out-Null
    }
    Write-Error @errorArgs
    $script:assigningPermission = $null #resetting the variable every time the processing fails

}

Function Set-AssignedPermissionFlag {

    $item.AssignedPermissions.Add($assigningPermission) | Out-Null
    $script:assigningPermission = $null #resetting the variable every time after it was set as a flag

}

$cmdArgs = @{
    Identity = $mailbox
    Confirm = $false
    ErrorAction = 'Stop'
}

foreach ($item in ($resultSettings | Where-Object {$_.FullAccess -or $_.SendAs -or $_.SendOnBehalf})){
    try {
        Get-Recipient -Identity $item.User -RecipientType UserMailbox,MailUser,MailUniversalSecurityGroup -ErrorAction Stop | Out-Null
        if ($item.FullAccess) {
            try {
                $assigningPermission = "Full Access"
                Add-MailboxPermission @cmdArgs -AccessRights FullAccess -User $item.User -InheritanceType All -AutoMapping:$false | Out-Null
                Set-AssignedPermissionFlag
            }
            catch{
                Catch-Error
            }
        }
        if ($item.SendAs) {
            try{
                $assigningPermission="Send As"
                Add-RecipientPermission @cmdArgs -AccessRights SendAs -Trustee $item.User | Out-Null
                Set-AssignedPermissionFlag
            }
            catch{
                Catch-Error
            }
        }
        if ($item.SendOnBehalf) {
            try{
                $assigningPermission="Send on Behalf"
                #-GrantSendOnBehalfTo overwrites the previous values!!!
                [System.Collections.ArrayList]$permissionsSendOnBehalf=(Get-Mailbox -Identity $Mailbox).GrantSendOnBehalfTo
                $permissionsSendOnBehalf.Add($item.User) | Out-Null
                Set-Mailbox @cmdArgs -GrantSendOnBehalfTo $permissionsSendOnBehalf
                Set-AssignedPermissionFlag
                Remove-Variable permissionsSendOnBehalf -Confirm:$false -Force
            }
            catch{
                Catch-Error
            }
        }
    }
    catch {
        if ($Error[0].CategoryInfo.Reason -eq "ManagementObjectNotFoundException"){
            $warnings.Add($Error[0].Exception) | Out-Null
        }
        else{
            Catch-Error
        }
    }
}

$warnings | Write-Warning

Write-Output $resultSettings | Where-Object {$_.AssignedPermissions} | Select-Object User, AssignedPermissions
