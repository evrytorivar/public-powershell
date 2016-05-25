<#
.SYNOPSIS
    Digs through provided RpcHistoryLogs (exported with Get-RpcHistory.ps1) and compiles a list.
.DESCRIPTION
    This script will parse RpcHistoryLogs provided and compile a list of the CAS-users that has connected to the
    CAS server.
.PARAMETER RpcHistoryFiles
    String array with absolute paths to RpcHistoryFiles (e.g. "c:\temp\log.csv","C:\temp\log2.csv").
    Parameter is mandatory.
.EXAMPLE
    Parse the logs provided
    .\Get-HistoryCasConnections.ps1 -RpcHistoryFiles "C:\temp\log.csv","C:\temp\log2.csv"
.INPUTS
    Does not accept pipeline input.
.OUTPUTS
    Does not provide pipeline output.
.NOTES
    NAME: Get-RpcCasConnections.ps1
    VERSION: 1.1
    AUTHOR: Tor Ivar Larsen
    CREATED: 22.01.2016
    LASTEDIT: 25.05.2016
#>
[CMDLetBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [string[]]$RpcHistoryFiles
)

$Users = ""
Foreach ($file in $RpcHistoryFiles)
{
    If (Test-Path "$file")
    {
        $Users += Import-Csv -Path "$file" -Delimiter ";" -Encoding UTF8
    }
}

$Mailboxes = @()
$Done = @()

Foreach ($User in $Users)
{
    $ClientName = $User.ClientName
    $Connected = $User.ConnectedToCAS
    $DateTime = $User.DateTime
    
    $ErrorActionPreference = "SilentlyContinue"
    $Mailbox = Get-Mailbox "$ClientName" | Select-Object Name,Alias,PrimarySMTPAddress,ServerName -ErrorAction Stop
    $ErrorActionPreference = "Continue"

    If ($Mailbox)
    {
        $Mailbox | Add-Member -MemberType NoteProperty -Name "CAS-Server" -Value $Connected
        $Mailbox | Add-Member -MemberType NoteProperty -Name "ConnectedTime" -Value $DateTime
        $Mailboxes += $Mailbox
        $Done += $ClientName
        Write-Verbose "Processed: $($Mailbox.Alias)"
    }
    Else
    {
        Write-Output "Could not process $ClientName."
    }
}

Write-Verbose "Users processed: $(($Mailboxes | Measure-Object).Count)"

$Mailboxes | Sort-Object -Property Name | Select-Object Name,Alias,CAS-server,ConnectedTime,PrimarySMTPAddress,ServerName |
    Export-Csv C:\scripts\CONNECTED_CAS_USERS.csv -Delimiter ";" -Encoding UTF8 -NoTypeInformation
