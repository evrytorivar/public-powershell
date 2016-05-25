<#
.SYNOPSIS
    Collects logs from server named in parameter, or localhost if no computername is provided.
.DESCRIPTION
    Collects logs from server named in parameter. After initial collection of logs
    the script will compile a list of users listed in the logfiles.
.PARAMETER ComputerName
    Either FQDN or NETBIOS-name of Exchange-server with ClientAccessRole installed.
    Parameter is mandatory.
.PARAMETER LogDir
    The path where Exchange RPC ClientAccess logs can be found.
    Parameter is not mandatory.
.EXAMPLE
    Get logs from ComputerName "MAILSRV001", default logdir
    .\Get-RPCHistory.ps1 -ComputerName "MAILSRV001"
.INPUTS
    Does not accept pipeline input.
.OUTPUTS
    Does not provide pipeline output.
.NOTES
    NAME: Get-RPCHistory
    VERSION: 1.0
    AUTHOR: Tor Ivar Larsen
    CREATED: 22.01.2016
    LASTEDIT: 22.01.2016
#>
Param
(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName=$null,
    [Parameter(Mandatory=$false)]
    [string]$LogDir = "\Microsoft\Exchange Server\V14\Logging\RPC Client Access",
    [Parameter(Mandatory=$false)]
    [int]$days = -5
)

If ((!$ComputerName) -or ($ComputerName -like "localhost"))
{
    $ComputerName = $env:COMPUTERNAME
    $LogDir = "$env:ProgramFiles\Microsoft\Exchange Server\V14\Logging\RPC Client Access"
}
Else
{
    $LogDir = "\\$ComputerName\C$\Program Files\Microsoft\Exchange Server\V14\Logging\RPC Client Access"
}

$Mailboxes = @()
$Done = @()
Foreach ($LogFile in (Get-ChildItem -Path $LogDir -Filter "*.LOG" | Where-Object {$_.LastWriteTime -gt (Get-Date).AddDays($days)}))
{
    Get-Content -Path "$($LogFile.FullName)" |
    Foreach-Object {
        If (!($_.StartsWith("#"))) {
            $Arr = $_.Split(",")
            $DateTime = $Arr[0]
            $DateTime = $DateTime.Substring(0,$DateTime.Length-1)
            $ClientName = $Arr[3]
            If ($Done -notcontains $ClientName) {
                #$Mailbox = Get-Mailbox "$ClientName" | Select-Object Name,Alias,PrimarySMTPAddress,ServerName
                $Mailbox = New-Object -TypeName PSObject
                $Mailbox | Add-Member -MemberType NoteProperty -Name "ClientName" -Value $ClientName
                $Mailbox | Add-Member -MemberType NoteProperty -Name "ConnectedToCAS" -Value $ComputerName
                $Mailbox | Add-Member -MemberType NoteProperty -Name "DateTime" -Value $DateTime
                $Mailboxes += $Mailbox
                $Done += $ClientName
            }
        }
    }
}

$Mailboxes | Sort-Object -Property ClientName | Export-Csv "C:\scripts\Exchange\Users_$ComputerName.csv" -Delimiter ";" -Encoding UTF8 -NoTypeInformation
