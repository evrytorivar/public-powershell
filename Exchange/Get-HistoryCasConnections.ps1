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
