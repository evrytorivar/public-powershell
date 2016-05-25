#Requires -Module RemoteAccess
<#
.SYNOPSIS
    Provides an easy way to list NRPT on the Remote Access Server.
.DESCRIPTION
    Fetches NRPT from Remote Access server, and outputs objects to the pipeline.
.EXAMPLE
    .\Get-Nrpt.ps1 | Export-Csv -Path NRPT.csv -Delimiter ";" -NoTypeInformation -Encoding UTF8
.INPUTS
    Does not accept pipeline input.
.OUTPUTS
    Outputs objects to the pipeline with two noteproperties: NameSpace and NameServers.
.NOTES
    NAME: Get-Nrpt
    VERSION: 1.01
    AUTHOR: Tor Ivar Larsen
    CREATED: 2016-05-11
    LASTEDIT: 2016-05-19
#>

Foreach ($Entry in (Get-DAClientDnsConfiguration).NrptEntry)
{
    $NameSpace = ($Entry | Select-Object -ExpandProperty NameSpace).ToString()
    $ProxyType = $Entry.DirectAccessProxyType
    $NameServers = ""

    If ($Entry.DirectAccessDnsServers)
    {
        $NameServers = ($Entry | Select-Object -ExpandProperty DirectAccessDnsServers).ToString()
    }

    $obj = New-Object -TypeName psobject
    $obj | Add-Member -MemberType NoteProperty -Name 'NameSpace' -Value $NameSpace
    $obj | Add-Member -MemberType NoteProperty -Name 'NameServers' -Value $NameServers
    $obj | Add-Member -MemberType NoteProperty -Name 'DirectAccessProxyType' -Value $ProxyType

    $obj
        
}
