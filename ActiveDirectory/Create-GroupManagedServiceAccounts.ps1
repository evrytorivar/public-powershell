<#
.SYNOPSIS
	Creates Group Managed Service Accounts for SQL Server or General purpose.
.DESCRIPTION
	Creates Group Managed Service Accounts for SQL Server or General purpose.
    All parameters are mandatory.
.PARAMETER ServerName
	The server on which GMSA is going to be used.
.PARAMETER UserOU
    String: distinguished name of the OU where users will be created.
.PARAMETER GroupOU
    String: distinguished name of the OU where groups will be created.
.PARAMETER Purpose
    String: SQL (for SQL server use) or General (for other GMSA)
.PARAMETER FQDN
    String: the domain where users are located (domain.local, corpnet.loc, etc.)
.EXAMPLE
    Create GMSA for use on SQL server "SQLSRV001" in domain "domain.local", provided userOU and groupOU
	.\Create-GroupManagedServiceAccount.ps1 -ServerName "SQLSRV001" -UserOU "OU=Users,OU=Organization,DC=domain,DC=local" -GroupOU "OU=Groups,OU=Organization,DC=domain,DC=local" -Purpose SQL -FQDN "domain.local"
.INPUTS
	It is not possible to pipe output to this script.
.OUTPUTS
    No outputs from this script. Cannot be piped to cmdlets.
.NOTES
	NAME: Create-GroupManagedServiceAccounts.ps1
	VERSION: BETA
	AUTHOR: Tor Ivar Larsen
    CREATED: 2015-12-21
	LASTEDIT: 2015-12-21
#>
[CMDLetBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [string]$ServerName,
    [Parameter(Mandatory=$true)]
    [string]$UserOU,
    [Parameter(Mandatory=$true)]
    [string]$GroupOU,
    [Parameter(Mandatory=$true)]
    [string]$Purpose,
    [Parameter(Mandatory=$true)]
    [string]$FQDN
)

Try
{
    Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
    Write-Error "Could not load module Active Directory:"
    Exit
}

$KDSRootKeyExists = Get-KDSRootKey

If (!$KDSRootKeyExists)
{
    Write-Warning "KDS Root Key is not configured in AD."

    Write-Host "Configure KDS Root Key by doing the following:"
    Write-Host "##########################################################################################################################"
    Write-Host ""
    Write-Host "1. Log into a domain controller or administrative server with RSAT ADDS installed."
    Write-Host ""
    Write-Host "2. Run either of these commands in an elevated PowerShell console:"
    Write-Host ""
    Write-Host "'Add-KDSRootKey -EffectiveImmediately' will add a KDS Root Key that becomes effective some time during the next 10 hours."
    Write-Host ""
    Write-Host "'Add-KDSRootKey -EffectiveTime ((Get-Date).AddHours(-10))' will add a KDS Root Key that actually is effective immediately."
    Write-Host ""
    Exit
}

Switch ($Purpose)
{
    "SQL" {

        $Accounts = @{  "ag" = "SQL Server Agent";
                        "as" = "SQL Server Analysis Services";
                        "db" = "SQL Server Database Engine";
                        "is" = "SQL Server Integration Services";
                        "rs" = "SQL Server Reporting Services"}

        Foreach ($Key in ($Accounts.Keys))
        {
            $Success = $null
            $UserName = "`_sql$Key$ServerName`-ms"
            $GroupName = "acl.msg._sql$Key$ServerName-ms"
            $UserDescription = "$($Accounts.Item($Key)) $ServerName"
            $GroupDescription = "$($Accounts.Item($Key)) MSSQLSERVER"
            $DnsHostName = "$UserName`.$FQDN"

            Write-Host 'New-ADGroup -Name '$GroupName' -GroupScope Global -Path "'$GroupOU'" -Description "'$Description'"'
            Write-Host 'Add-ADGroupMember -Identity "'$GroupName'" -Members "'$($ServerName)`$'"'
            Write-Host 'New-ADServiceAccount -name '$UserName' -Description "'$UserDescription'" -Path "'$UserOU'" -DnsHostName '$DnsHostName' -PrincipalsAllowedToRetrieveManagedPassword "'$GroupName'"'
        }
    }

    "General" {
        $Success = $null
        $UserName = "`_$ServerName`-ms"
        $GroupName = "acl.msg.$ServerName-ms"
        $UserDescription = "Group Managed Service Account used by $ServerName"
        $GroupDescription = "Group allowed to retrieve password for user $UserName"
        $DnsHostName = "$UserName`.$FQDN"

        Write-Host 'New-ADGroup -Name '$GroupName' -GroupScope Global -Path "'$GroupOU'" -Description "'$Description'"'
        Write-Host 'Add-ADGroupMember -Identity "'$GroupName'" -Members "'$($ServerName)`$'"'
        Write-Host 'New-ADServiceAccount -name '$UserName' -Description "'$UserDescription'" -Path "'$UserOU'" -DnsHostName '$DnsHostName' -PrincipalsAllowedToRetrieveManagedPassword "'$GroupName'"'

    }
    default {Write-Host "Purpose not recognized. Try again."}
}

<#
These commands must be run on the server in question after a reboot.
Install-ADServiceAccount _sqlag001-ms
Install-ADServiceAccount _sqldb001-ms
Install-ADServiceAccount _sqlas001-ms
Install-ADServiceAccount _sqlrs001-ms
Install-ADServiceAccount _sqlis001-ms
#>
