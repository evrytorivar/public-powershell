<#
.SYNOPSIS
	Installs windows updates on local computer.
.DESCRIPTION
	The script fetches all windows updates available from either Microsoft Update or WSUS (depends on configuration of host), and installs them.
.PARAMETER AutoRestart
	If switch is used, the computer will automatically restart when updates are done installing.
.PARAMETER Criteria
    Can be used if you need to provide more detailed criteria for Windows Update search.
.EXAMPLE
    Install all windows updates and restart computer automatically
	Install-WindowsUpdates.ps1 -AutoRestart
.INPUTS
	It is not possible to pipe output to this script.
.OUTPUTS
    No outputs from this script. Cannot be piped to cmdlets.
.NOTES
	NAME: Install-WindowsUpdates.ps1
	VERSION: 1.0
	AUTHOR: Tor Ivar Larsen
    CREATED: 2015-12-21
	LASTEDIT: 2015-12-21
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$False,Position=1)]
    [Switch]$AutoRestart,
    [Parameter(Mandatory=$False,Position=2)]
    [String]$Criteria="IsInstalled=0 and IsAssigned=1 and IsHidden=0 and Type='Software'"
)

function Get-WIAStatusValue($value) 
{ 
   switch -exact ($value) 
   { 
      0   {"NotStarted"} 
      1   {"InProgress"} 
      2   {"Succeeded"} 
      3   {"SucceededWithErrors"} 
      4   {"Failed"} 
      5   {"Aborted"} 
   }  
}

If ($AutoRestart)
{
    Write-Host "#######################################################################" -ForegroundColor Red
    Write-Host "#AutoRestart switch given. Script will automatically restart computer!#" -ForegroundColor Red
    Write-Host "#######################################################################" -ForegroundColor Red
}

#Update session
$UpdateSession = New-Object -ComObject Microsoft.Update.Session

#Search for relevant updates.
Write-Verbose "Searching for updates."
$Searcher = $UpdateSession.CreateUpdateSearcher()
$SearchResult = $Searcher.Search($Criteria).Updates


#Install updates.
Write-Verbose "Preparing for installation."
$Installer = New-Object -ComObject Microsoft.Update.Installer
$RebootRequired = $False

$Counter = 0
$UpdateCount = ($SearchResult | Measure-Object).Count
Foreach ($Update in $SearchResult)
{
    $Counter++
    Write-Progress -Activity "Installing Windows Updates" "Processing Update $Counter of $UpdateCount`: $($Update.Title)" -PercentComplete (($Counter / $UpdateCount) * 100)
    Write-Output "Processing $($Update.Title):"

    #Create updates collection
    $UpdatesCollection = New-Object -ComObject Microsoft.Update.UpdateColl
    #Accept EULA if necessary
    if ( $Update.EulaAccepted -eq 0 ) { $Update.AcceptEula() }
    $TMP = $UpdatesCollection.Add($Update)

    #Download update
    $UpdatesDownloader = $UpdateSession.CreateUpdateDownloader() 
    $UpdatesDownloader.Updates = $UpdatesCollection
    $DownloadResult = $UpdatesDownloader.Download()
    $FGC = "Green"
    If ($DownloadResult.ResultCode -eq 3) { $FGC = "Yellow" }
    ElseIf (($DownloadResult.ResultCode -eq 4) -or ($DownloadResult.ResultCode -eq 5)) { $FGC = "Red" }
    
    $Message = "  - Download {0}" -f (Get-WIAStatusValue $DownloadResult.ResultCode) 
    Write-Host $Message -ForegroundColor $FGC

    #Install update
    $UpdatesInstaller = $UpdateSession.CreateUpdateInstaller()
    $UpdatesInstaller.Updates = $UpdatesCollection 
    $InstallResult = $UpdatesInstaller.Install()
    $FGC = "Green"
    If ($InstallResult.ResultCode -eq 3) { $FGC = "Yellow" }
    ElseIf (($InstallResult.ResultCode -eq 4) -or ($InstallResult.ResultCode -eq 5)) { $FGC = "Red" }

    $Message = "  - Install {0}" -f (Get-WIAStatusValue $InstallResult.ResultCode)
    Write-Host $Message -ForegroundColor $FGC

    If (($InstallResult.ResultCode -eq 2) -or ($InstallResult.ResultCode -eq 3))
    {
        # Specify the registry key 
        $Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install"
        $Name = "LastSuccessTime"
        $date = (Get-Date).AddHours(-1)
        $Value = Get-Date -Date $date -Format "yyyy-MM-dd HH:mm:ss"

        If (Test-Path $Path) { Set-ItemProperty -Path $Path -Name $Name -Value "$Value" }
        Else
        {
            New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results" -Name "Install" -Force
            Set-ItemProperty -Path $Path -Name $Name -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
        }
    }

    If ($InstallResult.rebootRequired)
    {
        $RebootRequired = $True
    }
    if ((Get-Service -Name wuauserv).Status -ne "Running")
    { 
        Write-Output "WSUS Service is stopped: Restarting the service" 
        Set-Service  wuauserv -StartupType Automatic 
        Start-Service -name wuauserv 
    }
}

#Reboot if required by updates.
If ($RebootRequired -and $AutoRestart)
{
    Restart-Computer -Force -Confirm:$False
}
ElseIf ($RebootRequired)
{
    Write-Output "Reboot required by one or more updates, but '-AutoRestart' switch not present. Add switch '-AutoRestart' to automatically restart after all updates are installed."
}
Else
{
    Write-Output "Reboot not required."
}
