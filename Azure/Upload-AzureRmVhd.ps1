<#
.SYNOPSIS
	Script for uploading a VHD file to Azure Blob Storage.
.DESCRIPTION
    This script simplifies the process of uploading a VHD file to Azure RM.
.PARAMETER VHDFolder
	String containing the folder where VHD is located. No trailing backslash.
.PARAMETER VHDName
    String containing the VHD name of virtual disk to be uploaded. Must include .vhd extension.
.PARAMETER DestinationURL
    String containing the destination URL where VHD will be uploaded. No trailing slash.
.PARAMETER RGName
    String containing the Resource Group name where VHD will be uploaded.
.EXAMPLE
	.\Upload-AzureRMVHD.ps1 -VHDFolder "C:\VHDFiles" -VHDName "MyCustomVHD.vhd" -DestinationURL "https://mystorageblob.blob.core.windows.net/<conainername>" -RGName "ResourceGroup1"
.INPUTS
	No piping to this script.
.OUTPUTS
	No output to pipeline from this script.
.NOTES
	NAME: Upload-AzureRMVHD.ps1
	VERSION: 1.0
	AUTHOR: Tor Ivar Larsen
    CREATED: 2016-04-01
	LASTEDIT: 2016-05-25
#>

[CMDLetBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [string]$VHDFolder,
    [Parameter(Mandatory=$true)]
    [string]$VHDName,
    [Parameter(Mandatory=$true)]
    [string]$DestinationURL,
    [Parameter(Mandatory=$true)]
    [string]$RGName
)


If (!(Test-Path "$VHDFolder\$VHDName"))
{
    Write-Output "ERROR: File not found. Try again."
    Exit
}

$VHDFile = Get-ChildItem -Path "$VHDFolder\$VHDName"
If ($VHDFile.Extension -ne ".vhd")
{
    Write-Output "ERROR: You seem to have provided a file that is not a VHD file (Azure does not support VHDX format, only Fixed VHD). You can convert VHDX to VHD with Convert-VHD on a Hyper-V host or Windows 10 with Hyper-V installed."
    Exit
}


Try
{
    Write-Output "INFORMATION: Checking if you are logged in."
    Get-AzureRMSubscription -ErrorAction Stop
}
Catch
{
    Write-Output "INFORMATION: You are not logged in to Azure Resource Manager. Please log in."
    Login-AzureRMAccount
}

#Select azureRMsubscription
Write-Output "INFORMATION: Select your Azure Subscription in the gridview."
$Subscription = Get-AzureRMSubscription | Out-GridView -PassThru

Select-AzureRMSubscription -SubscriptionID $($Subscription.SubscriptionId)
Write-Output "INFORMATION: Subscription selected: $($Subscription.SubscriptionName)"
   
Write-Output "INFORMATION: $VHDFolder\$VHDName found. Initiating import process."
Add-AzureRMVhd -Destination "$DestinationURL/$VHDName" -LocalFilePath "$VHDFolder\$VHDName" -ResourceGroupName "$RGName"
