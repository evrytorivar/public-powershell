<#
When doing UPS Sync from Sharepoint, you can encounter error 6306 and "unable to process create message".
https://social.msdn.microsoft.com/Forums/en-US/b0f3d2b7-736f-432c-bc68-783648454cb3/error-6306-and-unable-to-process-create-message-when-trying-to-create-a-ad-connection?forum=sharepointadminprevious
The forum post mentioned points to a possible explanation, namely some invalid Auxiliary classes in AD Schema.
Rather than using the provided ldifde-command, I wanted to use powershell, and wrote a very simple and very small script for this purpose.
This script will get all the Auxiliary classes, and export them to a csv in the folder where script is executed.
No inputs from pipeline or outputs to pipeline from this script.
It goes without saying that this script must be executed from an admin-server or domain controller.
#>


([DirectoryServices.ActiveDirectory.ActiveDirectorySchema]::GetCurrentSchema()).FindAllClasses() |
    Where-Object {$_.Type -eq "Auxiliary"} |
    Select-Object Name, subClassOf |
    Export-Csv -Path "$PSScriptRoot\Exported_Schema_Classes_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv" -Encoding UTF8 -NoTypeInformation -Delimiter ";"
