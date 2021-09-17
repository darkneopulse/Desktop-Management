# This script checks that Office 16 is installed, then creates a registry key that configures delegate sent items to be stored in the From address sent items.
# I recommend backing up your registry prior to running. I take no responisbility for the use of this script.
# Documentation = https://docs.microsoft.com/en-us/exchange/troubleshoot/user-and-shared-mailboxes/sent-mail-is-not-saved
# Created by Darkneopulse

$Office16 = Test-Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Preferences'
if ($office16 -eq $True){
    reg add HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Preferences /v DelegateSentItemsStyle /t REG_DWORD /d 1 /f
    Write-Output "Registry key added successfully"
} 
else {
    Write-Output "Office 16 is not installed or Outlook is not setup"
}
