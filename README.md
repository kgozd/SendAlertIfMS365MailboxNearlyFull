# SendAlertIfMS365MailboxNearlyFull
This script is a valuable tool for every MS365 administrator. You can schedule its execution in Task Scheduler, and the script will send an alert to your email about users whose mailbox free space is less than 10%.


## Requirements

This script for proper working requires several things:

- PowerShell in version 5.1 (Not Core Edition!!!)
- Active MS365 account 
- An email adress from which messeges would be send (for example gmail)
- Installed 2 powershell Modules (AzureAD, ExchangeOnlineManagement)
- Enabled script execution in PowerShell

## Installation
Set execution policy in powershell to allow script execution

Use these commands to install required modules

```PowerShell
  Install-Module -Name AzureAD
  Install-Module -Name ExchangeOnlineManagement
```
    
Next go to Google account > security> 2FA > Application Password and generate one


Remember to put your configuration emails etc in CredsConfig.json

***Using env variable would be more robust option, I used json for simplicity.
