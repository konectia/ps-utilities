# PowerShell Utilities

PowerShell utilities.
Use:
```
powershell -File [utilidad].ps1
```

# Installation
Windows by default only executes signed scripts (see [execution policies](https:/go.microsoft.com/fwlink/?LinkID=135170)) so might be neccessary to update the policy to allow local scripts executions:

It also may be neccesary to install the NuGet client and some dependencies.

# Utilities

## excel-mail-remover.ps1

Search all the excel files in a directory (and subdirectories) that meet the pattern '* _subscribed_members.xlsx'. 
For each file that it finds, it looks for a file with the name that starts the same but changing the suffix '_subscribed_members.xlsx' for '_unsuscribed_members.xlsx' and eliminates the entries of the first file that are in the second one. In addition it also eliminates duplicate entries and orders the output.

Parameters:
 * unsubscribedSuffix: Unsubscribed file suffix ('_unsuscribed_members' by default)
 * subscribedSuffix:  Subscribed file suffix ('_subscribed_members' by default)
 * directory: Directory to process (script direcoty by default)
 * update: If no set, files are not modified (by default is not set)
 * reportXls: If set summary report file (by default empty)

Example of use:
```
powershell .\excel-mail-remover.ps1 -directory C:\data\mails -reportXls summary.xlsx -update
```