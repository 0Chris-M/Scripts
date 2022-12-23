## Powershell Scripts (.ps1, .psm1, .psd1, etc)



### Incident Response
| Script | Description|
| ------------- |:-------------:|
| Export-Drive.ps1 | Collates list of all files written to disk over previous 4 months. Can adjust timeframe to suit use case.  |
| Get-WinEvents.ps1 | Grabs Windows Event Logs and compresses to zip archive for easier downloading. |


### Office Hardening
| Script | Description|
| ------------- |:-------------:|
| Office-Macros_Disable_All.ps1 | Iterate through all SIDs on machine and set appropriate reg keys to disable all Macro functionality.  |
| Office-Macros_Restore_Disabled.ps1 | Restore all settings/reg keys changed in above script.     |
| Office-Macros_Securely_Enable.ps1 | Iterate through all SIDs on machine and securely allow macros through Trusted Locations. |
| Office-Macros_Restore_Securely_Enable.ps1  | Restore all settings/reg keys changed in above script. |
| Office_Hardening.ps1 | Implements several strategies to reduce the Attack Surface of Microsoft Office environments. |
| Office_Restore_Hardening.ps1 | Restore all settings/reg keys changed in above script. |
| Office_Follina_Mitigation.ps1 | Descr |




### ToDo's
- [ ] Office Hardening -> Echo script progress and tasks completed
- [ ] Office Hardening -> Reassess loaded hives code and for loop
- [ ] Office Hardening -> Test Zone.Identifier for proper Internet Macros blocking
