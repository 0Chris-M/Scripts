#####################################################################
# Restore hardening


 #####################################################################
                   ## Code for HKEY_USERS ##
 ##################################################################### 
     
# Regex pattern for SIDs
$PatternSID = 'S-1-\d+-\d+-\d+-\d+\-\d+\-\d+$'
 
# Get Username, SID, and location of ntuser.dat for all users
$ProfileList = gp 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object {$_.PSChildName -match $PatternSID} | 
    Select  @{name="SID";expression={$_.PSChildName}}, 
            @{name="UserHive";expression={"$($_.ProfileImagePath)\ntuser.dat"}}, 
            @{name="Username";expression={$_.ProfileImagePath -replace '^(.*[\\\/])', ''}}
 
# Get all user SIDs found in HKEY_USERS (ntuder.dat files that are loaded)
$LoadedHives = gci Registry::HKEY_USERS | ? {$_.PSChildname -match $PatternSID} | Select @{name="SID";expression={$_.PSChildName}}
 
# Get all users that are not currently logged
$UnloadedHives = Compare-Object $ProfileList.SID $LoadedHives.SID | Select @{name="SID";expression={$_.InputObject}}, UserHive, Username
 
# Loop through each profile on the machine
Foreach ($item in $ProfileList) {
    # Load User ntuser.dat if it's not already loaded
    IF ($item.SID -in $UnloadedHives.SID) {
        reg load HKU\$($Item.SID) $($Item.UserHive) | Out-Null
    }
     
     #####################################################################
    # Modify registry entries below:
  
    ##################################################################### 
    # Restore ActiveX
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\common\security -Name disableallactivex -Force

    # Restore add-ins for excel, project, powerpoint, visio, word, access, publisher, and OneNote.
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security -Name disablealladdins -Force    
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\"ms project"\security -Name disablealladdins -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security -Name disablealladdins -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\visio\security -Name disablealladdins -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security -Name disablealladdins -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\access\security -Name disablealladdins -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\publisher\security -Name disablealladdins -Force    
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\onenote\security -Name disablealladdins -Force    
    
    
    ### Restore DDE
      # Common
      Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\common\security -Name blockedextensions -Force          
      
      # Excel
      Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\"external content" -Name disableddeserverlaunch -Force      
      Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\"external content" -Name disableddeserverlookup -Force
      Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\"external content" -Name enableblockunsecurequeryfiles -Force
      Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name DataConnectionWarnings -Force      
      Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name RichDataConnectionWarnings -Force
      Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name WorkbookLinkWarnings -Force
  
      # Word
      Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\options -Name dontupdatelinks -Force      
      Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\options\wordmail -Name dontupdatelinks -Force      
  
  
    # Restore OLE
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\security -Name PackagerPrompt -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\security -Name PackagerPrompt -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\PowerPoint\security -Name PackagerPrompt -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\security -Name PackagerPrompt -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\security -Name PackagerPrompt -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Publisher\security -Name PackagerPrompt -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\security -Name PackagerPrompt -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\onenote\security -Name PackagerPrompt -Force      


    # Restore Excel extension hardening
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security -Name extensionhardening -Force      


    # Restore Office File Validation (OFV) for Excel, PowerPoint, and Word
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\excel\security\filevalidation -Name enableonload -Force      
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\word\security\filevalidation -Name enableonload -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\powerpoint\security\filevalidation -Name enableonload -Force


    #####################################################################
 
    # Unload ntuser.dat        
    IF ($item.SID -in $UnloadedHives.SID) {
        ### Garbage collection and closing of ntuser.dat ###
        [gc]::Collect()
        reg unload HKU\$($Item.SID) | Out-Null
    }
}
