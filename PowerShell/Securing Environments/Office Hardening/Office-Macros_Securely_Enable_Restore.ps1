#####################################################################
# Restore settings enforced by: Office Macros - Secure  worklet.


# Pattern to cycle through SIDs on machine
$PatternSID = 'S-1-\d+-\d+-\d+-\d+\-\d+\-\d+$'
 
# Get Username, SID, and location of ntuser.dat for all users on device
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
    # Office Common
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\Common\Security -Name automationsecurity -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Common\Security -Name vbaoff -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Common\Security -Name macroruntimescanscope -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Common\Security\'Trusted Locations' -Name 'allow user locations' -Force

    # Access
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security -Name blockcontentexecutionfrominternet -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Documents' -Name disabletrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security -Name vbawarnings -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations' -Name allownetworklocations -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations' -Name alllocationsdisabled -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations'\'All Applications'\Location1 -Name path -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations'\'All Applications'\Location1 -Name description -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Access\Security\'Trusted Locations'\'All Applications'\Location1 -Name allowsubfolders -Force
  
    # Excel 
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name excelbypassencryptedmacroscan -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name blockcontentexecutionfrominternet -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name accessvbom -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Documents' -Name disabletrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name vbawarnings -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security -Name xl4macrowarningfollowvba -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations' -Name allownetworklocations -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations' -Name alllocationsdisabled -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations'\'All Applications'\Location1 -Name path -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations'\'All Applications'\Location1 -Name description -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Excel\Security\'Trusted Locations'\'All Applications'\Location1 -Name allowsubfolders -Force
          
    # Outlook
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Outlook\Security -Name level -Force
         
    # Powerpoint
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security -Name powerpointbypassencryptedmacroscan -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security -Name blockcontentexecutionfrominternet -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security -Name accessvbom -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Documents' -Name disabletrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security -Name vbawarnings -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations' -Name allownetworklocations -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations' -Name alllocationsdisabled -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations'\'All Applications'\Location1 -Name path -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations'\'All Applications'\Location1 -Name description -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Powerpoint\Security\'Trusted Locations'\'All Applications'\Location1 -Name allowsubfolders -Force
          
    # MS Project
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations' -Name allownetworklocations -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations' -Name alllocationsdisabled -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security -Name vbawarnings -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations'\'All Applications'\Location1 -Name path -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations'\'All Applications'\Location1 -Name description -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\'MS Project'\Security\'Trusted Locations'\'All Applications'\Location1 -Name allowsubfolders -Force
        
    # MS Publisher
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Publisher\Security -Name automationsecuritypublisher -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Publisher\Security -Name vbawarnings -Force
         
    # Visio
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Application -Name createvbaprojects -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Application -Name loadvbaprojectsfromtext -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations' -Name allownetworklocations -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security -Name blockcontentexecutionfrominternet -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations' -Name alllocationsdisabled -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Documents' -Name disabletrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security -Name vbawarnings -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations'\'All Applications'\Location1 -Name path -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations'\'All Applications'\Location1 -Name description -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Visio\Security\'Trusted Locations'\'All Applications'\Location1 -Name allowsubfolders -Force
        
    # Word
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security -Name wordbypassencryptedmacroscan -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security -Name blockcontentexecutionfrominternet -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security -Name accessvbom -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Documents' -Name disabletrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Documents' -Name disablenetworktrusteddocuments -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security -Name vbawarnings -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations' -Name allownetworklocations -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations' -Name alllocationsdisabled -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations'\'All Applications'\Location1 -Name path -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations'\'All Applications'\Location1 -Name description -Force
    Remove-ItemProperty -Path registry::HKEY_USERS\$($Item.SID)\Software\Policies\Microsoft\Office\16.0\Word\Security\'Trusted Locations'\'All Applications'\Location1 -Name allowsubfolders -Force
      

    #####################################################################
 
    # Unload ntuser.dat        
    IF ($item.SID -in $UnloadedHives.SID) {
        ### Garbage collection and closing of ntuser.dat ###
        [gc]::Collect()
        reg unload HKU\$($Item.SID) | Out-Null
    }
}
