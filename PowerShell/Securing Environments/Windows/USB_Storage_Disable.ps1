# Disables the USB port for USB Storage only, and Deny's R/W/X capabilities. This Policy applies to the Local Machine. 
# Peripheral Devices (mouse, KB, Printer, etc.) are unaffected.
# Requires system reboot for changes to take effect
$remusb = Test-Path -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices"
If ($remusb -eq $False)
{
	New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices" -Force
}
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices" -Name "Deny_All" -Value 1 -Type DWord -Force
$remusb = Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR"
If ($remusb -eq $False)
{
	New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR" -Force
}
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR" -Name "Start" -Value 4 -Type DWord -Force
