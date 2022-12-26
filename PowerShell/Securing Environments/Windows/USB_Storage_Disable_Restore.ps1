# Re-Enables the USB port for USB Storage, and restores R/W/X capabilities
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR" -Name "Start" -Value 3 -Type DWord -Force
Remove-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices" -Force
