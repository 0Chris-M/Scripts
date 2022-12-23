$cmnsec = Test-Path -Path HKLM:\Software\Policies\Google\Chrome\ExtensionInstallForcelist
If ($cmnsec -eq $False)
{
	New-Item -Path HKLM:\Software\Policies\Google\Chrome\ExtensionInstallForcelist -Force
}
Set-ItemProperty -Path HKLM:\Software\Policies\Google\Chrome\ExtensionInstallForcelist -Name 1 -Value "cjpalhdlnbpafiamejdnhcphjbkeiagm;https://clients2.google.com/service/update2/crx" -Type String -Force