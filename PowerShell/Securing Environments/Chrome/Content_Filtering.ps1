# Pulls uBlock Origin from Official Repository in Chrome store and installs to Chrome.
# As this is placed on the forced extension list, user cannot turn off or remove
# https://chrome.google.com/webstore/detail/ublock-origin/cjpalhdlnbpafiamejdnhcphjbkeiagm?hl=en

$chrmext = Test-Path -Path HKLM:\Software\Policies\Google\Chrome\ExtensionInstallForcelist
If ($chrmext -eq $False)
{
	New-Item -Path HKLM:\Software\Policies\Google\Chrome\ExtensionInstallForcelist -Force
}
Set-ItemProperty -Path HKLM:\Software\Policies\Google\Chrome\ExtensionInstallForcelist -Name 1 -Value "cjpalhdlnbpafiamejdnhcphjbkeiagm;https://clients2.google.com/service/update2/crx" -Type String -Force
