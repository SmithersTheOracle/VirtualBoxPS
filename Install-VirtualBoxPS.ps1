Start-Transcript -Append -Confirm:$false -Force -IncludeInvocationHeader -Path "$($env:USERPROFILE)\PowerVBox_install.log"
try{
 Write-Host '[INFO] Installing VirtualBox module'
 Write-Host '[INFO] Creating VirtualBox WebSrv startup task'
 if (((& cmd /c schtasks /query /fo csv /tn `"\Psuedo Services\VirtualBox\VirtualBox API Web Service`") | ConvertFrom-Csv).TaskName -ne '\Psuedo Services\VirtualBox\VirtualBox API Web Service') {
  & cmd /c schtasks /create /ru `"SYSTEM`" /rp /tn `"\Psuedo Services\VirtualBox\VirtualBox API Web Service`" /XML `"$((Get-Location).Path)\VirtualBox API Web Service.xml`" | Write-Verbose
 }
 else {Write-Host "[INFO] VirtualBox WebSrv startup task already exists"}
 Write-Host '[INFO] Starting VirtualBox WebSrv'
 if ((Test-Path -LiteralPath "$($env:VBOX_MSI_INSTALL_PATH)VBoxWebSrv.exe") -and !(Get-Process VBoxWebSrv)) {& cmd /c schtasks.exe /run /tn `"\Psuedo Services\VirtualBox\VirtualBox API Web Service`" | Write-Verbose}
 Write-Host '[INFO] Copying files'
 Copy-Item -Path "$((Get-Location).Path)\*.wsdl" -Destination "$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\" -Force -Confirm:$false | Write-Verbose
 foreach ($pspath in (($env:PSModulePath).Split(';'))) {
  if (!(Test-Path "$($pspath)\Oracle.PowerVBox\1.0\")) {New-Item -ItemType Directory -Path "$($pspath)\Oracle.PowerVBox\1.0\" -Force -Confirm:$false | Write-Verbose}
  Copy-Item -Path "$((Get-Location).Path)\*.psm*" -Destination "$($pspath)\Oracle.PowerVBox\1.0\" -Force -Confirm:$false | Write-Verbose
  Copy-Item -Path "$((Get-Location).Path)\*.psd*" -Destination "$($pspath)\Oracle.PowerVBox\1.0\" -Force -Confirm:$false | Write-Verbose
 }
 Write-Host '[SUCCESS] Installation complete' -ForegroundColor Green
 Pause
}
catch{
 Write-Host "[ERROR] Installer encountered a fatal error" -ForegroundColor Red
 $Error[0]
}
finally {
 Stop-Transcript
}