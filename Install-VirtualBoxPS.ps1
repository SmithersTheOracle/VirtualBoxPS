Start-Transcript -Append -Confirm:$false -Force -IncludeInvocationHeader -Path "$($env:USERPROFILE)\VirtualBoxPS_install.log"
try{
 Write-Host '[INFO] Installing VirtualBox module'
 Write-Host '[INFO] Creating VirtualBox WebSrv startup task'
 if (((& cmd /c schtasks /query /fo csv /tn `"\Pseudo Services\VirtualBox\VirtualBox API Web Service`") | ConvertFrom-Csv).TaskName -ne '\Pseudo Services\VirtualBox\VirtualBox API Web Service') {
  do {
   $user = Get-Credential -Message "Enter the credentials for the user the VirtualBox Web Server will run-as (the password will not be saved, it's only used to create the task)" -UserName "$($env:USERNAME)"
   foreach ($name in (Get-LocalUser).Name) {
    if ($user.GetNetworkCredential().UserName.ToLower() -eq $name.ToLower()) {$flag = $true}
   }
   if (!$flag) {Write-Host "[Error] User $($user.GetNetworkCredential().UserName) doesn't exist on this computer. Try again." -ForegroundColor Red}
  } until ($flag)
  & cmd /c schtasks /create /ru `"$($env:COMPUTERNAME)\$($user.GetNetworkCredential().UserName)`" /rp `"$($user.GetNetworkCredential().Password)`" /np /tn `"\Pseudo Services\VirtualBox\VirtualBox API Web Service`" /XML `"$((Get-Location).Path)\VirtualBox API Web Service.xml`" | Write-Verbose
 }
 else {Write-Host "[INFO] VirtualBox WebSrv startup task already exists"}
 Write-Host '[INFO] Starting VirtualBox WebSrv'
 if ((Test-Path -LiteralPath "$($env:VBOX_MSI_INSTALL_PATH)VBoxWebSrv.exe") -and !(Get-Process VBoxWebSrv -ErrorAction SilentlyContinue)) {& cmd /c schtasks.exe /run /tn `"\Pseudo Services\VirtualBox\VirtualBox API Web Service`" | Write-Verbose}
 Write-Host '[INFO] Copying files'
 if (!(Test-Path "$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\")) {New-Item -ItemType Directory -Path "$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\" -Force -Confirm:$false | Write-Verbose}
 Copy-Item -Path "$((Get-Location).Path)\*.wsdl" -Destination "$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\" -Force -Confirm:$false | Write-Verbose
 Copy-Item -Path "$((Get-Location).Path)\*.json" -Destination "$($env:VBOX_MSI_INSTALL_PATH)sdk\" -Force -Confirm:$false | Write-Verbose
 $pspath = (($env:PSModulePath).Split(';') | Where-Object {$_ -like 'c:\windows\system32*'})
 if (!(Test-Path "$($pspath)\VirtualBoxPS\")) {New-Item -ItemType Directory -Path "$($pspath)\VirtualBoxPS\" -Force -Confirm:$false | Write-Verbose}
 Copy-Item -Path "$((Get-Location).Path)\*.psm*" -Destination "$($pspath)\VirtualBoxPS\" -Force -Confirm:$false | Write-Verbose
 Copy-Item -Path "$((Get-Location).Path)\*.psd*" -Destination "$($pspath)\VirtualBoxPS\" -Force -Confirm:$false | Write-Verbose
 Write-Host '[SUCCESS] Installation complete' -ForegroundColor Green
 Pause
}
catch{
 Write-Host "[ERROR] Installer encountered a fatal error:`n" -ForegroundColor Red
 Write-Host $Error[0] -ForegroundColor Red
 Pause
}
finally {
 Stop-Transcript
}