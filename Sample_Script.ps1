$creds = Get-Credential -Message 'Enter credentials' -UserName $env:USERNAME
Import-Module VirtualBoxPS -Verbose -Force -NoClobber
Start-VirtualBoxSession -Credential $creds -Verbose
Get-VirtualBoxVM -Verbose
Get-VirtualBoxDisks -Verbose
Pause