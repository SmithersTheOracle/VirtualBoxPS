<##################################################
# 
#          VirtualBoxPS Sample Script
# 
# This is a sample script provided to give a basic
# understanding of how to use the VirtualBoxPS
# PowerShell module. For more detailed information
# about this module use:
#   Get-Module VirtualBoxPS -All
# 
# To get command related help import the module
# and use Get-Help to for each command. Ex:
#   Get-Help Get-VirtualBoxVM -Full
# 
##################################################>

# get credentials to login to the VirtualBox Web Service
$creds = Get-Credential -Message 'Enter credentials' -UserName $env:USERNAME
# import the VirtualBoxPS module
Import-Module VirtualBoxPS -Verbose -Force -NoClobber
# login to the VirtualBox Web Service and create a global session ID
Start-VirtualBoxSession -Credential $creds -Verbose
# get all VM information from the VirtualBox inventory
Get-VirtualBoxVM -Verbose | Select-Object *
# get all virtual disk information from the VirtualBox inventory
Get-VirtualBoxDisks -Verbose | Select-Object *
# pause the script and wait for user to press the Enter key before exiting
Pause