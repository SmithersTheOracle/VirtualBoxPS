<##################################################
# 
#          VirtualBoxPS Sample Script
# 
# This is a sample script provided to give a basic
# understanding of how to use the VirtualBoxPS
# PowerShell module. For more detailed information
# about this module use:
#   Get-Module -ListAvailable -Name VirtualBoxPS -All
# 
# To get command related help import the module
# and use Get-Help to for each command. Ex:
#   Get-Help Get-VirtualBoxVM -Full
# 
# **NOTE: IF YOU HAVE A VM WITH A NAME MATCHING 
# "othertest" OR VIRTUAL DISKS WITH NAMES MATCHING 
# "testdisk" THEY WILL BE DELETED!**
# 
##################################################>

# import the VirtualBoxPS module
Import-Module VirtualBoxPS -ArgumentList COM -Verbose -Force -NoClobber
# create a new VM named OtherTest64 with an OsTypeId of Other_64 and CpuCount of 2
New-VirtualBoxVM -Name OtherTest64 -OsTypeId Other_64 -CpuCount 2 -Verbose
# enable the VRDE server for the OtherTest64 VM
Enable-VirtualBoxVMVRDEServer -Name OtherTest64 -Verbose | Out-Null
# edit the VRDE server for the OtherTest64 VM to AuthType of Null and TcpPort of 43389
Edit-VirtualBoxVMVRDEServer -Name OtherTest64 -AuthType Null -TcpPort 43389 -Verbose | Out-Null
# edit the OtherTest64 VM to Description of Test description
Edit-VirtualBoxVM -Name OtherTest64 -Description "Test description" -Verbose
# find the OtherTest64 VM and pipeline the machine object to Edit-VirtualBoxVM to edit the CpuCount to 4
Get-VirtualBoxVM othertest6 | Edit-VirtualBoxVM -CpuCount 4 -Verbose
# edit the OtherTest64 VM to its VirtualBox default icon
Edit-VirtualBoxVM -Name OtherTest64 -Icon '' -Verbose
# perform a WhatIf removal of the OtherTest64 VM
Remove-VirtualBoxVM -Name OtherTest64 -Whatif -Verbose
# remove the OtherTest64 VM from the VirtualBox inventory and silence all confirmation prompts
Remove-VirtualBoxVM -Name OtherTest64 -Confirm:$false -Verbose
# import the OtherTest64 VM back into the VirtualBox inventory
Import-VirtualBoxVM -Name OtherTest64 -Location "$($env:HOMEDRIVE)$($env:HOMEPATH)\VirtualBox VMs\OtherTest64" -Verbose
# create a new 4MB VMDK virtual disk named TestDisk on the current user's desktop and display a progress bar
New-VirtualBoxDisk -AccessMode ReadWrite -Format VMDK -Location "$($env:HOMEDRIVE)$($env:HOMEPATH)\Desktop" -LogicalSize 4194304 -Name TestDisk -VariantType Standard -ProgressBar -Verbose
# mount the virtual disk named TestDisk to any machines matching the name othertest to the IDE controller primary master port
Mount-VirtualBoxDisk -Name testdisk -MachineName othertest -Controller IDE -ControllerPort 0 -ControllerSlot 0 -Verbose
# create a new 4MB VMDK virtual disk named TestDisk2 on the current user's desktop and display a progress bar
New-VirtualBoxDisk -AccessMode ReadWrite -Format VMDK -Location "$($env:HOMEDRIVE)$($env:HOMEPATH)\Desktop" -LogicalSize 4194304 -Name TestDisk2 -VariantType Standard -ProgressBar -Verbose
# mount the virtual disk named TestDisk2 to any machines matching the name othertest to the IDE controller primary slave port
Mount-VirtualBoxDisk -Name testdisk2.vmdk -MachineName othertest -Controller IDE -ControllerPort 0 -ControllerSlot 1 -Verbose
# dismount the virtual disk named TestDisk2 from any machines matching the name othertest
Dismount-VirtualBoxDisk -Name testdisk2.vmdk -MachineName othertest -Confirm:$false -Verbose
# remount the virtual disk named TestDisk2 to any machines matching the name othertest to the IDE controller primary slave port
Mount-VirtualBoxDisk -Name testdisk2.vmdk -MachineName othertest -Controller IDE -ControllerPort 0 -ControllerSlot 1 -Verbose
# get all virtual disks attached to the OtherTest64 VM and dismount them only from that machine
(Get-VirtualBoxVM -Name OtherTest64 -Verbose).IMediumAttachments.IMedium | Dismount-VirtualBoxDisk -MachineName OtherTest64 -Confirm:$false -Verbose
# remove TestDisk.vmdk from the VirtualBox inventory
Remove-VirtualBoxDisk -Name testdisk.vmdk -Verbose
# reimport TestDisk.vmdk to the VirtualBox inventory
Import-VirtualBoxDisk -FileName "$($env:HOMEDRIVE)$($env:HOMEPATH)\Desktop\TestDisk.vmdk" -AccessMode ReadWrite -Verbose
# remove TestDisk.vmdk from the VirtualBox inventory, delete it from the host machine, and display a progress bar
Remove-VirtualBoxDisk -Name testdisk.vmdk -DeleteFromHost -Confirm:$false -ProgressBar -Verbose
# recreate the TestDisk.vmdk
New-VirtualBoxDisk -AccessMode ReadWrite -Format VMDK -Location "$($env:HOMEDRIVE)$($env:HOMEPATH)\Desktop" -LogicalSize 4194304 -Name TestDisk -VariantType Standard -ProgressBar -Verbose
# remount TestDisk.vmdk to the OtherTest64 VM to the IDE controller primary master port
Mount-VirtualBoxDisk -Name testdisk.vmdk -MachineName othertest -Controller IDE -ControllerPort 0 -ControllerSlot 0 -Verbose
# remount TestDisk2.vmdk to the OtherTest64 VM to the IDE controller primary slave port
Mount-VirtualBoxDisk -Name testdisk2.vmdk -MachineName othertest -Controller IDE -ControllerPort 0 -ControllerSlot 1 -Verbose
# remove the OtherTest64 VM from the VirtualBox inventory
Remove-VirtualBoxVM -Name OtherTest64 -ProgressBar -Confirm:$false -Verbose
# reimport the OtherTest64 VM to the VirtualBox inventory
Import-VirtualBoxVM -Name OtherTest64 -Location "$($env:HOMEDRIVE)$($env:HOMEPATH)\VirtualBox VMs\OtherTest64" -Verbose
# remove the OtherTest64 VM from the VirtualBox inventory and detach all its media
Remove-VirtualBoxVM -Name OtherTest64 -DetachAllReturnNone -ProgressBar -Confirm:$false -Verbose
# reimport the OtherTest64 VM to the VirtualBox inventory
New-VirtualBoxVM -Name OtherTest64 -OsTypeId Other_64 -CpuCount 2 -Verbose
# remount TestDisk.vmdk to the OtherTest64 VM to the IDE controller primary master port
Mount-VirtualBoxDisk -Name testdisk.vmdk -MachineName othertest -Controller IDE -ControllerPort 0 -ControllerSlot 0 -Verbose
# remount TestDisk2.vmdk to the OtherTest64 VM to the IDE controller primary slave port
Mount-VirtualBoxDisk -Name testdisk2.vmdk -MachineName othertest -Controller IDE -ControllerPort 0 -ControllerSlot 1 -Verbose
# delete the OtherTest64 VM and detach all its media and delete the disks from the host machine
Remove-VirtualBoxVM -Name OtherTest64 -DetachAllReturnHardDisksOnly -ProgressBar -Confirm:$false -Verbose
# recreate the OtherTest64 VM
New-VirtualBoxVM -Name OtherTest64 -OsTypeId Other_64 -CpuCount 2 -Verbose
# recreate TestDisk.vmdk
New-VirtualBoxDisk -AccessMode ReadWrite -Format VMDK -Location "$($env:HOMEDRIVE)$($env:HOMEPATH)\Desktop" -LogicalSize 4194304 -Name TestDisk -VariantType Standard -ProgressBar -Verbose
# remount TestDisk.vmdk to the OtherTest64 VM to the IDE controller primary master port
Mount-VirtualBoxDisk -Name testdisk.vmdk -MachineName othertest -Controller IDE -ControllerPort 0 -ControllerSlot 0 -Verbose
# recreate TestDisk2.vmdk
New-VirtualBoxDisk -AccessMode ReadWrite -Format VMDK -Location "$($env:HOMEDRIVE)$($env:HOMEPATH)\Desktop" -LogicalSize 4194304 -Name TestDisk2 -VariantType Standard -ProgressBar -Verbose
# remount TestDisk2.vmdk to the OtherTest64 VM to the IDE controller primary slave port
Mount-VirtualBoxDisk -Name testdisk2.vmdk -MachineName othertest -Controller IDE -ControllerPort 0 -ControllerSlot 1 -Verbose
# delete the OtherTest64 VM and detach all its media and delete them from the host machine (will not currently affect virtual optical media)
Remove-VirtualBoxVM -Name OtherTest64 -Full -ProgressBar -Confirm:$false -Verbose
# get all VM information from the VirtualBox inventory
Get-VirtualBoxVM -Verbose | Select-Object *
# get all virtual disk information from the VirtualBox inventory
Get-VirtualBoxDisk -Verbose | Select-Object *
# pause the script and wait for user to press the Enter key before exiting
Pause