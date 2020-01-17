# **VirtualBoxPS**
A PowerShell module to manage a Virtual Box environment. This module was developed for an older version of Virtual Box so I am sure it needs some updating.

---

# IMPORTANT
This module is currently in a **very** early development phase. You should not try to use it in a production environment. If you do use it, you do so at your own risk! Please realize that it uses the VirtualBox web service in plain text mode by default. If you provide any information to this web service, it will **NOT** be encrypted. Adding SSL support is possible, but has not been added natively to this module's installer.

---

### **TOPIC**
about_VirtualBoxPS

### **SHORT DESCRIPTION**
These functions are wrappers to the APIs that you can use to manage a virtual machine infrastructure based on the VirtualBox application, which you can download for free from Oracle at [VirtualBox's homepage](http://www.virtualbox.org).

### **LONG DESCRIPTION**
The free virtualization application from Oracle, VirtualBox, offers an application SDK which at this point does not include native PowerShell support. This module is an attempt to utilize the VirtualBox Web Service to perform common management tasks for virtual machines running in the VirtualBox environment.

This module currently requires PowerShell 5.0 (use $PSVersionTable.PSVersion to check your primary version and $Host.Version to check the version of each host type like PS vs. PS_ISE) or higher to run. PS 5.0 introduced classes, which are used somewhat heavily in the module. However, like any ongoing project, all of this is subject to change without notice. The module has currently been tested on Windows 7 and Windows 10, and is being developed on Windows 7.

### **THE GOAL**
This module is being designed to provide as much of the capability as VirtualBox's VBoxManage.exe command line tool or better. The idea is that the module will support greater security, portability, and expandability. That being said, the API being used will be the VirtualBox Web Service. VBoxWebSrv.exe can be launched using certificate based encryption, so it can be accessed using https (still working on a certificate provider for this.) It's also web based so it can be setup to be accessed remotely.

### **CONTRIBUTION**
If you would like to contribute to this project, visit [our thread](https://forums.virtualbox.org/viewtopic.php?f=34&t=54027) at the VirtualBox API Forum. Also, note that there is at least one other project on GitHub pursuing a similar goal, which is also posted in that thread. It is planned to merge [NNVirtualBoxPowerShellMode](https://github.com/ajbrehm/NNVirtualBoxPowerShellModule) and [VirtualBoxPS](#-virtualboxps) in the future after local bugs have been ironed out.
	
### **INSTALLATION INSTRUCTIONS**
For this module to work, a simple "installation" script has been included for Windows to copy all of the required files to the correct locations. Install-VirtualBoxPS.ps1 will need to have vboxweb.wsdl, vboxwebService.wsdl, VirtualBox API Web Service.xml, VirtualBoxPS.psd1, and VirtualBoxPS.psm1 in the same folder. To install, run Install-VirtualBoxPS.ps1 with an elevated PowerShell session. This script will automatically copy the files into the following locations:
	
	$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\vboxweb.wsdl
	$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\vboxwebService.wsdl
	C:\Windows\system32\WindowsPowerShell\v1.0\Modules\VirtualBoxPS\VirtualBoxPS.psd1
	C:\Windows\system32\WindowsPowerShell\v1.0\Modules\VirtualBoxPS\VirtualBoxPS.psm1
	
	
Additionally, a new startup task will be created in Task Scheduler by importing the xml file:
	
	\Pseudo Services\VirtualBox\VirtualBox API Web Service
    
### **_$VBOX_ GLOBAL VARIABLE**
When you import the module, a global variable is created for the main VirtualBox Web Service object. This variable, $vbox, is used by many of the support functions. It will be removed when you remove the module. If the variable does not exist when you call a function, that requires it, it will be recreated in the global scope.
    
    PS C:\> $vbox

	SoapVersion                          : Default
	AllowAutoRedirect                    : False
	CookieContainer                      :
	ClientCertificates                   : {}
	EnableDecompression                  : False
	UserAgent                            : Mozilla/4.0 (compatible; MSIE 6.0; MS Web Services Client Protocol
										   4.0.30319.42000)
	Proxy                                :
	UnsafeAuthenticatedConnectionSharing : False
	Credentials                          :
	UseDefaultCredentials                : False
	ConnectionGroupName                  :
	PreAuthenticate                      : False
	Url                                  : http://localhost:18083/
	RequestEncoding                      :
	Timeout                              : 100000
	Site                                 :
	Container                            :
    
### **CUSTOM OBJECTS**
The Get-VirtualBoxVM is used to get virtual machine objects and most other functions in the module will take pipelined input from this function when everything is completed. Get-VirtualBoxVM writes an array of custom objects to the pipeline with commonly used properties.

### **KNOWN ISSUES**
* Submit-VirtualBoxVMProcess isn't recieving anything from stdout or stderr.

### **WORK AROUNDS**
* Start-VirtualBoxVM -Type Gui -Encrypted will still display a password prompt, even though the VM starts properly.
	>(Workaround) Press cancel when prompted for the disk password, then press cancel when asked what to do next.
* Submit-VirtualBoxVMProcess will crash and abort your VM if VBox tools closes before the command completes.
	>(Workaround) If you are sending a command that will cause this (shutdown commands), use the supplied -Bypass switch.
    
### **VERSION**
	0.2.8.2
	January 17, 2020
    
### **SEE ALSO**
	Get-VirtualBox
	Start-VirtualBoxSession
	Stop-VirtualBoxSession
	Start-VirtualBoxWebSrv
	Stop-VirtualBoxWebSrv
	Restart-VirtualBoxWebSrv
	Update-VirtualBoxWebSrv
	Get-VirtualBoxVM
	Suspend-VirtualBoxVM
	Resume-VirtualBoxVM
	Start-VirtualBoxVM
	Stop-VirtualBoxVM
	New-VirtualBoxVM
	Remove-VirtualBoxVM
	Import-VirtualBoxVM
	Edit-VirtualBoxVM
	Get-VirtualBoxDisk
	New-VirtualBoxDisk
	Submit-VirtualBoxVMProcess
	Submit-VirtualBoxVMPowerShellScript
