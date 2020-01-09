# [![Logo](logo_sm.png)] **VirtualBoxPS**
A PowerShell module to manage a Virtual Box environment. This module was developed for an older version of Virtual Box
so I am sure it needs some updating.

---

# [![!](exclaim_sm.png)] IMPORTANT
This module is currently in a **very** early development phase. You should not try to use it 
in a production environment. If you do use it, you do so at your own risk! Please realize 
that it uses the VirtualBox web proxy in plain text mode by default. If you provide any 
information to this web service, it will **NOT** be encrypted. Adding SSL support is possible, 
but has not been added natively to this module's installer.

---

### **TOPIC**
about_VirtualBoxPS

### **SHORT DESCRIPTION**
These functions are wrappers to the APIs that you can use to manage a virtual machine 
infrastructure based on the VirtualBox application, which you can download for free from 
Oracle at [VirtualBox's homepage.](http://www.virtualbox.org)

### **LONG DESCRIPTION**
The free virtualization application from Oracle, VirtualBox, offers an application SDK which
at this point does not include native PowerShell support. This module is an attempt to utilize
the VirtualBox Web Service to perform common management tasks for virtual machines running
in the VirtualBox environment.
	
### **INSTALLATION INSTRUCTIONS**
For this module to work, a simple "installation" script has been included for Windows to copy 
all of the required files to the correct locations. Install-VirtualBoxPS.ps1 will need to have 
vboxweb.wsdl, vboxwebService.wsdl, VirtualBox API Web Service.xml, VirtualBoxPS.psd1, and
VirtualBoxPS.psm1 in the same folder. To install, run Install-VirtualBoxPS.ps1 with an
elevated PowerShell session. This script will automatically copy the files into the following
locations:
	
	$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\vboxweb.wsdl
	$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\vboxwebService.wsdl
	$($env:PSModulePath)\Oracle.PowerVBox\1.0\VirtualBoxPS.psd1
	$($env:PSModulePath)\Oracle.PowerVBox\1.0\VirtualBoxPS.psm1
	
	
Additionally, a new startup task will be created in Task Scheduler by importing the xml file:
	
	\Pseudo Services\VirtualBox\VirtualBox API Web Service
    
### **_$VBOX_ GLOBAL VARIABLE**
When you import the module, a global variable is created for the main VirtualBox Web Service object. 
This variable, $vbox, is used by many of the support functions. It will be removed when you
remove the module. If the variable does not exist when you call a function, that requires it,
it will be recreated in the global scope.
    
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
The Get-VirtualBoxVM is often used to get virtual machine objects and most other functions in 
the module will take pipelined input from this function. Get-VirtualBoxVM writes a custom 
object to the pipeline with commonly used properties.
    
### **VERSION**
	0.1.12.10
	January 8, 2020
    
### **SEE ALSO**
	Get-VirtualBox
	Start-VirtualBoxSession
	Stop-VirtualBoxSession
	Start-VirtualBoxWebSrv
	Stop-VirtualBoxWebSrv
	Restart-VirtualBoxWebSrv
	Update-VirtualBoxWebSrv
	Get-VirtualBoxProcess
	Get-VirtualBoxVM
	Suspend-VirtualBoxVM
	Start-VirtualBoxVM
	Stop-VirtualBoxVM
	Get-VirtualBoxDisks
