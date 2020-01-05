# VirtualBoxPS
A PowerShell module to manage a Virtual Box environment. This module was developed for an older version of Virtual Box
so I am sure it needs some updating.

TOPIC
    about_VirtualBoxPS

SHORT DESCRIPTION
    These functions are wrappers to the APIs that you can use to manage a virtual machine 
	infrastructure based on the VirtualBox application, which you can download for free from 
	Oracle at http://www.virtualbox.org

LONG DESCRIPTION
    The free virtualization application from Oracle, VirtualBox, offers an application SDK which
    at this point does not include native PowerShell support. This module is an attempt to utilize
    the VirtualBox Web Service to perform common management tasks for virtual machines running
    in the VirtualBox environment.
	
INSTALLATION INSTRUCTIONS
	For this module to work, a simple "intallation" script has been included for Windows to copy 
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
	
	\Psuedo Services\VirtualBox\VirtualBox API Web Service
    
$VBOX GLOBAL VARIABLE
    When you import the module, a global variable is created for the main VirtualBox COM object. 
    This variable, $vbox, is used by many of the support functions. It will be removed when you
    remove the module. If the variable does not exist when you call a function, that requires it,
    it will be recreated in the global scope.
    
    PS S:\> $vbox

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
    
CUSTOM OBJECTS  
    The Get-VirtualBoxVM is often used to get virtual machine objects and most other functions in 
    the module will take pipelined input from this function. Get-VirtualBoxVM writes a custom 
    object to the pipeline with commonly used properties.
    
VERSION
    0.1.8.7
    January 5, 2020
    
SEE ALSO
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
