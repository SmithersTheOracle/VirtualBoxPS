# Requires -version 2.0
<#
TODO:
Finish Integrating Progress bar
Create new VM
Create a new Disk
#>
<#
****************************************************************
* DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
* THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
* YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
* DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
****************************************************************
#>
#########################################################################################
# Classe Definitions
class VirtualBoxVM {
    [ValidateNotNullOrEmpty()]
    [string]$Name
    [ValidateNotNullOrEmpty()]
    [string]$Id
    [string]$Uuid
    [string]$Description
    [string]$MemoryMB
    [string]$State
    [bool]$Running
    [string]$Info
    [string]$GuestOS
    [string]$ISession
}
Update-TypeData -TypeName VirtualBoxVM -DefaultDisplayPropertySet @("UUID","Name","MemoryMB","Description","State","GuestOS") -Force
class VirtualBoxVHD {
    [string]$Name
    [string]$Description
    [string]$Format
    [string]$Size
    [string]$LogicalSize
    [string[]]$VMIds
    [string[]]$VMNames
    [string]$State
    [string[]]$Variant
    [string]$Location
    [string]$HostDrive
    [string]$MediumFormat
    [string]$Type
    [string]$Parent
    [string[]]$Children
    [string]$Id
    [string]$ReadOnly
    [string]$AutoReset
    [string]$LastAccessError
}
Update-TypeData -TypeName VirtualBoxVHD -DefaultDisplayPropertySet @("Name","Description","Format","Size","LogicalSize","VMIds","VMNames") -Force
class VirtualBoxWebSrvTask {
    [string]$Name
    [string]$Path
    [string]$Status
}
Update-TypeData -TypeName VirtualBoxWebSrvTask -DefaultDisplayPropertySet @("Name","Path","Status") -Force
class VirtualBoxError {
    [string]Call ($ErrInput) {
        if ($ErrInput){return $ErrInput.ToString().Substring($ErrInput.ToString().IndexOf('"')).Split('"')[1]}
        else {return $null}
    }
    [string]Code ($ErrInput) {
        if ($ErrInput){return $ErrInput.ToString().Substring($ErrInput.ToString().IndexOf('rc=')+3).Remove(10)}
        else {return $null}
    }
    [string]Description ($ErrInput) {
        if ($ErrInput){return $ErrInput.ToString().Substring($ErrInput.ToString().IndexOf('rc=')+14).Split('(')[0].TrimEnd(' ')}
        else {return $null}
    }
}
#########################################################################################
# Variable Declarations
$authtype = "VBoxAuth"
$vboxwebsrvtask = New-Object VirtualBoxWebSrvTask
$vboxerror = New-Object VirtualBoxError
#########################################################################################
# Includes
# N/A
#########################################################################################
# Function Definitions
Function Get-VirtualBox {
<#
.SYNOPSIS
Get the VirtualBox Web Service.
.DESCRIPTION
Create a PowerShell object for the VirtualBox Web Service object. This command is run when the VirtualBoxPS module is loaded.
.EXAMPLE
PS C:\> $vbox = Get-VirtualBox
Creates a $vbox variable to referece the VirtualBox Web Service
.NOTES
NAME        :  Get-VirtualBox
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Start-VirtualBoxSession
.INPUTS
None
.OUTPUTS
$global:vbox
#>
[cmdletbinding()]
Param() # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
} # Begin
Process {
 # create vbox app
 Write-Verbose 'Creating the VirtualBox Web Service object ($global:vbox)'
 #$global:vbox = New-Object -ComObject "VirtualBox.VirtualBox"
 $global:vbox = New-WebServiceProxy -Uri "$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\vboxwebService.wsdl" -Namespace "VirtualBox" -Class "VirtualBox"
 # write variable to the pipeline
 $global:vbox
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Start-VirtualBoxSession {
<#
.SYNOPSIS
Starts a VirtualBox Web Service session and populates the $global:ivbox managed object reference.
.DESCRIPTION
Create a PowerShell object reference to the VirtualBox Web Service managed object.
.EXAMPLE
PS C:\> Start-VirtualBoxSession -Protocol "http" -Domain "localhost" -Port "18083" -Credential $Credential
Populates the $global:ivbox variable to referece the VirtualBox Web Service managed object
.NOTES
NAME        :  Start-VirtualBoxSession
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBox
Stop-VirtualBoxSession
.INPUTS
Pos0: string       : <"http">
Pos1: string       : <"localhost">
Pos2: string       : <"18083">
Pos3: pscredential :
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(Position=0)]
[ValidateSet("http","https")]
  [string]$Protocol = "http",
# localhost ONLY for now since we haven't enabled https
[Parameter(Position=1)]
  [string]$Domain = "localhost",
[Parameter(Position=2)]
  [string]$Port = "18083",
[Parameter(Position=3)]
  [pscredential]$Credential
) # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
 # get global vbox variable or create it if it doesn't exist create it
 if (-Not $global:vbox) {$global:vbox = Get-VirtualBox}
 # refresh vboxwebsrv variable
 if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
 # start the websrvtask if it's not running
 if ($global:vboxwebsrvtask.Status -and $global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
 # set the target web service url
 $global:vbox.Url = "$($Protocol)://$($Domain):$($Port)"
 # if a session already exists, stop it
 if ($global:ivbox) {Stop-VirtualBoxSession}
} # Begin
Process {
 try {
  # login to web service
  Write-Verbose 'Creating the VirtualBox Web Service session ($global:ivbox)'
  $global:ivbox = $global:vbox.IWebsessionManager_logon($Credential.GetNetworkCredential().UserName,$Credential.GetNetworkCredential().Password)
 }
 catch {
  Write-Verbose '$_.Exception'
  Write-Host $_.Exception -ForegroundColor Red
 }
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Stop-VirtualBoxSession {
<#
.SYNOPSIS
Closes the VirtualBox Web Service session.
.DESCRIPTION
Instruct the VirtualBox Web Service to close the current managed object session.
.EXAMPLE
PS C:\> Stop-VirtualBoxSession
.NOTES
NAME        :  Stop-VirtualBoxSession
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBox
Start-VirtualBoxSession
.INPUTS
None
.OUTPUTS
None
#>
[cmdletbinding(DefaultParameterSetName="UserPass")]
Param() # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
 # get global vbox variable or create it if it doesn't exist create it
 if (-Not $global:vbox) {$global:vbox = Get-VirtualBox}
 # refresh vboxwebsrv variable
 if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
 # start the websrvtask if it's not running
 if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
} # Begin
Process {
 # login to web service
 Write-Verbose 'Creating the VirtualBox Web Service session ($global:ivbox)'
 if ($global:ivbox) {
  try {
   # tell vboxwebsrv to end the current session
   $global:vbox.IWebsessionManager_logoff($global:ivbox)
   $global:ivbox = $null
  } # end try
  catch {
   Write-Verbose '$_.Exception'
   Write-Host $_.Exception -ForegroundColor Red
  }
 }
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Start-VirtualBoxWebSrv {
<#
.SYNOPSIS
Starts the VirtualBox Web Service.
.DESCRIPTION
Starts the VirtualBox Web Service using schtask.exe.
.EXAMPLE
PS C:\> Start-VirtualBoxWebSrv
Starts the VirtualBox Web Service if it isn't already running
.NOTES
NAME        :  Start-VirtualBoxWebSrv
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Stop-VirtualBoxWebSrv
Restart-VirtualBoxWebSrv
Update-VirtualBoxWebSrv
.INPUTS
None
.OUTPUTS
None
#>
[cmdletbinding()]
Param() # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
} # Begin
Process {
 try {
  # refresh the vboxwebsrv scheduled task
  Write-Verbose 'Running Update-VirtualBoxWebSrv cmdlet'
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  Write-Verbose "$($global:vboxwebsrvtask.Name) status: $($global:vboxwebsrvtask.Status)"
  if ($global:vboxwebsrvtask.Status -and $global:vboxwebsrvtask.Status -ne 'Running') {
   # start the web service task
   Write-Verbose "Starting the VirtualBox Web Service ($($global:vboxwebsrvtask.Name))"
   & cmd /c schtasks.exe /run /tn `"$($global:vboxwebsrvtask.Path)$($global:vboxwebsrvtask.Name)`"
  }
  else {
   # return a message
   return "The VBoxWebSrv task is already running"
  }
 }
 catch {
  Write-Verbose '$_.Exception'
  Write-Host $_.Exception -ForegroundColor Red
 }
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Stop-VirtualBoxWebSrv {
<#
.SYNOPSIS
Stops the VirtualBox Web Service.
.DESCRIPTION
Stops the VirtualBox Web Service using schtask.exe.
.EXAMPLE
PS C:\> Stop-VirtualBoxWebSrv
Stops the VirtualBox Web Service it is running
.NOTES
NAME        :  Stop-VirtualBoxWebSrv
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Start-VirtualBoxWebSrv
Restart-VirtualBoxWebSrv
Update-VirtualBoxWebSrv
.INPUTS
None
.OUTPUTS
None
#>
[cmdletbinding(DefaultParameterSetName="UserPass")]
Param() # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
} # Begin
Process {
 # login to web service
 Write-Verbose 'Ending the VirtualBox Web Service'
 try {
  # tell vboxwebsrv to end the current session
  & cmd /c schtasks.exe /end /tn `"$($global:vboxwebsrvtask.Path)$($global:vboxwebsrvtask.Name)`"
 } # end try
 catch {
  Write-Verbose '$_.Exception'
  Write-Host $_.Exception -ForegroundColor Red
 } # end catch
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Restart-VirtualBoxWebSrv {
<#
.SYNOPSIS
Restarts the VirtualBox Web Service.
.DESCRIPTION
Stops then starts the VirtualBox Web Service using schtask.exe.
.EXAMPLE
PS C:\> Restart-VirtualBoxWebSrv
Restarts the VirtualBox Web Service
.NOTES
NAME        :  Restart-VirtualBoxWebSrv
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Start-VirtualBoxWebSrv
Stop-VirtualBoxWebSrv
Update-VirtualBoxWebSrv
.INPUTS
None
.OUTPUTS
None
#>
[cmdletbinding()]
Param() # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
} # Begin
Process {
 # restart the web service task
 Stop-VirtualBoxWebSrv
 Start-VirtualBoxWebSrv
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Update-VirtualBoxWebSrv {
<#
.SYNOPSIS
Gets the updated status of the VirtualBox Web Service.
.DESCRIPTION
Gets the updated status of the VirtualBox Web Service using schtask.exe.
.EXAMPLE
PS C:\> if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
Returns the updated status of the VirtualBox Web Service
.NOTES
NAME        :  Update-VirtualBoxWebSrv
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Start-VirtualBoxWebSrv
Stop-VirtualBoxWebSrv
Restart-VirtualBoxWebSrv
.INPUTS
None
.OUTPUTS
[VirtualBoxWebSrvTask]$vboxwebsrvtask
#>
[cmdletbinding()]
Param() # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
} # Begin
Process {
 # refresh the web service task information
 try {
  Write-Verbose 'Updating $global:vboxwebsrvtask'
  $tempjoin = $()
  $tempobj = (& cmd /c schtasks.exe /query /fo csv | ConvertFrom-Csv | Where-Object {$_.TaskName -match 'VirtualBox API Web Service'}).TaskName.Split("\")
  $vboxwebsrvtask = New-Object VirtualBoxWebSrvTask
  for ($a=0;$a-lt$tempobj.Count;$a++) {
   if ($a -lt $tempobj.Count-1) {
    $tempjoin += $tempobj[$a].Insert($tempobj[$a].Length,'\')
   }
   else {
    $vboxwebsrvtask.Name = $tempobj[$a]
    $vboxwebsrvtask.Path = [string]::Join('\',$tempjoin)
   }
  }
  $vboxwebsrvtask.Status = (& cmd /c schtasks.exe /query /fo csv | ConvertFrom-Csv | Where-Object {$_.TaskName -match 'VirtualBox API Web Service'}).Status
 } # end try
 catch {
  Write-Verbose '$_.Exception'
  Write-Host $_.Exception -ForegroundColor Red
 } # end catch
 if (!$vboxwebsrvtask) {throw 'Failed to update $vboxwebsrvtask'}
 return $vboxwebsrvtask
} # Process
End {
 Write-Verbose $vboxwebsrvtask
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Get-VirtualBoxVM {
<#
.SYNOPSIS
Get a VirtualBox virtual machine
.DESCRIPTION
Retrieve any or all VirtualBox machines by name, by state or all. The default usage, without any parameters is to display all running virtual machines. Use -IncludeRaw to add the native COM object for the virtual machine.
.PARAMETER Name
The name of a virtual machine. IMPORTANT: Names are case sensitive.
.PARAMETER All
Return all virtual machines regardless of state. Valid values are:
"Stopped","Running","Saved","Teleported","Aborted","Paused","Stuck","Snapshotting",
"Starting","Stopping","Restoring","TeleportingPausedVM","TeleportingIn","FaultTolerantSync",
"DeletingSnapshotOnline","DeletingSnapshot", and "SettingUp"
.PARAMETER State
Return virtual machines based on their state.
.PARAMETER IncludeRaw
Include the raw or native COM object for the virtual machine.
.EXAMPLE
PS C:\> Get-VirtualBoxVM
ID          : 96c58d09-37be-46b1-9f4b-d37ea6da4005
Name        : Win2008 R2 Standard
MemoryMB    : 1500
Description : Windows 2008 R2 Standard DC jdhlab.local
State       : Running
OS          : Windows2008_64

ID          : ed29417c-869a-45bf-bbf3-79a407ade630
Name        : CoreDC01
MemoryMB    : 512
Description :
State       : Running
OS          : Windows2008_64

ID          : 2dd7f99a-d209-4b1c-ad79-2fa34e2c229a
Name        : Ubuntu
MemoryMB    : 1024
Description : v11.04 Natty Narwhal
State       : Running
OS          : Ubuntu_64

Return all running virtual machines
.EXAMPLE
PS C:\> Get-VirtualBoxVM -Name CoreDC01
ID          : ed29417c-869a-45bf-bbf3-79a407ade630
Name        : CoreDC01
MemoryMB    : 512
Description :
State       : Running
OS          : Windows2008_64

Retrieve a machine by name. Names are case sensitive
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Saved
ID          : 2dd7f99a-d209-4b1c-ad79-2fa34e2c229a
Name        : Ubuntu
MemoryMB    : 1024
Description : v11.04 Natty Narwhal
State       : Saved
OS          : Ubuntu_64

Get suspended virtual machines
.NOTES
NAME        :  Get-VirtualBoxVM
VERSION     :  0.9
LAST UPDATED:  6/13/2011
AUTHOR      :  Jeffery Hicks
.LINK
Stop-VirtualBoxVM
Start-VirtualBoxVM
Suspend-VirtualBoxVM
.INPUTS
Strings for virtual machine names
.OUTPUTS
Custom Object
#>
[cmdletbinding(DefaultParameterSetName="All")]
Param(
[Parameter(Position=0)]
  [string[]]$Name,
  [switch]$SkipCheck,
[Parameter(ParameterSetName="All")]
  [switch]$All,
[Parameter(ParameterSetName="All")]
[ValidateSet("Stopped","Running","Saved","Teleported","Aborted",
   "Paused","Stuck","Snapshotting","Starting","Stopping",
   "Restoring","TeleportingPausedVM","TeleportingIn","FaultTolerantSync",
   "DeletingSnapshotOnline","DeletingSnapshot","SettingUp")]
  [string]$State
) # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
 # check global vbox variable and create it if it doesn't exist
 if (-Not $global:vbox) {$global:vbox = Get-VirtualBox}
 # refresh vboxwebsrv variable
 if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
 # start the websrvtask if it's not running
 if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
 if (-Not $global:ivbox) {Start-VirtualBoxSession}
 if (!$Name) {$All = $true}
} # Begin
Process {
 Write-Verbose "Getting virtual machine inventory"
 # initialize array object to hold virtual machine values
 $vminventory = @()
 # get virtual machine inventory
 foreach ($vmid in ($global:vbox.IVirtualBox_getMachines($global:ivbox))) {
   $tempobj = New-Object VirtualBoxVM
   $tempobj.Name = $global:vbox.IMachine_getName($vmid)
   $tempobj.Description = $global:vbox.IMachine_getDescription($vmid)
   $tempobj.State = $global:vbox.IMachine_getState($vmid)
   $tempobj.GuestOS = $global:vbox.IMachine_getOSTypeId($vmid)
   $tempobj.MemoryMb = $global:vbox.IMachine_getMemorySize($vmid)
   $tempobj.Id = $vmid
   $tempobj.Uuid = $global:vbox.IMachine_getId($vmid)
   # decode state
   Switch ($tempobj.State) {
    1 {$tempobj.State = "Stopped"}
    2 {$tempobj.State = "Saved"}
    3 {$tempobj.State = "Teleported"}
    4 {$tempobj.State = "Aborted"}
    5 {$tempobj.State = "Running"}
    6 {$tempobj.State = "Paused"}
    7 {$tempobj.State = "Stuck"}
    8 {$tempobj.State = "Snapshotting"}
    9 {$tempobj.State = "Starting"}
    10 {$tempobj.State = "Stopping"}
    11 {$tempobj.State = "Restoring"}
    12 {$tempobj.State = "TeleportingPausedVM"}
    13 {$tempobj.State = "TeleportingIn"}
    14 {$tempobj.State = "FaultTolerantSync"}
    15 {$tempobj.State = "DeletingSnapshotOnline"}
    16 {$tempobj.State = "DeletingSnapshot"}
    17 {$tempobj.State = "SettingUp"}
    Default {$tempobj.State = $tempobj.State}
   }
   $vminventory += $tempobj
 } # end foreach loop inventory
 # filter virtual machines
 if ($Name -and $Name -ne "*") {
  Write-Verbose "Filtering virtual machines by name: $Name"
  foreach ($vm in $vminventory) {
   Write-Verbose "Matching $($vm.Name) to $($Name)"
   if ($vm.Name -match $Name) {
    $obj += $vm
   }
  }
 } # end if $Name and not *
 if ($All -or $Name -eq "*") {
  Write-Verbose "Filtering all virtual machines"
  $obj = $vminventory
 } # end if $All or $Name *
 if ($State) {
  Write-Verbose "Filtering virtual machines with a state of $State"
  $obj=$vminventory | where {$_.State -eq $State}
 } # end if $State
 Write-Verbose "Found $(($obj | Measure-Object).count) virtual machines"
 if ($obj) {
  # write virtual machines object to the pipeline
  $obj
 } # end if $obj
 else {
  Write-Host "[Warning] No matching virtual machines found" -ForegroundColor DarkYellow
 } # end else
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Suspend-VirtualBoxVM {
<#
.SYNOPSIS
Suspend a virtual machine
.DESCRIPTION
Suspends or saves the state of a running virtual machine.
.PARAMETER Name
The Name of a running virtual machine.
.PARAMETER UUID
The UUID of a running virtual machine.
.EXAMPLE
PS C:\> Get-VirtualBoxVM | Suspend-VirtualBoxVM
Suspend all running virtual machines
.NOTES
NAME        :  Suspend-VirtualBoxVM
VERSION     :  0.1b.garbage.doesntwork
LAST UPDATED:  1/5/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxVM
Start-VirtualBoxVM
Stop-VirtualBoxVM
.INPUTS
String[]    :  
.OUTPUTS
None
#>
[cmdletbinding(DefaultParametersetName='Name')]
Param(
[Parameter(ParameterSetName="Name",
Mandatory=$true,
HelpMessage="Enter a virtual box machine Name",
ValueFromPipeline=$true,
ValueFromPipelineByPropertyName=$true)]
[Parameter(Position=0)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
  [switch]$SkipCheck,
[Parameter(ParameterSetName="Uuid",
HelpMessage="Enter a virtual box machine UUID",
ValueFromPipelineByPropertyName=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Uuid
) # Param
Begin {
 Write-Verbose "Ending $($myinvocation.mycommand)"
 #get global vbox variable or create it if it doesn't exist create it
 if (-Not $global:vbox) {$global:vbox = Get-VirtualBox}
 # refresh vboxwebsrv variable
 if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
 # start the websrvtask if it's not running
 if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
} # Begin
Process {
 try {
  if ($Name) {
   foreach ($item in $Name) {
    # get the virtual machine by Name
    $vmachines = Get-VirtualBoxVM $item -SkipCheck
    if ($vmachines) {
     foreach ($vmachine in $vmachines) {
      Write-Verbose "Suspending $($vmachine.Name)"
      # create a session object
      Write-Verbose "Creating a session object"
      $isession = $global:vbox.IWebsessionManager_getSessionObject($vmachine.Id)
      # lock the vm session
      Write-Verbose "Locking the machine session"
      $global:vbox.IMachine_lockMachine($vmachine.Id,$isession,1)
      # suspend the vm
      Write-Verbose "Saving the virtual machine state"
      $global:vbox.IMachine_saveState($vmachine.Id)
     } # foreach matched $vmachine
    } # end if $vmachines
    else {Write-Verbose "Failed to find virtual machine with the name $item"}
   } # foreach $Name
  } # end if Name
 } # Try
 catch {
  Write-Verbose '$_.Exception'
  Write-Host $_.Exception -ForegroundColor Red
 } # Catch
 if ($isession) {
  Write-Verbose "Unlocking the machine session"
  $global:vbox.ISession_unlockMachine($isession)
 } # end if $isession
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Start-VirtualBoxVM {
<#
.SYNOPSIS
Start a virtual machine
.DESCRIPTION
Start virtual box machines. The default is to start them in GUI mode. You can also run them headless mode which will start a new, hidden process window. If the machine(s) disk(s) are encrypted, you must specify the -Encrypted switch and supply credentials using the -Credential parameter. The default username (identifier) is the name of the virtual machine.
.PARAMETER Name
The name of at least one virtual machine.
.PARAMETER ProgressBar
A switch to display a progress bar.
.PARAMETER SkipCheck
A switch to skip service update (for development use).
.PARAMETER Type
Specifies whether to run the virtual machine in GUI or headless mode.
.PARAMETER Encrypted
A switch to specify use of disk encryption.
.PARAMETER Credential
Powershell credentials. Must be provided if the -Encrypted switch is used.
.EXAMPLE
PS C:\> Start-VirtualBoxVM "Win10"
Starts the virtual machine called Win10 in GUI mode.
.EXAMPLE
PS C:\> Start-VirtualBoxVM "2016 Core" -Headless -Encrypted -Credential
Starts the virtual machine called 2016 Core in headless mode and provides credentials to decrypt the disk(s) on boot.
.NOTES
NAME        :  Suspend-VirtualBoxVM
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxVM
Start-VirtualBoxVM
Stop-VirtualBoxVM
.INPUTS
String[]    :  
.OUTPUTS
None
#>
[CmdletBinding(DefaultParametersetName='None')]
Param(
[Parameter(Position=0,Mandatory=$true,HelpMessage="Enter one or more virtual machine name(s)",
ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck,
[Parameter(Position=1,Mandatory=$false,HelpMessage="Enter the requested start type (Headless or Gui)",
ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
[ValidateSet("Headless","Gui")]
  [string]$Type = 'Gui',
[Parameter(ParameterSetName='Encryption',Mandatory=$false,HelpMessage="Use this switch if VM disk(s) are encrypted")]
  [switch]$Encrypted,
[Parameter(ParameterSetName='Encryption',Mandatory=$true,HelpMessage="Enter the credentials to unlock the VM disk(s)")]
  [pscredential]$Credential
) # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
 #get global vbox variable or create it if it doesn't exist create it
 if (-Not $global:vbox) {$global:vbox = Get-VirtualBox}
 # refresh vboxwebsrv variable
 if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
 # start the websrvtask if it's not running
 if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
} # Begin
Process {
 # get vm inventory (ids)
 $imachines = Get-VirtualBoxVM -SkipCheck $Name
 # get hard disk inventory (ids)
 $imediums = $global:vbox.IVirtualBox_getHardDisks($global:ivbox)
 if ($imachines) {
  foreach ($imachine in $imachines) {
  if ($imachine.State -match 'PoweredOff') {
   # create isession for the machine
   $imachine.ISession = $global:vbox.IWebsessionManager_getSessionObject($imachine.Id)
   if (-not $Encrypted) {
    foreach ($imachine in $imachines) {
     # start the vm in headless mode
     $iprogress = $global:vbox.IMachine_launchVMProcess($imachine.Id, $imachine.ISession, $Type.ToLower(),$null)
     $ipercent = $global:vbox.IProgress_getOperationPercent($iprogress)
     if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($global:vbox.IProgress_getOperationDescription($iprogress)): $($ipercent)%” -percentComplete ($ipercent)}
     do {
      $ipercent = $global:vbox.IProgress_getOperationPercent($iprogress)
      if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($global:vbox.IProgress_getOperationDescription($iprogress)): $($ipercent)%” -percentComplete ($ipercent)}
     } until ($ipercent -eq '100') # continue once the progress reaches 100%
    } # end foreach imachine
   } # end if not Encrypted
   elseif ($Encrypted) {
    foreach ($imachine in $imachines) {
     # start the vm in $Type mode
     Write-Verbose "Starting VM $($imachine.Name)"
     $iprogress = $global:vbox.IMachine_launchVMProcess($imachine.Id, $imachine.ISession, $Type.ToLower(),$null)
     $ipercent = $global:vbox.IProgress_getOperationPercent($iprogress)
     if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($global:vbox.IProgress_getOperationDescription($iprogress)): $($ipercent)%” -percentComplete ($ipercent)}
     Write-Verbose "Waiting for VM $($imachine.Name) to pause"
     do {
      # get the current machine state
      $machinestate = $global:vbox.IMachine_getState($imachine.Id)
      $ipercent = $global:vbox.IProgress_getOperationPercent($iprogress)
      if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($global:vbox.IProgress_getOperationDescription($iprogress)): $($ipercent)%” -percentComplete ($ipercent)}
     } until ($machinestate -eq 'Paused') # continue once the vm pauses for password
     Write-Verbose "VM $($imachine.Name) paused"
     # create new session object for iconsole
     Write-Verbose "Getting IConsole Session object for VM $($imachine.Name)"
     $iconsole = $global:vbox.ISession_getConsole($imachine.ISession)
     Write-Verbose "foreach disk"
     foreach ($disk in $imediums) {
      Write-Verbose "Processing disk $disk"
      Write-Verbose $imachine.Uuid
      foreach ($diskguid in $global:vbox.IMedium_getMachineIds($disk)) {
       Write-Verbose $diskguid
      }
      if ($imachine.Uuid -contains $global:vbox.IMedium_getMachineIds($disk)) {
       try {
        Write-Verbose "Checking for Password against disk"
        # check the password against the vm disk
        $global:vbox.IMedium_checkEncryptionPassword($disk, $Credential.GetNetworkCredential().Password)
        Write-Verbose  "The image is configured for encryption and the password is correct"
        # pass disk encryption password to the vm console
        Write-Verbose "Sending Identifier: $($imachine.Name) with password: $($Credential.Password)"
        $global:vbox.IConsole_addDiskEncryptionPassword($iconsole, $imachine.Name, $Credential.GetNetworkCredential().Password, $false)
        Write-Verbose "Password sent"
       }
       catch {
        Write-Verbose '$_.Exception'
        Write-Host $_.Exception -ForegroundColor Red
       }
      }
     }
    } # end foreach imachine
   } # end elseif Encrypted
   else {
    Write-Host "[Error] Invalid Type selected." -ForegroundColor Red
   } # else error
  } # end if $machine.State -match 'PoweredOff'
  else {
   Write-Host "[Error] Only VMs that have been powered off can be started. The state of $($imachine.Name) is $($imachine.State)" -ForegroundColor Red
  } # else error
  } # foreach $imachine in $imachines
 } # end if imachines
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Stop-VirtualBoxVM {
<#
.SYNOPSIS
Stop a virtual machine
.DESCRIPTION
Stop one or more virtual box machines by powering them off. You may also provide the -Acpi switch to send an ACPI shutdown signal. Alternatively, if a machine will not respond to an ACPI shutdown signal, you may try the -PsShutdown switch which will send a shutdown command via PowerShell. Credentials will be required if -PsShutdown is used.
.PARAMETER Name
The name of at least one virtual machine.
.PARAMETER Acpi
A switch to send an ACPI shutdown signal to the machine.
.PARAMETER PsShutdown
A switch to send the Stop-Computer PowerShell command to the machine.
.PARAMETER Credential
Administrator credentials for the machine. Required for PsShutdown
.EXAMPLE
PS C:\> Stop-VirtualBoxVM "Win10"
Stops the virtual machine called Win10
.EXAMPLE
PS C:\> Get-VirtualBoxVM | Stop-VirtualBoxVM
Stops all running virtual machines
.NOTES
NAME        :  Stop-VirtualBoxVM
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxVM
Start-VirtualBoxVM
Suspend-VirtualBoxVM
.INPUTS
String[]    :  
.OUTPUTS
None
#>
[cmdletbinding(DefaultParameterSetName="None")]
Param(
[Parameter(Position=0,Mandatory=$true,HelpMessage="Enter a virtual machine name",
ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
  [switch]$SkipCheck,
[Parameter(ParameterSetName="Acpi")]
  [switch]$Acpi,
[Parameter(ParameterSetName="PsShutdown")]
  [switch]$PsShutdown,
[Parameter(ParameterSetName="PsShutdown",Mandatory=$true)]
  [pscredential]$Credential
) # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
 # get global vbox variable or create it if it doesn't exist create it
 if (-Not $global:vbox) {$global:vbox = Get-VirtualBox}
 # refresh vboxwebsrv variable
 if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
 # start the websrvtask if it's not running
 if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
} # Begin
Process {
 foreach ($item in $name) {
  #get the virtual machine
  $imachine = Get-VirtualBoxVM -SkipCheck $item
  if ($imachine) {
   if ($pscmdlet.ShouldProcess($imachine.name)) {
    # create Vbox session object
    Write-Verbose "Creating a session object"
    $imachine.ISession = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
    if ($Acpi) {
     Write-Verbose "ACPI Shutdown requested"
     if ($imachine.State -eq 'Running') {
      Write-verbose "Locking the machine session"
      $global:vbox.IMachine_lockMachine($imachine.Id,$imachine.ISession,1)
      # create iconsole session to vm
      Write-verbose "Creating IConsole session to the machine"
      $iconsole = $global:vbox.ISession_getConsole($imachine.ISession)
      #send ACPI shutdown signal
      Write-verbose "Sending ACPI Shutdown signal to the machine"
      $global:vbox.IConsole_powerButton($iconsole)
      # release the iconsole session
      Write-verbose "Releasing the IConsole session"
      $global:vbox.IManagedObjectRef_release($iconsole)
      # unlock the machine session
      Write-Verbose "Unlocking the machine session"
      $global:vbox.ISession_unlockMachine($imachine.ISession)
     }
     else {
      return "Only machines that are running may be stopped."
     }
    }
    elseif ($PsShutdown) {
     Write-Verbose "PowerShell Shutdown requested"
     if ($imachine.State -eq 'Running') {
      Write-verbose "Locking the machine session"
      $global:vbox.IMachine_lockMachine($imachine.Id,$imachine.ISession,1)
      # create iconsole session to vm
      Write-verbose "Creating IConsole session to the machine"
      $iconsole = $global:vbox.ISession_getConsole($imachine.ISession)
      # create iconsole guest session to vm
      Write-verbose "Creating IConsole guest session to the machine"
      $iconsoleguest = $global:vbox.IConsole_getGuest($iconsole)
      # create a guest session
      Write-Verbose "Creating a guest console session"
      $iguestsession = $global:vbox.IGuest_createSession($iconsoleguest,$Credential.GetNetworkCredential().UserName,$Credential.GetNetworkCredential().Password,$Credential.GetNetworkCredential().Domain,"PsShutdown")
      # wait 10 seconds for the session to be created successfully
      Write-Verbose "Waiting for guest console to establish successfully (timeout: 10s)"
      $iguestsessionstatus = $global:vbox.IGuestSession_waitFor($iguestsession, 1, 10000)
      Write-Verbose "Guest console status: $iguestsessionstatus"
      # create the powershell process in the guest machine and send it a stop-computer -force command and wait for 10 seconds
      Write-Verbose 'Sending PowerShell Stop-Computer -Force -Confirm:$false command (timeout: 10s)'
      $iguestprocess = $global:vbox.IGuestSession_processCreate($iguestsession, 'cmd.exe', @('/c','powershell -ExecutionPolicy Bypass -Command `"Stop-Computer -Force -Confirm:$false`"'), @(), 3, 10000) #"powershell.exe", '-ExecutionPolicy Bypass -Command Stop-Computer -Force -Confirm:$false', @(), 3, 10000)
      # release the iconsole guest session
      Write-Verbose "Releasing the IConsole guest session"
      # release the iconsole session
      Write-verbose "Releasing the IConsole session"
      $global:vbox.IManagedObjectRef_release($iconsole)
      # unlock the machine session
      Write-Verbose "Unlocking the machine session"
      $global:vbox.ISession_unlockMachine($imachine.ISession)
     }
     else {
      return "Only machines that are running may be stopped."
     }
    }
    else {
     Write-Verbose "Power-off requested"
     if ($imachine.State -eq 'Running') {
      Write-verbose "Locking the machine session"
      $global:vbox.IMachine_lockMachine($imachine.Id,$imachine.ISession,1)
      # create iconsole session to vm
      Write-verbose "Creating IConsole session to the machine"
      $iconsole = $global:vbox.ISession_getConsole($imachine.ISession)
      # Power off the machine
      Write-verbose "Powering off the machine"
      $iprogress = $global:vbox.IConsole_powerDown($iconsole)
      # release the iconsole session
      Write-verbose "Releasing the IConsole session"
      $global:vbox.IManagedObjectRef_release($iconsole)
      # unlock the machine session
      Write-Verbose "Unlocking the machine session"
      $global:vbox.ISession_unlockMachine($imachine.ISession)
     }
     else {
      return "Only machines that are running may be stopped."
     }
    }
   } #should process
  } #if vmachine
  else {
   return "No machines matching the name `"$($Name)`" found."
  } # end else
 } #foreach
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Get-VirtualBoxDisks {
<#
.SYNOPSIS
Get VirtualBox disks
.DESCRIPTION
Retrieve all VirtualBox disks.
.EXAMPLE
PS C:\Users\rich> Get-VirtualBoxDisks


Name        : GNS3 IOU VM_1.3-disk1.vmdk
Description :
Format      : VMDK
Size        : 1242759168
LogicalSize : 2147483648
VMIds       : {c9d4dc35-3967-4009-993d-1c23ab4ff22b}
VMNames     : {GNS3 IOU VM_1.3}

Name        : turnkey-lamp-disk1.vdi
Description :
Format      : vdi
Size        : 4026531840
LogicalSize : 21474836480
VMIds       : {a237e4f5-da5a-4fca-b2a6-80f9aea91a9b}
VMNames     : {WebSite}

Name        : 2016 Core.vhd
Description :
Format      : VHD
Size        : 7291584512
LogicalSize : 53687091200
VMIds       : {7353caa6-8cb6-4066-aec9-6c6a69a001b6}
VMNames     : {2016 Core}

Name        : Win10.vhd
Description :
Format      : VHD
Size        : 15747268096
LogicalSize : 53687091200
VMIds       : {15a4c311-3b89-4936-89c7-11d3340ced7a}
VMNames     : {Win10}

Gets virtual machine disks
.NOTES
NAME        :  Get-VirtualBoxDisks
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
None (Yet)
.INPUTS
None
.OUTPUTS
[VirtualBoxVHD[]]
#>
[cmdletbinding()]
Param(
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Starting $($myinvocation.mycommand)"
 # check global vbox variable and create it if it doesn't exist
 if (-Not $global:vbox) {$global:vbox = Get-VirtualBox}
 # refresh vboxwebsrv variable
 if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
 # start the websrvtask if it's not running
 if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
 if (-Not $global:ivbox) {Start-VirtualBoxSession}
} # Begin
Process {
 Write-Verbose "Getting virtual disk inventory"
 # initialize array object to hold virtual machine values
 $disks = @()
 # get virtual machine inventory
 foreach ($imediumid in ($global:vbox.IVirtualBox_getHardDisks($global:ivbox))) {
  Write-Verbose "Getting disk: $($imediumid)"
  $disk = New-Object VirtualBoxVHD
  $disk.Name = $global:vbox.IMedium_getName($imediumid)
  $disk.Description = $global:vbox.IMedium_getDescription($imediumid)
  $disk.Format = $global:vbox.IMedium_getFormat($imediumid)
  $disk.Size = $global:vbox.IMedium_getSize($imediumid)
  $disk.LogicalSize = $global:vbox.IMedium_getLogicalSize($imediumid)
  $disk.VMIds = $global:vbox.IMedium_getMachineIds($imediumid)
  foreach ($machineid in $disk.VMIds) {$disk.VMNames = (Get-VirtualBoxVM -SkipCheck | Where-Object {$_.Uuid -eq $machineid}).Name}
  $disk.State = $global:vbox.IMedium_getState($imediumid)
  $disk.Variant = $global:vbox.IMedium_getVariant($imediumid)
  $disk.Location = $global:vbox.IMedium_getLocation($imediumid)
  $disk.HostDrive = $global:vbox.IMedium_getHostDrive($imediumid)
  $disk.MediumFormat = $global:vbox.IMedium_getMediumFormat($imediumid)
  $disk.Type = $global:vbox.IMedium_getType($imediumid)
  $disk.Parent = $global:vbox.IMedium_getParent($imediumid)
  $disk.Children = $global:vbox.IMedium_getChildren($imediumid)
  $disk.Id = $imediumid
  $disk.ReadOnly = $global:vbox.IMedium_getReadOnly($imediumid)
  $disk.AutoReset = $global:vbox.IMedium_getAutoReset($imediumid)
  $disk.LastAccessError = $global:vbox.IMedium_getLastAccessError($imediumid)
  $disks += $disk
 } # end foreach loop inventory
 Write-Verbose "Found $(($disks | Measure-Object).count) disk(s)"
 if ($disks) {
  # write virtual machines object to the pipeline
  $disks
 } # end if $obj
 else {
  Write-Host "[Warning] No virtual disks found." -ForegroundColor DarkYellow
 } # end else
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
#########################################################################################
# Entry
if (!(Get-Process -ErrorAction Stop | Where-Object {$_.ProcessName -match 'VBoxWebSrv'})) {
 if (Test-Path "$($env:VBOX_MSI_INSTALL_PATH)VBoxWebSrv.exe") {
  Start-Process -FilePath "`"$($env:VBOX_MSI_INSTALL_PATH)VBoxWebSrv.exe`"" -ArgumentList "--authentication `"$authtype`"" -WindowStyle Hidden -Verbose
 }
 else {throw "VBoxWebSrv not found."}
} # end if VBoxWebSrv check
# get the global reference to the virtualbox web service object
Write-Verbose "Initializing VirtualBox environment"
$vbox = Get-VirtualBox
# get the web service task
Write-Verbose "Updating VirtualBoxWebSrv"
$vboxwebsrvtask = Update-VirtualBoxWebSrv
# define aliases
New-Alias -Name gvbox -Value Get-VirtualBox
New-Alias -Name stavboxs -Value Start-VirtualBoxSession
New-Alias -Name stovboxs -Value Stop-VirtualBoxSession
New-Alias -Name stavboxws -Value Start-VirtualBoxWebSrv
New-Alias -Name stovboxws -Value Stop-VirtualBoxWebSrv
New-Alias -Name resvboxws -Value Restart-VirtualBoxWebSrv
New-Alias -Name refvboxws -Value Update-VirtualBoxWebSrv
New-Alias -Name gvboxvm -Value Get-VirtualBoxVM
New-Alias -Name suvboxvm -Value Suspend-VirtualBoxVM
New-Alias -Name stavboxvm -Value Start-VirtualBoxVM
New-Alias -Name stovboxvm -Value Stop-VirtualBoxVM
New-Alias -Name gvboxd -Value Get-VirtualBoxDisks
# export module members
Export-ModuleMember -Alias * -Function * -Variable @('vbox','vboxwebsrvtask','vboxerror')