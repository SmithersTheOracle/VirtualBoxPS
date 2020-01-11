# Requires -version 3.0
<#
TODO:
Finish integrating progress bar
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
# Class Definitions
class VirtualBoxVM {
    [ValidateNotNullOrEmpty()]
    [string]$Name
    [ValidateNotNullOrEmpty()]
    [string]$Id
    [string]$MMachine
    [guid]$Guid
    [string]$Description
    [string]$MemoryMB
    [string]$State
    [bool]$Running
    [string]$Info
    [string]$GuestOS
    [string]$ISession
    [string]$MSession
    [string]$IConsole
    [string]$MConsole
    [string]$IProgress
    [string]$IConsoleGuest
    [string]$IGuestSession
    [uint64]$IPercent
}
Update-TypeData -TypeName VirtualBoxVM -DefaultDisplayPropertySet @("GUID","Name","MemoryMB","Description","State","GuestOS") -Force
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
    static [array]op_Addition($A,$B) {
        [array]$C = $null
        $C += [VirtualBoxVHD]@{Name=$A.Name;Description=$A.Description;Format=$A.Format;Size=$A.Size;LogicalSize=$A.LogicalSize;VMIds=$A.VMIds;VMNames=$A.VMNames;State=$A.State;Variant=$A.Variant;Location=$A.Location;HostDrive=$A.HostDrive;MediumFormat=$A.MediumFormat;Type=$A.Type;Parent=$A.Parent;Children=$A.Children;Id=$A.Id;ReadOnly=$A.ReadOnly;AutoReset=$A.AutoReset;LastAccessError=$A.LastAccessError}
        $C += [VirtualBoxVHD]@{Name=$B.Name;Description=$B.Description;Format=$B.Format;Size=$B.Size;LogicalSize=$B.LogicalSize;VMIds=$B.VMIds;VMNames=$B.VMNames;State=$B.State;Variant=$B.Variant;Location=$B.Location;HostDrive=$B.HostDrive;MediumFormat=$B.MediumFormat;Type=$B.Type;Parent=$B.Parent;Children=$B.Children;Id=$B.Id;ReadOnly=$B.ReadOnly;AutoReset=$B.AutoReset;LastAccessError=$B.LastAccessError}
        return $C
    }
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
 Write-Output $global:vbox
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
[Parameter(HelpMessage="Enter protocol to be used to connect to the web service (Default: http)",
Mandatory=$false,Position=0)]
[ValidateSet("http","https")]
  [string]$Protocol = "http",
# localhost ONLY for now since we haven't enabled https
[Parameter(HelpMessage="Enter the domain or IP address running the web service (Default: localhost)",
Mandatory=$false,Position=1)]
  [string]$Domain = "localhost",
[Parameter(HelpMessage="Enter the TCP port the web service is listening on (Default: 18083)",
Mandatory=$false,Position=2)]
  [string]$Port = "18083",
[Parameter(HelpMessage="Enter the credentials used to run the web service",
Mandatory=$true,Position=3)]
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
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
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
   Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
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
   & cmd /c schtasks.exe /run /tn `"$($global:vboxwebsrvtask.Path)$($global:vboxwebsrvtask.Name)`" | Write-Verbose
  }
  else {
   # return a message
   return "The VBoxWebSrv task is already running"
  }
 }
 catch {
  Write-Verbose '$_.Exception'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
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
  & cmd /c schtasks.exe /end /tn `"$($global:vboxwebsrvtask.Path)$($global:vboxwebsrvtask.Name)`" | Write-Verbose
 } # end try
 catch {
  Write-Verbose '$_.Exception'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
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
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
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
Get VirtualBox virtual machine information
.DESCRIPTION
Retrieve any or all VirtualBox virtual machines by name/GUID, state, or all. The default usage, without any parameters is to display all virtual machines.
.PARAMETER Name
The name of a virtual machine.
.PARAMETER Guid
The GUID of a virtual machine.
.PARAMETER State
Return virtual machines based on their state. Valid values are:
"Stopped","Running","Saved","Teleported","Aborted","Paused","Stuck","Snapshotting",
"Starting","Stopping","Restoring","TeleportingPausedVM","TeleportingIn","FaultTolerantSync",
"DeletingSnapshotOnline","DeletingSnapshot", and "SettingUp"
.PARAMETER SkipCheck
A switch to skip service update (for development use).
.EXAMPLE
PS C:\> Get-VirtualBoxVM
UUID        : c9d4dc35-3967-4009-993d-1c23ab4ff22b
Name        : GNS3 IOU VM_1.3
MemoryMB    : 2048
Description : VM for GNS3 (development)
State       : Saved
GuestOS     : Debian

UUID        : a237e4f5-da5a-4fca-b2a6-80f9aea91a9b
Name        : WebSite
MemoryMB    : 512
Description : LAMP Server
State       : PoweredOff
GuestOS     : Other_64

UUID        : 7353caa6-8cb6-4066-aec9-6c6a69a001b6
Name        : 2016 Core
MemoryMB    : 1024
Description :
State       : PoweredOff
GuestOS     : Windows2016_64

UUID        : 15a4c311-3b89-4936-89c7-11d3340ced7a
Name        : Win10
MemoryMB    : 2048
Description :
State       : PoweredOff
GuestOS     : Windows10_64

Return all virtual machines
.EXAMPLE
PS C:\> Get-VirtualBoxVM -Name 2016
UUID        : 7353caa6-8cb6-4066-aec9-6c6a69a001b6
Name        : 2016 Core
MemoryMB    : 1024
Description :
State       : PoweredOff
GuestOS     : Windows2016_64

Retrieve a machine by name
.EXAMPLE
PS C:\> Get-VirtualBoxVM -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6
UUID        : 7353caa6-8cb6-4066-aec9-6c6a69a001b6
Name        : 2016 Core
MemoryMB    : 1024
Description :
State       : PoweredOff
GuestOS     : Windows2016_64

Retrieve a machine by GUID
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Saved
UUID        : c9d4dc35-3967-4009-993d-1c23ab4ff22b
Name        : GNS3 IOU VM_1.3
MemoryMB    : 2048
Description : VM for GNS3 (development)
State       : Saved
GuestOS     : Debian

Get suspended virtual machines
.NOTES
NAME        :  Update-VirtualBoxWebSrv
VERSION     :  1.1
LAST UPDATED:  1/8/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Start-VirtualBoxVM
Stop-VirtualBoxVM
Suspend-VirtualBoxVM
.INPUTS
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for virtual machine states
.OUTPUTS
System.Array[]
#>
[cmdletbinding(DefaultParameterSetName="All")]
Param(
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true,Position=0)]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true,Position=0)]
  [guid[]]$Guid,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter a virtual machine state you wish to filter by")]
[ValidateSet("PoweredOff","Running","Saved","Teleported","Aborted",
   "Paused","Stuck","Snapshotting","Starting","Stopping",
   "Restoring","TeleportingPausedVM","TeleportingIn","FaultTolerantSync",
   "DeletingSnapshotOnline","DeletingSnapshot","SettingUp")]
  [string]$State,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
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
 if (!$Name) {$All = $true}
} # Begin
Process {
 #$obj = New-Object VirtualBoxVM
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
   $tempobj.Guid = $global:vbox.IMachine_getId($vmid)
   $tempobj.ISession = $global:vbox.IWebsessionManager_getSessionObject($vmid)
   # decode state
   Switch ($tempobj.State) {
    1 {$tempobj.State = "PoweredOff"}
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
   Write-Verbose "Found $($tempobj.Name) and adding to inventory"
   $vminventory += $tempobj
 } # end foreach loop inventory
 # filter virtual machines
 if ($Name -and $Name -ne "*") {
  Write-Verbose "Filtering virtual machines by name: $Name"
  foreach ($vm in $vminventory) {
   Write-Verbose "Matching $($vm.Name) to $($Name)"
   if ($vm.Name -match $Name) {
    if ($State -and $vm.State -eq $State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Id=$vm.Id;Guid=$vm.Guid;ISession=$vm.ISession}}
    elseif (!$State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Id=$vm.Id;Guid=$vm.Guid;ISession=$vm.ISession}}
   }
  }
 } # end if $Name and not *
 elseif ($Guid) {
  Write-Verbose "Filtering virtual machines by GUID: $Guid"
  foreach ($vm in $vminventory) {
   Write-Verbose "Matching $($vm.Guid) to $($Guid)"
   if ($vm.Guid -match $Guid) {
    if ($State -and $vm.State -eq $State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Id=$vm.Id;Guid=$vm.Guid}}
    elseif (!$State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Id=$vm.Id;Guid=$vm.Guid}}
   }
  }
 } # end if $Guid
 elseif ($PSCmdlet.ParameterSetName -eq "All" -or $Name -eq "*") {
  if ($State) {
   Write-Verbose "Filtering all virtual machines by state: $State"
   foreach ($vm in $vminventory) {
    if ($vm.State -eq $State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Id=$vm.Id;Guid=$vm.Guid}}
   }
  }
  else {
   Write-Verbose "Filtering all virtual machines"
   foreach ($vm in $vminventory) {
    [VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Id=$vm.Id;Guid=$vm.Guid}
   }
  }
 } # end if All
 Write-Verbose "Found $(($obj | Measure-Object).count) virtual machine(s)"
 Write-Verbose "Found $($obj.Guid)"
 if ($obj) {
  # write virtual machines object to the pipeline as an array
  Write-Output ([System.Array]$obj)
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
Suspends a running virtual machine to the paused state.
.PARAMETER Machine
At least one running virtual machine object. The object must be wrapped as a [System.Array]. Can be received via pipeline input.
.PARAMETER Name
The Name of at least one running virtual machine.
.PARAMETER GUID
The GUID of at least one running virtual machine.
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Suspend-VirtualBoxVM
Suspend all running virtual machines
.EXAMPLE
PS C:\> Suspend-VirtualBoxVM -Name "2016"
Suspend the "2016 Core" virtual machine
.EXAMPLE
PS C:\> Suspend-VirtualBoxVM -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6
Suspend the virtual machine with GUID 7353caa6-8cb6-4066-aec9-6c6a69a001b6
.NOTES
NAME        :  Suspend-VirtualBoxVM
VERSION     :  1.0
LAST UPDATED:  1/9/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Resume-VirtualBoxVM
.INPUTS
System.Array[]:  Array for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs 
.OUTPUTS
None
#>
[CmdletBinding()]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
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
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {throw "Error: You must supply at least one VM object, name, or GUID."}
 # initialize $imachines array
 $imachines = @()
 # get vm inventory (by $Machine)
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
 }
 # get vm inventory (by $Name)
 elseif ($Name) {
  Write-Verbose "Getting VM inventory from Name(s)"
  $imachines = Get-VirtualBoxVM -Name $Name -SkipCheck
 }
 # get vm inventory (by $Guid)
 elseif ($Guid) {
  Write-Verbose "Getting VM inventory from GUID(s)"
  $imachines = Get-VirtualBoxVM -Guid $Guid -SkipCheck
 }
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.State -eq 'Running') {
     Write-Verbose "Suspending $($imachine.Name)"
     # get the machine session
     Write-Verbose "Getting the machine session"
     $imachine.ISession = $global:vbox.IWebsessionManager_getSessionObject($imachine.Id)
     # lock the vm session
     Write-Verbose "Locking the machine session"
     $global:vbox.IMachine_lockMachine($imachine.Id,$imachine.ISession,1)
     # get the machine IConsole session
     Write-Verbose "Getting the machine IConsole session"
     $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession)
     # suspend the vm
     Write-Verbose "Pausing the virtual machine"
     $global:vbox.IConsole_pause($imachine.IConsole)
    } # end if $imachine.State -eq 'Running'
    else {Write-Verbose "The requested virtual machine `"$($imachine.Name)`" can't be paused because it is not running (State: $($imachine.State))"}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {throw "No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception suspending machine'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession) {
     if ($global:vbox.IMachine_getSessionState($imachine.Id) > 1) {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
    $imachine.IProgress = $null
    $imachine.IPercent = $null
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Resume-VirtualBoxVM {
<#
.SYNOPSIS
Resume a virtual machine
.DESCRIPTION
Resumes a paused virtual machine to the running state.
.PARAMETER Name
The Name of at least one paused virtual machine.
.PARAMETER GUID
The GUID of at least one paused virtual machine.
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Paused | Resume-VirtualBoxVM
Resume all paused virtual machines
.EXAMPLE
PS C:\> Resume-VirtualBoxVM -Name "2016"
Resume the "2016 Core" virtual machine
.EXAMPLE
PS C:\> Resume-VirtualBoxVM -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6
Resume the virtual machine with GUID 7353caa6-8cb6-4066-aec9-6c6a69a001b6
.NOTES
NAME        :  Resume-VirtualBoxVM
VERSION     :  1.0
LAST UPDATED:  1/9/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Suspend-VirtualBoxVM
.INPUTS
System.Array[]:  Array for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
.OUTPUTS
None
#>
[CmdletBinding()]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
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
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {throw "Error: You must supply at least one VM object, name, or GUID."}
 # initialize $imachines array
 $imachines = @()
 # get vm inventory (by $Machine)
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
 }
 # get vm inventory (by $Name)
 elseif ($Name) {
  Write-Verbose "Getting VM inventory from Name(s)"
  $imachines = Get-VirtualBoxVM -Name $Name -SkipCheck
 }
 # get vm inventory (by $Guid)
 elseif ($Guid) {
  Write-Verbose "Getting VM inventory from GUID(s)"
  $imachines = Get-VirtualBoxVM -Guid $Guid -SkipCheck
 }
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.State -eq 'Paused') {
     Write-Verbose "Resuming $($imachine.Name)"
     # get the machine session
     Write-Verbose "Getting the machine session"
     $imachine.ISession = $global:vbox.IWebsessionManager_getSessionObject($imachine.Id)
     # lock the vm session
     Write-Verbose "Locking the machine session"
     $global:vbox.IMachine_lockMachine($imachine.Id,$imachine.ISession,1)
     # get the machine IConsole session
     Write-Verbose "Getting the machine IConsole session"
     $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession)
     # resume the vm
     Write-Verbose "resuming the virtual machine"
     $global:vbox.IConsole_resume($imachine.IConsole)
    } # end if $imachine.State -eq 'Running'
    else {Write-Verbose "The requested virtual machine `"$($imachine.Name)`" can't be resumed because it is not paused (State: $($imachine.State))"}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {throw "No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception resuming machine'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession) {
     if ($global:vbox.IMachine_getSessionState($imachine.Id) > 1) {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
    $imachine.IProgress = $null
    $imachine.IPercent = $null
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
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
Start virtual box machines by machine object, name, or GUID. The default Type is to start them in GUI mode. You can also run them headless mode which will start a new, hidden process window. If the machine(s) disk(s) are encrypted, you must specify the -Encrypted switch and supply credentials using the -Credential parameter. The username (identifier) is the name of the virtual machine by default, unless it has been otherwise specified.
.PARAMETER Machine
At least one virtual machine object. The object must be wrapped as a [System.Array]. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Type
Specifies whether to run the virtual machine in GUI or headless mode.
.PARAMETER Encrypted
A switch to specify use of disk encryption.
.PARAMETER Credential
Powershell credentials. Must be provided if the -Encrypted switch is used.
.PARAMETER ProgressBar
A switch to display a progress bar.
.PARAMETER SkipCheck
A switch to skip service update (for development use).
.EXAMPLE
PS C:\> Start-VirtualBoxVM "Win10"
Starts the virtual machine called Win10 in GUI mode.
.EXAMPLE
PS C:\> Start-VirtualBoxVM "2016" -Headless -Encrypted -Credential
Starts the virtual machine called "2016 Core" in headless mode and provides credentials to decrypt the disk(s) on boot.
.NOTES
NAME        :  Start-VirtualBoxVM
VERSION     :  1.1
LAST UPDATED:  1/8/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxVM
Stop-VirtualBoxVM
.INPUTS
System.Array[]:  Array for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
PsCredential[]:  Credential for virtual machine disks
.OUTPUTS
None
#>
[CmdletBinding(DefaultParameterSetName="Unencrypted")]
Param(
[Parameter(ParameterSetName='Unencrypted',Mandatory=$false,
ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)"
,Position=0)]
[Parameter(ParameterSetName='Encrypted',Mandatory=$false,
ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)"
,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(ParameterSetName='Unencrypted',Mandatory=$false,
ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
[Parameter(ParameterSetName='Encrypted',Mandatory=$false,
ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(ParameterSetName='Unencrypted',Mandatory=$false,
ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)")]
[ValidateNotNullorEmpty()]
[Parameter(ParameterSetName='Encrypted',Mandatory=$false,
ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(ParameterSetName='Unencrypted',Mandatory=$true,
HelpMessage="Enter the requested start type (Headless or Gui)",Position=1)]
[Parameter(ParameterSetName='Encrypted',Mandatory=$true,
HelpMessage="Enter the requested start type (Headless or Gui)",Position=1)]
[ValidateSet("Headless","Gui")]
  [string]$Type = 'Gui',
[Parameter(ParameterSetName='Encrypted',Mandatory=$true,
HelpMessage="Use this switch if VM disk(s) are encrypted")]
  [switch]$Encrypted,
[Parameter(ParameterSetName='Encrypted',Mandatory=$true,
HelpMessage="Enter the credentials to unlock the VM disk(s)")]
  [pscredential]$Credential,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
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
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {throw "Error: You must supply at least one VM object, name, or GUID."}
 # initialize $imachines array
 $imachines = @()
 # get vm inventory (by $Machine)
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  if ($Encrypted) {
   Write-Verbose "Getting virtual disks from Machine(s)"
   $disks = Get-VirtualBoxDisks -Machine $Machine -SkipCheck
  }
 }
 # get vm inventory (by $Name)
 elseif ($Name) {
  Write-Verbose "Getting VM inventory from Name(s)"
  $imachines = Get-VirtualBoxVM -Name $Name -SkipCheck
  if ($Encrypted) {
   Write-Verbose "Getting virtual disks from VM Name(s)"
   $disks = Get-VirtualBoxDisks -MachineName $Name -SkipCheck
  }
 }
 # get vm inventory (by $Guid)
 elseif ($Guid) {
  Write-Verbose "Getting VM inventory from GUID(s)"
  $imachines = Get-VirtualBoxVM -Guid $Guid -SkipCheck
  if ($Encrypted) {
   Write-Verbose "Getting virtual disks from VM GUID(s)"
   $disks = Get-VirtualBoxDisks -MachineGuid $Guid -SkipCheck
  }
 }
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
   if ($imachine.State -match 'PoweredOff') {
    if (-not $Encrypted) {
     # start the vm in $Type mode
     Write-Verbose "Starting VM $($imachine.Name) in $Type mode"
     $iprogress = $global:vbox.IMachine_launchVMProcess($imachine.Id, $imachine.ISession, $Type.ToLower(),$null)
     $ipercent = $global:vbox.IProgress_getOperationPercent($iprogress)
     if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($global:vbox.IProgress_getOperationDescription($iprogress)): $($ipercent)%” -percentComplete ($ipercent)}
     do {
      $ipercent = $global:vbox.IProgress_getOperationPercent($iprogress)
      if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($global:vbox.IProgress_getOperationDescription($iprogress)): $($ipercent)%” -percentComplete ($ipercent)}
     } until ($ipercent -eq '100') # continue once the progress reaches 100%
    } # end if not Encrypted
    elseif ($Encrypted) {
     # start the vm in $Type mode
     Write-Verbose "Starting VM $($imachine.Name) in $Type mode"
     $imachine.IProgress = $global:vbox.IMachine_launchVMProcess($imachine.Id, $imachine.ISession, $Type.ToLower(),$null)
     $imachine.IPercent = $global:vbox.IProgress_getOperationPercent($imachine.IProgress)
     if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($global:vbox.IProgress_getOperationDescription($imachine.IProgress)): $($imachine.IPercent)%” -percentComplete ($imachine.IPercent)}
     Write-Verbose "Waiting for VM $($imachine.Name) to pause"
     do {
      # get the current machine state
      $machinestate = $global:vbox.IMachine_getState($imachine.Id)
      $imachine.IPercent = $global:vbox.IProgress_getOperationPercent($imachine.IProgress)
      if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($global:vbox.IProgress_getOperationDescription($imachine.IProgress)): $($imachine.IPercent)%” -percentComplete ($imachine.IPercent)}
     } until ($machinestate -eq 'Paused') # continue once the vm pauses for password
     Write-Verbose "VM $($imachine.Name) paused"
     # create new session object for iconsole
     Write-Verbose "Getting IConsole Session object for VM $($imachine.Name)"
     $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession)
     foreach ($disk in $disks) {
      Write-Verbose "Processing disk $disk"
      try {
       Write-Verbose "Checking for Password against disk"
       # check the password against the vm disk
       $global:vbox.IMedium_checkEncryptionPassword($disk.Id, $Credential.GetNetworkCredential().Password)
       Write-Verbose  "The image is configured for encryption and the password is correct"
       # pass disk encryption password to the vm console
       Write-Verbose "Sending Identifier: $($imachine.Name) with password: $($Credential.Password)"
       $global:vbox.IConsole_addDiskEncryptionPassword($imachine.IConsole, $imachine.Name, $Credential.GetNetworkCredential().Password, $false)
       Write-Verbose "Password sent"
      } # Try
      catch {
       Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
       return
      } # Catch
     } # end foreach $disk in $disks
    } # end elseif Encrypted
   } # end if $machine.State -match 'PoweredOff'
   else {throw "Only VMs that have been powered off can be started. The state of $($imachine.Name) is $($imachine.State)"}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {throw "No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception starting machine'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession) {
     if ($global:vbox.IMachine_getSessionState($imachine.Id) > 1) {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
    $imachine.IProgress = $null
    $imachine.IPercent = $null
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
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
[cmdletbinding(DefaultParameterSetName="PowerOff")]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(ParameterSetName="Acpi",Mandatory=$true,
HelpMessage="Use this switch to send the ACPI Shutdown command to the VM")]
  [switch]$Acpi,
[Parameter(ParameterSetName="PsShutdown",Mandatory=$true,
HelpMessage="Use this switch send the Stop-Computer PowerShell command to the guest OS")]
  [switch]$PsShutdown,
[Parameter(ParameterSetName="PsShutdown",Mandatory=$true,
HelpMessage="Enter the credentials to login to the guest OS")]
  [pscredential]$Credential,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
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
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {throw "Error: You must supply at least one VM object, name, or GUID."}
 # initialize $imachines array
 $imachines = @()
 # get vm inventory (by $Machine)
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
 }
 # get vm inventory (by $Name)
 elseif ($Name) {
  Write-Verbose "Getting VM inventory from Name(s)"
  $imachines = Get-VirtualBoxVM -Name $Name -SkipCheck
 }
 # get vm inventory (by $Guid)
 elseif ($Guid) {
  Write-Verbose "Getting VM inventory from GUID(s)"
  $imachines = Get-VirtualBoxVM -Guid $Guid -SkipCheck
 }
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
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
      $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession)
      #send ACPI shutdown signal
      Write-verbose "Sending ACPI Shutdown signal to the machine"
      $global:vbox.IConsole_powerButton($imachine.IConsole)
     }
     else {return "Only machines that are running may be stopped."}
    }
    elseif ($PsShutdown) {
     Write-Verbose "PowerShell Shutdown requested"
     if ($imachine.State -eq 'Running') {
      Write-verbose "Locking the machine session"
      $global:vbox.IMachine_lockMachine($imachine.Id,$imachine.ISession,1)
      # create iconsole session to vm
      Write-verbose "Creating IConsole session to the machine"
      $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession)
      # create iconsole guest session to vm
      Write-verbose "Creating IConsole guest session to the machine"
      $imachine.IConsoleGuest = $global:vbox.IConsole_getGuest($imachine.IConsole)
      # create a guest session
      Write-Verbose "Creating a guest console session"
      $imachine.IGuestSession = $global:vbox.IGuest_createSession($imachine.IConsoleGuest,$Credential.GetNetworkCredential().UserName,$Credential.GetNetworkCredential().Password,$Credential.GetNetworkCredential().Domain,"PsShutdown")
      # wait 10 seconds for the session to be created successfully - this needs to be merged with the previous call
      Write-Verbose "Waiting for guest console to establish successfully (timeout: 10s)"
      $iguestsessionstatus = $global:vbox.IGuestSession_waitFor($imachine.IGuestSession, 1, 10000)
      Write-Verbose "Guest console status: $iguestsessionstatus"
      # create the powershell process in the guest machine and send it a stop-computer -force command and wait for 10 seconds
      Write-Verbose 'Sending PowerShell Stop-Computer -Force -Confirm:$false command (timeout: 10s)'
      $iguestprocess = $global:vbox.IGuestSession_processCreate($imachine.IGuestSession, 'C:\\Windows\\System32\\cmd.exe', [array]@('cmd.exe','/c','powershell.exe','-ExecutionPolicy','Bypass','-Command','Stop-Computer','-Force','-Confirm:$false'), [array]@(), 3, 10000)
     }
     else {return "Only machines that are running may be stopped."}
    }
    else {
     Write-Verbose "Power-off requested"
     if ($imachine.State -eq 'Running') {
      Write-verbose "Locking the machine session"
      $global:vbox.IMachine_lockMachine($imachine.Id,$imachine.ISession,1)
      # create iconsole session to vm
      Write-verbose "Creating IConsole session to the machine"
      $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession)
      # Power off the machine
      Write-verbose "Powering off the machine"
      $iprogress = $global:vbox.IConsole_powerDown($imachine.IConsole)
     }
     else {return "Only machines that are running may be stopped."}
    }
   } #foreach
  } # end if $imachines
  else {throw "No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception starting machine'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession) {
     if ($global:vbox.IMachine_getSessionState($imachine.Id) > 1) {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    # next 2 ifs only for in-guest sessions
    if ($imachine.IGuestSession) {
     # release the iconsole session
     Write-verbose "Releasing the IGuestSession for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IGuestSession)
    } # end if $imachine.IConsole
    if ($imachine.IConsoleGuest) {
     # release the iconsole session
     Write-verbose "Releasing the IConsoleGuest for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsoleGuest)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
    $imachine.IProgress = $null
    $imachine.IPercent = $null
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
    # next 2 only for in-guest sessions
    $imachine.IGuestSession = $null
    $imachine.IConsoleGuest = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Get-VirtualBoxDisks {
<#
.SYNOPSIS
Get VirtualBox disk information
.DESCRIPTION
Retrieve VirtualBox disks by machine object, machine name, machine GUID, or all.
.PARAMETER Machine
At least one virtual machine object. The object must be wrapped as a [System.Array]. Can be received via pipeline input.
.PARAMETER MachineName
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER MachineGuid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.EXAMPLE
PS C:\> Get-VirtualBoxVM -Name 2016 | Get-VirtualBoxDisks

Name        : 2016 Core.vhd
Description :
Format      : VHD
Size        : 7291584512
LogicalSize : 53687091200
VMIds       : {7353caa6-8cb6-4066-aec9-6c6a69a001b6}
VMNames     : {2016 Core}

Gets virtual machine by machine object from pipeline input
.EXAMPLE
PS C:\> Get-VirtualBoxDisks -MachineName 2016

Name        : 2016 Core.vhd
Description :
Format      : VHD
Size        : 7291584512
LogicalSize : 53687091200
VMIds       : {7353caa6-8cb6-4066-aec9-6c6a69a001b6}
VMNames     : {2016 Core}

Gets virtual machine by Name
.EXAMPLE
PS C:\> Get-VirtualBoxDisks -MachineGuid 7353caa6-8cb6-4066-aec9-6c6a69a001b6

Name        : 2016 Core.vhd
Description :
Format      : VHD
Size        : 7291584512
LogicalSize : 53687091200
VMIds       : {7353caa6-8cb6-4066-aec9-6c6a69a001b6}
VMNames     : {2016 Core}

Gets virtual machine by GUID
.EXAMPLE
PS C:\> Get-VirtualBoxDisks

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

Gets all virtual machine disks
.NOTES
NAME        :  Get-VirtualBoxDisks
VERSION     :  1.1
LAST UPDATED:  1/8/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
None (Yet)
.INPUTS
System.Array[]:  Array for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
.OUTPUTS
System.Array[]
#>
[cmdletbinding(DefaultParameterSetName="All")]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="MachineName",Mandatory=$true,Position=0)]
[Alias('Name')]
  [string[]]$MachineName,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="MachineGuid",Mandatory=$true,Position=0)]
  [guid[]]$MachineGuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
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
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - MachineName: `"$MachineName`""
 Write-Verbose "Pipeline - MachineGuid: `"$MachineGuid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $MachineName -or $MachineGuid)) {throw "Error: You must supply at least one VM object, name, or GUID."}
 $disks = @()
 $obj = @()
 # get virtual machine inventory
 Write-Verbose "Getting virtual disk inventory"
 foreach ($imediumid in ($global:vbox.IVirtualBox_getHardDisks($global:ivbox))) {
  Write-Verbose "Getting disk: $($imediumid)"
  $disk = New-Object VirtualBoxVHD
  $disk.Name = $global:vbox.IMedium_getName($imediumid)
  $disk.Description = $global:vbox.IMedium_getDescription($imediumid)
  $disk.Format = $global:vbox.IMedium_getFormat($imediumid)
  $disk.Size = $global:vbox.IMedium_getSize($imediumid)
  $disk.LogicalSize = $global:vbox.IMedium_getLogicalSize($imediumid)
  $disk.VMIds = $global:vbox.IMedium_getMachineIds($imediumid)
  foreach ($machineid in $disk.VMIds) {$disk.VMNames = (Get-VirtualBoxVM -Guid $machineid -SkipCheck).Name}
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
  [VirtualBoxVHD[]]$disks += [VirtualBoxVHD]@{Name=$disk.Name;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;VMIds=$disk.VMIds;VMNames=$disk.VMNames;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;}
 } # end foreach loop inventory
 # filter by machine object
 if ($Machine) {
  foreach ($disk in $disks) {
   $matched = $false
   foreach ($vmname in $disk.VMNames) {
    Write-Verbose "Matching $vmname to $($Machine.Name)"
    if ($vmname -match $Machine.Name) {Write-Verbose "Matched $vmname to $($Machine.Name)";$matched = $true}
   }
   if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;VMIds=$disk.VMIds;VMNames=$disk.VMNames;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;}}
  }
 }
 # filter by machine name
 elseif ($MachineName) {
  foreach ($disk in $disks) {
   $matched = $false
   foreach ($vmname in $disk.VMNames) {
    Write-Verbose "Matching $vmname to $MachineName"
    if ($vmname -match $MachineName) {Write-Verbose "Matched $vmname to $MachineName";$matched = $true}
   }
   if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;VMIds=$disk.VMIds;VMNames=$disk.VMNames;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;}}
  }
 }
 # filter by machine GUID
 elseif ($MachineGuid) {
  foreach ($disk in $disks) {
   $matched = $false
   foreach ($vmguid in $disk.VMIds) {
    Write-Verbose "Matching $vmguid to $MachineGuid"
    if ($vmguid -eq $MachineGuid) {Write-Verbose "Matched $vmguid to $MachineGuid";$matched = $true}
   }
   if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;VMIds=$disk.VMIds;VMNames=$disk.VMNames;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;}}
  }
 }
 # no filter
 else {foreach ($disk in $disks) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;VMIds=$disk.VMIds;VMNames=$disk.VMNames;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;}}}
 Write-Verbose "Found $(($obj | Measure-Object).count) disk(s)"
 if ($obj) {
  # write virtual machines object to the pipeline as an array
  Write-Output ([System.Array]$obj)
 } # end if $obj
 else {
  Write-Host "[Warning] No virtual disks found." -ForegroundColor DarkYellow
 } # end else
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Submit-VirtualBoxVMProcess {
<#
.SYNOPSIS
Start a guest virtual machine process
.DESCRIPTION
Will start the requested process, with optional arguments, in the guest operating system.
.PARAMETER Machine
At least one running virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The Name of at least one running virtual machine.
.PARAMETER GUID
The GUID of at least one running virtual machine.
.PARAMETER PathToExecutable
The full path to the executable.
.PARAMETER Arguments
An array of arguments to pass the executable.
.PARAMETER Credential
Administrator/Root credentials for the machine.
.EXAMPLE
PS C:\> Submit-VirtualBoxVMProcess Win10 'cmd.exe' '/c','shutdown','/s','/f' -Credential $credentials
Runs cmd.exe in the virtual machine guest OS with the argument list "/c shutdown /s /f"
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Where-Object {$_.GuestOS -match 'windows'} | Submit-VirtualBoxVMProcess -PathToExecutable 'C:\\Windows\\System32\\gpupdate.exe' -Credential $credentials
Runs gpupdate.exe on all running virtual machines with a Windows guest OS
.NOTES
NAME        :  Submit-VirtualBoxVMProcess
VERSION     :  1.0
LAST UPDATED:  1/11/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Submit-VirtualBoxVMPowerShellScript
.INPUTS
System.Array[]:  Array for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for process to create
String[]      :  Strings for arguments to process
PsCredential[]:  Credential for virtual machine disks
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
Mandatory=$true,ParameterSetName="Machine",Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
Mandatory=$true,ParameterSetName="Name",Position=0)]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
Mandatory=$true,ParameterSetName="Guid")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Enter the full path to the executable",
Position=1,Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string]$PathToExecutable,
[Parameter(HelpMessage="Enter an array of arguments to use when creating the process",
Position=2)]
  [string[]]$Arguments,
[Parameter(Mandatory=$true,
HelpMessage="Enter the credentials to login to the guest OS")]
  [pscredential]$Credential,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
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
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {throw "Error: You must supply at least one VM object, name, or GUID."}
 if ($Arguments) {$Arguments = ,$PathToExecutable + $Arguments}
 # initialize $imachines array
 $imachines = @()
 # get vm inventory (by $Machine)
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
 }
 # get vm inventory (by $Name)
 elseif ($Name) {
  Write-Verbose "Getting VM inventory from Name(s)"
  $imachines = Get-VirtualBoxVM -Name $Name -SkipCheck
 }
 # get vm inventory (by $Guid)
 elseif ($Guid) {
  Write-Verbose "Getting VM inventory from GUID(s)"
  $imachines = Get-VirtualBoxVM -Guid $Guid -SkipCheck
 }
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    Write-verbose "Locking the machine session"
    $global:vbox.IMachine_lockMachine($imachine.Id,$imachine.ISession,1)
    # create iconsole session to vm
    Write-verbose "Creating IConsole session to the machine"
    $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession)
    # create iconsole guest session to vm
    Write-verbose "Creating IConsole guest session to the machine"
    $imachine.IConsoleGuest = $global:vbox.IConsole_getGuest($imachine.IConsole)
    # create a guest session
    Write-Verbose "Creating a guest console session"
    $imachine.IGuestSession = $global:vbox.IGuest_createSession($imachine.IConsoleGuest,$Credential.GetNetworkCredential().UserName,$Credential.GetNetworkCredential().Password,$Credential.GetNetworkCredential().Domain,"PsShutdown")
    # wait 10 seconds for the session to be created successfully - this needs to be merged with the previous call
    Write-Verbose "Waiting for guest console to establish successfully (timeout: 10s)"
    $iguestsessionstatus = $global:vbox.IGuestSession_waitFor($imachine.IGuestSession, 1, 10000)
    Write-Verbose "Guest console status: $iguestsessionstatus"
    # create the process in the guest machine and send it a list of arguments
    Write-Verbose "Sending `"$($PathToExecutable) -- $($Arguments)`" command (timeout: 10s)"
    $iguestprocess = $global:vbox.IGuestSession_processCreate($imachine.IGuestSession, $PathToExecutable, $Arguments, [array]@(), 3, 10000)
    #Write-Verbose $global:vbox.IGuestProcessEvent_getProcess($imachine.iguestsessionstatus)
   } #foreach
  } # end if $imachines
  else {throw "No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception running process in guest machine'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession) {
     if ($global:vbox.IMachine_getSessionState($imachine.Id) > 1) {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    # next 2 ifs only for in-guest sessions
    if ($imachine.IGuestSession) {
     # release the iconsole session
     Write-verbose "Releasing the IGuestSession for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IGuestSession)
    } # end if $imachine.IConsole
    if ($imachine.IConsoleGuest) {
     # release the iconsole session
     Write-verbose "Releasing the IConsoleGuest for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsoleGuest)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
    $imachine.IProgress = $null
    $imachine.IPercent = $null
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
    # next 2 only for in-guest sessions
    $imachine.IGuestSession = $null
    $imachine.IConsoleGuest = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($myinvocation.mycommand)"
} # End
} # end function
Function Submit-VirtualBoxVMPowerShellScript {
<#
.SYNOPSIS
Start a guest virtual machine process
.DESCRIPTION
Will start the requested process, with optional arguments, in the guest operating system.
.PARAMETER Machine
At least one running virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The Name of at least one running virtual machine.
.PARAMETER GUID
The GUID of at least one running virtual machine.
.PARAMETER PathToExecutable
The full path to the executable.
.PARAMETER Arguments
An array of arguments to pass the executable.
.PARAMETER Credential
Administrator/Root credentials for the machine.
.EXAMPLE
PS C:\> Submit-VirtualBoxVMPowerShellScript Win10 'cmd.exe' '/c','shutdown','/s','/f' -Credential $credentials
Runs cmd.exe in the virtual machine guest OS with the argument list "/c shutdown /s /f"
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Where-Object {$_.GuestOS -match 'windows'} | Submit-VirtualBoxVMPowerShellScript -PathToExecutable 'C:\\Windows\\System32\\gpupdate.exe' -Credential $credentials
Runs gpupdate.exe on all running virtual machines with a Windows guest OS
.NOTES
NAME        :  Submit-VirtualBoxVMPowerShellScript
VERSION     :  1.0
LAST UPDATED:  1/11/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Submit-VirtualBoxVMPowerShellScript
.INPUTS
System.Array[]:  Array for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for process to create
String[]      :  Strings for arguments to process
PsCredential[]:  Credential for virtual machine disks
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
Mandatory=$true,ParameterSetName="Machine",Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
Mandatory=$true,ParameterSetName="Name",Position=0)]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
Mandatory=$true,ParameterSetName="Guid")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(Position=1,Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string]$ScriptBlock,
[Parameter(Mandatory=$true,
HelpMessage="Enter the credentials to login to the guest OS")]
  [pscredential]$Credential,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
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
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {throw "Error: You must supply at least one VM object, name, or GUID."}
 # initialize $imachines array
 $imachines = @()
 # get vm inventory (by $Machine)
 if ($Machine) {
  foreach ($item in $Name) {
   Write-Verbose "Submitting PowerShell command to VM $($Machine.Name) by VM object"
   Submit-VirtualBoxVMProcess -Machine $Machine -PathToExecutable "cmd.exe" -Arguments "/c","powershell","-ExecutionPolicy","Bypass","-Command",$ScriptBlock -Credential $Credential -SkipCheck
  }
 }
 # get vm inventory (by $Name)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Submitting PowerShell command to VM $($Name) by Name"
   Submit-VirtualBoxVMProcess -Name $Name -PathToExecutable "cmd.exe" -Arguments "/c","powershell","-ExecutionPolicy","Bypass","-Command",$ScriptBlock -Credential $Credential -SkipCheck
  }
 }
 # get vm inventory (by $Guid)
 elseif ($Guid) {
  foreach ($item in $Name) {
   Write-Verbose "Submitting PowerShell command to VM $((Get-VirtualBoxVM -Guid $Guid -SkipCheck).Name) by GUID"
   Submit-VirtualBoxVMProcess -Guid $Guid -PathToExecutable "cmd.exe" -Arguments "/c","powershell","-ExecutionPolicy","Bypass","-Command",$ScriptBlock -Credential $Credential -SkipCheck
  }
 }
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
New-Alias -Name revboxvm -Value Resume-VirtualBoxVM
New-Alias -Name stavboxvm -Value Start-VirtualBoxVM
New-Alias -Name stovboxvm -Value Stop-VirtualBoxVM
New-Alias -Name gvboxd -Value Get-VirtualBoxDisks
New-Alias -Name subvboxvmp -Value Submit-VirtualBoxVMProcess
New-Alias -Name subvboxvmpss -Value Submit-VirtualBoxVMPowerShellScript
# export module members
Export-ModuleMember -Alias * -Function * -Variable @('vbox','vboxwebsrvtask','vboxerror')