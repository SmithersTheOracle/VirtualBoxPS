# Requires -version 5.0
<#
TODO:
Add support for credential arrays
Create a new Disk
Modify a VM
-WhatIf support (Extremely low priority)
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
# subclasses
class IProgress {
    [ValidateNotNullOrEmpty()]
    [string]$Id
    [ValidateNotNullOrEmpty()]
	[guid]$Guid
	[string]$Description
	[string]$Initiator
	[bool]$Cancelable
	[uint64]$Percent
	[long]$TimeRemaining
	[bool]$Completed
	[bool]$Canceled
	[long]$ResultCode
	[string]$ErrorInfo
	[uint64]$OperationCount
	[uint64]$Operation
	[string]$OperationDescription
	[uint64]$OperationPercent
	[uint64]$OperationWeight
	[uint64]$Timeout
    [IProgress]Fetch ([string]$Id) {
        $Variable = [IProgress]::new()
        if ($Variable){
			$Variable.Id = $Id
			$Variable.Guid = $global:vbox.IProgress_getId($Id)
			$Variable.Description = $global:vbox.IProgress_getDescription($Id)
			$Variable.Initiator = $global:vbox.IProgress_getInitiator($Id)
			$Variable.Cancelable = $global:vbox.IProgress_getCancelable($Id)
			$Variable.Percent = $global:vbox.IProgress_getPercent($Id)
			$Variable.TimeRemaining = $global:vbox.IProgress_getTimeRemaining($Id)
			$Variable.Completed = $global:vbox.IProgress_getCompleted($Id)
			$Variable.Canceled = $global:vbox.IProgress_getCanceled($Id)
			#$Variable.ResultCode = $global:vbox.IProgress_getResultCode($Id)
			#$Variable.ErrorInfo = $global:vbox.IProgress_getErrorInfo($Id)
			$Variable.OperationCount = $global:vbox.IProgress_getOperationCount($Id)
			$Variable.Operation = $global:vbox.IProgress_getOperation($Id)
			$Variable.OperationDescription = $global:vbox.IProgress_getOperationDescription($Id)
			$Variable.OperationPercent = $global:vbox.IProgress_getOperationPercent($Id)
			$Variable.OperationWeight = $global:vbox.IProgress_getOperationWeight($Id)
			#$Variable.getTimeout = $global:vbox.IProgress_getTimeout($Id)
            return $Variable
        }
        else {return $null}
    }
    [IProgress]Update ([string]$Id) {
        $Variable = [IProgress]::new()
        if ($Variable){
			$Variable.Id = $Id
			$Variable.Initiator = $global:vbox.IProgress_getInitiator($Id)
			$Variable.Percent = $global:vbox.IProgress_getPercent($Id)
			$Variable.TimeRemaining = $global:vbox.IProgress_getTimeRemaining($Id)
			$Variable.Completed = $global:vbox.IProgress_getCompleted($Id)
			$Variable.Canceled = $global:vbox.IProgress_getCanceled($Id)
			#$Variable.ResultCode = $global:vbox.IProgress_getResultCode($Id)
			#$Variable.ErrorInfo = $global:vbox.IProgress_getErrorInfo($Id)
			$Variable.OperationCount = $global:vbox.IProgress_getOperationCount($Id)
			$Variable.Operation = $global:vbox.IProgress_getOperation($Id)
			$Variable.OperationDescription = $global:vbox.IProgress_getOperationDescription($Id)
			$Variable.OperationPercent = $global:vbox.IProgress_getOperationPercent($Id)
			$Variable.OperationWeight = $global:vbox.IProgress_getOperationWeight($Id)
			#$Variable.getTimeout = $global:vbox.IProgress_getTimeout($Id)
            return $Variable
        }
        else {return $null}
    }
}
Update-TypeData -TypeName IProgress -DefaultDisplayPropertySet @("GUID","Description") -Force
# property classes
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
    [IProgress]$IProgress = [IProgress]::new()
    [string]$IConsoleGuest
    [string]$IGuestSession
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
class ISystemPropertiesSupported {
    [string[]]$ParavirtProviders
    [string[]]$ClipboardModes
    [string[]]$DndModes
    [string[]]$FirmwareTypes
    [string[]]$PointingHidTypes
    [string[]]$KeyboardHidTypes
    [string[]]$VfsTypes
    [string[]]$ImportOptions
    [string[]]$ExportOptions
    [string[]]$RecordingAudioCodecs
    [string[]]$RecordingVideoCodecs
    [string[]]$RecordingVsMethods
    [string[]]$RecordingVrcModes
    [string[]]$GraphicsControllerTypes
    [string[]]$CloneOptions
    [string[]]$AutostopTypes
    [string[]]$VmProcPriorities
    [string[]]$NetworkAttachmentTypes
    [string[]]$NetworkAdapterTypes
    [string[]]$PortModes
    [string[]]$UartTypes
    [string[]]$USBControllerTypes
    [string[]]$AudioDriverTypes
    [string[]]$AudioControllerTypes
    [string[]]$StorageBuses
    [string[]]$StorageControllerTypes
    [string[]]$ChipsetTypes
    [uint64]$MinGuestRam
    [uint64]$MaxGuestRam
    [uint64]$MinGuestVRam
    [uint64]$MaxGuestVRam
    [uint64]$MinGuestCpuCount
    [uint64]$MaxGuestCpuCount
}
# method classes
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
} # probably going to drop this in a future version - see the IVirtualBoxErrorInfo class for replacement
class IVirtualBoxErrorInfo {
# https://www.virtualbox.org/sdkref/group___virtual_box___c_o_m__result__codes.html
# https://www.virtualbox.org/sdkref/interface_i_virtual_box_error_info.html
    [long]resultCode ($Id) {
        if ($Id -ne $null){return $global:vbox.IVirtualBoxErrorInfo_getResultCode($Id)}
        else {return $null}
    } # Result code of the error. Usually, it will be the same as the result code returned by the method that provided this error information, but not always.
    <#
    Example: onWin32, CoCreateInstance() will most likely return E_NOINTERFACE upon 
    unsuccessful component instantiation attempt, but not the value the component factory returned. 
    Value is typed ’long’, not ’result’, to make interface usable from scripting languages.
    ***Note: In MS COM, there is no equivalent. In XPCOM, it is the same as nsIException::result.
    #>
    [long]resultDetail ($Id) {
        if ($Id -ne $null){return $global:vbox.IVirtualBoxErrorInfo_getResultDetail($Id)}
        else {return $null}
    } # Optional result data of this error. This will vary depending on the actual error usage. By default this attribute is not being used.
    [guid]interfaceID ($Id) {
        if ($Id -ne $null){return $global:vbox.IVirtualBoxErrorInfo_getInterfaceID($Id)}
        else {return $null}
    } # UUID of the interface that defined the error.
    <#
    ***Note: In MS COM, it is the same as IErrorInfo::GetGUID, except for the data type. In XPCOM, there is no equivalent.
    #>
    [string]component ($Id) {
        if ($Id -ne $null){return $global:vbox.IVirtualBoxErrorInfo_getComponent($Id)}
        else {return $null}
    <#
    ***Note: In MS COM, it is the same as IErrorInfo::GetSource. In XPCOM, there is no equivalent.
    #>
    } # Name of the component that generated the error.
    [string]text ($Id) {
        if ($Id -ne $null){return $global:vbox.IVirtualBoxErrorInfo_getText($Id)}
        else {return $null}
    } # Text description of the error.
    <#
    ***Note: In MS COM, it is the same as IErrorInfo::GetDescription. In XPCOM, it is the same as nsIException::message.
    #>
    [string]next ($Id) {
        if ($Id -ne $null){return $global:vbox.IVirtualBoxErrorInfo_getNext($Id)}
        else {return $null}
    } # Next error object if there is any, or null otherwise.
    <#
    ***Note: In MS COM, there is no equivalent. In XPCOM, it is the same as nsIException::inner.
    #>
} # The IVirtualBoxErrorInfo interface represents extended error information.
class GuestSessionWaitForFlag {
    [uint64]ToULong ([string]$FromStr) {
        if ($FromStr){
            $ToULong = $null
            Switch ($FromStr) {
                'None'      {$ToULong = 0} # No waiting flags specified. Do not use this.
                'Start'     {$ToULong = 1} # Wait for the guest session being started.
                'Terminate' {$ToULong = 2} # Wait for the guest session being terminated.
                'Status'    {$ToULong = 3} # Wait for the next guest session status change.
                Default     {$ToULong = 0} # Default to 0.
            }
            return [uint64]$ToULong
        }
        else {return $null}
    }
    [string]ToStr ([uint64]$FromLong) {
        if ($FromLong){
            $ToStr = $null
            Switch ($FromLong) {
                0       {$ToStr = 'None'} # No waiting flags specified. Do not use this.
                1       {$ToStr = 'Start'} # Wait for the guest session being started.
                2       {$ToStr = 'Terminate'} # Wait for the guest session being terminated.
                3       {$ToStr = 'Status'} # Wait for the next guest session status change.
                Default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Unsigned Long
class LockType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'   {$ToInt = 0} # Placeholder value, do not use when obtaining a lock.
                'Shared' {$ToInt = 1} # Request only a shared lock for remote-controlling the machine. Such a lock allows changing certain VM settings which can be safely modified for a running VM.
                'Write'  {$ToInt = 2} # Lock the machine for writing. This requests an exclusive lock, i.e. there cannot be any other API client holding any type of lock for this VM concurrently. Remember that a VM process counts as an API client which implicitly holds the equivalent of a shared lock during the entire VM runtime.
                'VM'     {$ToInt = 3} # Lock the machine for writing, and create objects necessary for running a VM in this process.
                Default  {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Placeholder value, do not use when obtaining a lock.
                1       {$ToStr = 'Shared'} # Request only a shared lock for remote-controlling the machine. Such a lock allows changing certain VM settings which can be safely modified for a running VM.
                2       {$ToStr = 'Write'} # Lock the machine for writing. This requests an exclusive lock, i.e. there cannot be any other API client holding any type of lock for this VM concurrently. Remember that a VM process counts as an API client which implicitly holds the equivalent of a shared lock during the entire VM runtime.
                3       {$ToStr = 'VM'} # Lock the machine for writing, and create objects necessary for running a VM in this process.
                Default {$ToStr = 'None'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class ProcessCreateFlag {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'None'                    {$ToInt = 0} # No flag set
                'WaitForProcessStartOnly' {$ToInt = 1} # Only use the specified timeout value to wait for starting the guest process - the guest process itself then uses an infinite timeout.
                'IgnoreOrphanedProcesses' {$ToInt = 2} # Do not report an error when executed processes are still alive when VBoxService or the guest OS is shutting down.
                'Hidden'                  {$ToInt = 3} # Do not show the started process according to the guest OS guidelines.
                'Profile'                 {$ToInt = 4} # Utilize the user’s profile data when exeuting a process. Only available for Windows guests at the moment.
                'WaitForStdOut'           {$ToInt = 5} # The guest process waits until all data from stdout is read out.
                'WaitForStdErr'           {$ToInt = 6} # The guest process waits until all data from stderr is read out.
                'ExpandArguments'         {$ToInt = 7} # Expands environment variables in process arguments. ***Note: This is not yet implemented and is currently silently ignored. We will document the protocolVersion number for this feature once it appears, so don’t use it till then.
                'UnquotedArguments'       {$ToInt = 8} # Work around for Windows and OS/2 applications not following normal argument quoting and escaping rules. The arguments are passed to the application without any extra quoting, just a single space between each. ***Note: Present since VirtualBox 4.3.28 and 5.0 beta 3.
                Default                   {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'None'} # No flag set
                1       {$ToStr = 'WaitForProcessStartOnly'} # Only use the specified timeout value to wait for starting the guest process - the guest process itself then uses an infinite timeout.
                2       {$ToStr = 'IgnoreOrphanedProcesses'} # Do not report an error when executed processes are still alive when VBoxService or the guest OS is shutting down.
                3       {$ToStr = 'Hidden'} # Do not show the started process according to the guest OS guidelines.
                4       {$ToStr = 'Profile'} # Utilize the user’s profile data when exeuting a process. Only available for Windows guests at the moment.
                5       {$ToStr = 'WaitForStdOut'} # The guest process waits until all data from stdout is read out.
                6       {$ToStr = 'WaitForStdErr'} # The guest process waits until all data from stderr is read out.
                6       {$ToStr = 'ExpandArguments'} # Expands environment variables in process arguments. ***Note: This is not yet implemented and is currently silently ignored. We will document the protocolVersion number for this feature once it appears, so don’t use it till then.
                6       {$ToStr = 'UnquotedArguments'} # Work around for Windows and OS/2 applications not following normal argument quoting and escaping rules. The arguments are passed to the application without any extra quoting, just a single space between each. ***Note: Present since VirtualBox 4.3.28 and 5.0 beta 3.
                Default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class ProcessWaitForFlag {
    [uint64]ToULong ([string]$FromStr) {
        if ($FromStr){
            $ToULong = $null
            Switch ($FromStr) {
                'None'      {$ToULong = 0} # No waiting flags specified. Do not use this.
                'Start'     {$ToULong = 1} # Wait for the process being started.
                'Terminate' {$ToULong = 2} # Wait for the process being terminated.
                'StdIn'     {$ToULong = 3} # Wait for stdin becoming available.
                'StdOut'    {$ToULong = 4} # Wait for data becoming available on stdout.
                'StdErr'    {$ToULong = 5} # Wait for data becoming available on stderr.
                Default     {$ToULong = 0} # Default to 0.
            }
            return [uint64]$ToULong
        }
        else {return $null}
    }
    [string]ToStr ([uint64]$FromLong) {
        if ($FromLong){
            $ToStr = $null
            Switch ($FromLong) {
                0       {$ToStr = 'None'} # No waiting flags specified. Do not use this.
                1       {$ToStr = 'Start'} # Wait for the process being started.
                2       {$ToStr = 'Terminate'} # Wait for the process being terminated.
                3       {$ToStr = 'StdIn'} # Wait for stdin becoming available.
                4       {$ToStr = 'StdOut'} # Wait for data becoming available on stdout.
                5       {$ToStr = 'StdErr'} # Wait for data becoming available on stderr.
                Default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Unsigned Long
class VBoxEventType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Invalid'       {$ToInt = 0} # must always be first - not sure what this means Oracle...
                'Any'           {$ToInt = 1} # Wildcard for all events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                'Vetoable'      {$ToInt = 2} # Wildcard for all vetoable events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                'MachineEvent'  {$ToInt = 3} # Wildcard for all machine events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                'SnapshotEvent' {$ToInt = 4} # Wildcard for all snapshot events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                'InputEvent'    {$ToInt = 5} # Wildcard for all input device (keyboard, mouse) events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                'LastWildcard'  {$ToInt = 6} # Last wildcard.
                Default         {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Invalid'} # must always be first - not sure what this means Oracle...
                1       {$ToStr = 'Any'} # Wildcard for all events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                2       {$ToStr = 'Vetoable'} # Wildcard for all vetoable events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                3       {$ToStr = 'MachineEvent'} # Wildcard for all machine events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                4       {$ToStr = 'SnapshotEvent'} # Wildcard for all snapshot events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                5       {$ToStr = 'InputEvent'} # Wildcard for all input device (keyboard, mouse) events. Events of this type are never delivered, and only used in IEventSource::registerListener() call to simplify registration.
                6       {$ToStr = 'LastWildcard'} # Last wildcard.
                Default {$ToStr = 'Invalid'} # Default to Invalid.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class Handle {
    [uint64]ToULong ([string]$FromStr) {
        if ($FromStr){
            $ToULong = $null
            Switch ($FromStr) {
                'StdIn'  {$ToULong = 0} # 0 is usually stdin.
                'StdOut' {$ToULong = 1} # 1 is usually stdout.
                'StdErr' {$ToULong = 2} # 2 is usually stderr.
                Default  {$ToULong = 0} # Default to 0.
            }
            return [uint64]$ToULong
        }
        else {return $null}
    }
    [string]ToStr ([uint64]$FromLong) {
        if ($FromLong){
            $ToStr = $null
            Switch ($FromLong) {
                0       {$ToStr = 'StdIn'} # 0 is usually stdin.
                1       {$ToStr = 'StdOut'} # 1 is usually stdout.
                2       {$ToStr = 'StdErr'} # 2 is usually stderr.
                Default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Unsigned Long
#########################################################################################
# Variable Declarations
$authtype = "VBoxAuth"
$vboxwebsrvtask = New-Object VirtualBoxWebSrvTask
# probably going to drop this in a future version - see the IVirtualBoxErrorInfo class for replacement
$vboxerror = New-Object VirtualBoxError
$global:systempropertiessupported = New-Object ISystemPropertiesSupported
# global automatic method variables
$global:ivirtualboxerrorinfo = New-Object IVirtualBoxErrorInfo
$global:guestsessionwaitforflag = New-Object GuestSessionWaitForFlag
$global:locktype = New-Object LockType
$global:processcreateflag = New-Object ProcessCreateFlag
$global:processwaitforflag = New-Object ProcessWaitForFlag
$global:vboxeventtype = New-Object VBoxEventType
$global:handle = New-Object Handle
#########################################################################################
# Includes
# N/A
#########################################################################################
# Function Definitions
Function Get-VirtualBox {
<#
.SYNOPSIS
Get the VirtualBox Web Service
.DESCRIPTION
Create a PowerShell reference object for the VirtualBox Web Service. This command is run by default when the VirtualBoxPS module is loaded.
.EXAMPLE
PS C:\> $vbox = Get-VirtualBox
Creates a $vbox variable to reference the VirtualBox Web Service
.NOTES
NAME        :  Get-VirtualBox
VERSION     :  1.0
LAST UPDATED:  1/4/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Start-VirtualBoxSession
Stop-VirtualBoxSession
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
 $global:vbox = New-WebServiceProxy -Uri "$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\vboxwebService.wsdl" -Namespace "VirtualBox" -Class "VirtualBoxWebSrv"
 # if a session exists (probably because the module is being re-imported) try to get all the global data again
 if ($global:ivbox) {
  try {
   # get guest OS type IDs
   Write-Verbose 'Fetching guest OS type data ($global:iguestostype)'
   $global:iguestostype = $global:vbox.IVirtualBox_getGuestOSTypes($global:ivbox)
  } # Try
  catch {
   Write-Verbose 'Exception fetching guest OS type data'
   Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
  } # Catch
  try {
   # create a local copy of capabilities for quick reference
   Write-Verbose 'Fetching system properties object ($global:systemproperties)'
   $global:systemproperties = $global:vbox.IVirtualBox_getSystemProperties($global:ivbox)
  } # Try
  catch {
   Write-Verbose 'Exception fetching system properties'
   Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
  } # Catch
  try {
   Write-Verbose 'Fetching supported system properties ($global:systempropertiessupported)'
   Write-Verbose 'Fetching system properties: ParavirtProviders'
   $global:systempropertiessupported.ParavirtProviders = $global:vbox.ISystemProperties_getSupportedParavirtProviders($global:systemproperties)
   Write-Verbose 'Fetching system properties: ClipboardModes'
   $global:systempropertiessupported.ClipboardModes = $global:vbox.ISystemProperties_getSupportedClipboardModes($global:systemproperties)
   Write-Verbose 'Fetching system properties: DndModes'
   $global:systempropertiessupported.DndModes = $global:vbox.ISystemProperties_getSupportedDnDModes($global:systemproperties)
   Write-Verbose 'Fetching system properties: FirmwareTypes'
   $global:systempropertiessupported.FirmwareTypes = $global:vbox.ISystemProperties_getSupportedFirmwareTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: PointingHidTypes'
   $global:systempropertiessupported.PointingHidTypes = $global:vbox.ISystemProperties_getSupportedPointingHIDTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: KeyboardHidTypes'
   $global:systempropertiessupported.KeyboardHidTypes = $global:vbox.ISystemProperties_getSupportedKeyboardHIDTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: VfsTypes'
   $global:systempropertiessupported.VfsTypes = $global:vbox.ISystemProperties_getSupportedVFSTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: ImportOptions'
   $global:systempropertiessupported.ImportOptions = $global:vbox.ISystemProperties_getSupportedImportOptions($global:systemproperties)
   Write-Verbose 'Fetching system properties: ExportOptions'
   $global:systempropertiessupported.ExportOptions = $global:vbox.ISystemProperties_getSupportedExportOptions($global:systemproperties)
   Write-Verbose 'Fetching system properties: RecordingAudioCodecs'
   $global:systempropertiessupported.RecordingAudioCodecs = $global:vbox.ISystemProperties_getSupportedRecordingAudioCodecs($global:systemproperties)
   Write-Verbose 'Fetching system properties: RecordingVideoCodecs'
   $global:systempropertiessupported.RecordingVideoCodecs = $global:vbox.ISystemProperties_getSupportedRecordingVideoCodecs($global:systemproperties)
   Write-Verbose 'Fetching system properties: RecordingVsMethods'
   $global:systempropertiessupported.RecordingVsMethods = $global:vbox.ISystemProperties_getSupportedRecordingVSMethods($global:systemproperties)
   Write-Verbose 'Fetching system properties: RecordingVrcModes'
   $global:systempropertiessupported.RecordingVrcModes = $global:vbox.ISystemProperties_getSupportedRecordingVRCModes($global:systemproperties)
   Write-Verbose 'Fetching system properties: GraphicsControllerTypes'
   $global:systempropertiessupported.GraphicsControllerTypes = $global:vbox.ISystemProperties_getSupportedGraphicsControllerTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: CloneOptions'
   $global:systempropertiessupported.CloneOptions = $global:vbox.ISystemProperties_getSupportedCloneOptions($global:systemproperties)
   Write-Verbose 'Fetching system properties: AutostopTypes'
   $global:systempropertiessupported.AutostopTypes = $global:vbox.ISystemProperties_getSupportedAutostopTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: VmProcPriorities'
   $global:systempropertiessupported.VmProcPriorities = $global:vbox.ISystemProperties_getSupportedVMProcPriorities($global:systemproperties)
   Write-Verbose 'Fetching system properties: NetworkAttachmentTypes'
   $global:systempropertiessupported.NetworkAttachmentTypes = $global:vbox.ISystemProperties_getSupportedNetworkAttachmentTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: NetworkAdapterTypes'
   $global:systempropertiessupported.NetworkAdapterTypes = $global:vbox.ISystemProperties_getSupportedNetworkAdapterTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: PortModes'
   $global:systempropertiessupported.PortModes = $global:vbox.ISystemProperties_getSupportedPortModes($global:systemproperties)
   Write-Verbose 'Fetching system properties: UartTypes'
   $global:systempropertiessupported.UartTypes = $global:vbox.ISystemProperties_getSupportedUartTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: UsbControllerTypes'
   $global:systempropertiessupported.UsbControllerTypes = $global:vbox.ISystemProperties_getSupportedUSBControllerTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: AudioDriverTypes'
   $global:systempropertiessupported.AudioDriverTypes = $global:vbox.ISystemProperties_getSupportedAudioDriverTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: AudioControllerTypes'
   $global:systempropertiessupported.AudioControllerTypes = $global:vbox.ISystemProperties_getSupportedAudioControllerTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: StorageBuses'
   $global:systempropertiessupported.StorageBuses = $global:vbox.ISystemProperties_getSupportedStorageBuses($global:systemproperties)
   Write-Verbose 'Fetching system properties: StorageControllerTypes'
   $global:systempropertiessupported.StorageControllerTypes = $global:vbox.ISystemProperties_getSupportedStorageControllerTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: ChipsetTypes'
   $global:systempropertiessupported.ChipsetTypes = $global:vbox.ISystemProperties_getSupportedChipsetTypes($global:systemproperties)
   Write-Verbose 'Fetching system properties: MinGuestRam'
   $global:systempropertiessupported.MinGuestRam = $global:vbox.ISystemProperties_getMinGuestRAM($global:systemproperties)
   Write-Verbose 'Fetching system properties: MaxGuestRam'
   $global:systempropertiessupported.MaxGuestRam = $global:vbox.ISystemProperties_getMaxGuestRAM($global:systemproperties)
   Write-Verbose 'Fetching system properties: MinGuestVRam'
   $global:systempropertiessupported.MinGuestVRam = $global:vbox.ISystemProperties_getMinGuestVRAM($global:systemproperties)
   Write-Verbose 'Fetching system properties: MaxGuestVRam'
   $global:systempropertiessupported.MaxGuestVRam = $global:vbox.ISystemProperties_getMaxGuestVRAM($global:systemproperties)
   Write-Verbose 'Fetching system properties: MinGuestCPUCount'
   $global:systempropertiessupported.MinGuestCPUCount = $global:vbox.ISystemProperties_getMinGuestCPUCount($global:systemproperties)
   Write-Verbose 'Fetching system properties: MaxGuestCPUCount'
   $global:systempropertiessupported.MaxGuestCPUCount = $global:vbox.ISystemProperties_getMaxGuestCPUCount($global:systemproperties)
  } # Try
  catch {
   Write-Verbose 'Exception fetching supported system properties'
   Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
  } # Catch
 }
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
Starts a VirtualBox Web Service session and populates the $global:ivbox managed object reference
.DESCRIPTION
Create a PowerShell managed object reference to the VirtualBox Web Service managed object.
.PARAMETER Protocol
The protocol of the VirtualBox Web Service. Default is http.
.PARAMETER Domain
The domain name or IP address of the VirtualBox Web Service. Default is localhost.
.PARAMETER Port
The TCP port of the VirtualBox Web Service. Default is 18083.
.PARAMETER Force
A switch to force updating global properties.
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
string       : string for protocol
string       : string for IP/FQDN
string       : string for TCP port
pscredential :
.OUTPUTS
$global:ivbox
#>
[cmdletbinding()]
Param(
[Parameter(HelpMessage="Enter protocol to be used to connect to the web service (Default: http)",
Mandatory=$false,Position=0)]
[ValidateSet("http","https")]
  [string]$Protocol = "http",
# localhost ONLY for now since we haven't enabled https
[Parameter(HelpMessage="Enter the domain name or IP address running the web service (Default: localhost)",
Mandatory=$false,Position=1)]
  [string]$Domain = "localhost",
[Parameter(HelpMessage="Enter the TCP port the web service is listening on (Default: 18083)",
Mandatory=$false,Position=2)]
  [string]$Port = "18083",
[Parameter(HelpMessage="Enter the credentials used to run the web service",
Mandatory=$true,Position=3)]
  [pscredential]$Credential,
[Parameter(HelpMessage="Use this switch to force updating global properties")]
  [switch]$Force
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
  if (!$global:iguestostype -or $Force) {
   try {
    # get guest OS type IDs
    Write-Verbose 'Fetching guest OS type data ($global:iguestostype)'
    $global:iguestostype = $global:vbox.IVirtualBox_getGuestOSTypes($global:ivbox)
   } # Try
   catch {
    Write-Verbose 'Exception fetching guest OS type data'
    Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
   } # Catch
  }
  if (!$global:systemproperties -or $Force) {
   try {
    # create a local copy of capabilities for quick reference
    Write-Verbose 'Fetching system properties object ($global:systemproperties)'
    $global:systemproperties = $global:vbox.IVirtualBox_getSystemProperties($global:ivbox)
   } # Try
   catch {
    Write-Verbose 'Exception fetching system properties'
    Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
   } # Catch
   try {
    Write-Verbose 'Fetching supported system properties ($global:systempropertiessupported)'
    Write-Verbose 'Fetching system properties: ParavirtProviders'
    $global:systempropertiessupported.ParavirtProviders = $global:vbox.ISystemProperties_getSupportedParavirtProviders($global:systemproperties)
    Write-Verbose 'Fetching system properties: ClipboardModes'
    $global:systempropertiessupported.ClipboardModes = $global:vbox.ISystemProperties_getSupportedClipboardModes($global:systemproperties)
    Write-Verbose 'Fetching system properties: DndModes'
    $global:systempropertiessupported.DndModes = $global:vbox.ISystemProperties_getSupportedDnDModes($global:systemproperties)
    Write-Verbose 'Fetching system properties: FirmwareTypes'
    $global:systempropertiessupported.FirmwareTypes = $global:vbox.ISystemProperties_getSupportedFirmwareTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: PointingHidTypes'
    $global:systempropertiessupported.PointingHidTypes = $global:vbox.ISystemProperties_getSupportedPointingHIDTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: KeyboardHidTypes'
    $global:systempropertiessupported.KeyboardHidTypes = $global:vbox.ISystemProperties_getSupportedKeyboardHIDTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: VfsTypes'
    $global:systempropertiessupported.VfsTypes = $global:vbox.ISystemProperties_getSupportedVFSTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: ImportOptions'
    $global:systempropertiessupported.ImportOptions = $global:vbox.ISystemProperties_getSupportedImportOptions($global:systemproperties)
    Write-Verbose 'Fetching system properties: ExportOptions'
    $global:systempropertiessupported.ExportOptions = $global:vbox.ISystemProperties_getSupportedExportOptions($global:systemproperties)
    Write-Verbose 'Fetching system properties: RecordingAudioCodecs'
    $global:systempropertiessupported.RecordingAudioCodecs = $global:vbox.ISystemProperties_getSupportedRecordingAudioCodecs($global:systemproperties)
    Write-Verbose 'Fetching system properties: RecordingVideoCodecs'
    $global:systempropertiessupported.RecordingVideoCodecs = $global:vbox.ISystemProperties_getSupportedRecordingVideoCodecs($global:systemproperties)
    Write-Verbose 'Fetching system properties: RecordingVsMethods'
    $global:systempropertiessupported.RecordingVsMethods = $global:vbox.ISystemProperties_getSupportedRecordingVSMethods($global:systemproperties)
    Write-Verbose 'Fetching system properties: RecordingVrcModes'
    $global:systempropertiessupported.RecordingVrcModes = $global:vbox.ISystemProperties_getSupportedRecordingVRCModes($global:systemproperties)
    Write-Verbose 'Fetching system properties: GraphicsControllerTypes'
    $global:systempropertiessupported.GraphicsControllerTypes = $global:vbox.ISystemProperties_getSupportedGraphicsControllerTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: CloneOptions'
    $global:systempropertiessupported.CloneOptions = $global:vbox.ISystemProperties_getSupportedCloneOptions($global:systemproperties)
    Write-Verbose 'Fetching system properties: AutostopTypes'
    $global:systempropertiessupported.AutostopTypes = $global:vbox.ISystemProperties_getSupportedAutostopTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: VmProcPriorities'
    $global:systempropertiessupported.VmProcPriorities = $global:vbox.ISystemProperties_getSupportedVMProcPriorities($global:systemproperties)
    Write-Verbose 'Fetching system properties: NetworkAttachmentTypes'
    $global:systempropertiessupported.NetworkAttachmentTypes = $global:vbox.ISystemProperties_getSupportedNetworkAttachmentTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: NetworkAdapterTypes'
    $global:systempropertiessupported.NetworkAdapterTypes = $global:vbox.ISystemProperties_getSupportedNetworkAdapterTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: PortModes'
    $global:systempropertiessupported.PortModes = $global:vbox.ISystemProperties_getSupportedPortModes($global:systemproperties)
    Write-Verbose 'Fetching system properties: UartTypes'
    $global:systempropertiessupported.UartTypes = $global:vbox.ISystemProperties_getSupportedUartTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: UsbControllerTypes'
    $global:systempropertiessupported.UsbControllerTypes = $global:vbox.ISystemProperties_getSupportedUSBControllerTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: AudioDriverTypes'
    $global:systempropertiessupported.AudioDriverTypes = $global:vbox.ISystemProperties_getSupportedAudioDriverTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: AudioControllerTypes'
    $global:systempropertiessupported.AudioControllerTypes = $global:vbox.ISystemProperties_getSupportedAudioControllerTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: StorageBuses'
    $global:systempropertiessupported.StorageBuses = $global:vbox.ISystemProperties_getSupportedStorageBuses($global:systemproperties)
    Write-Verbose 'Fetching system properties: StorageControllerTypes'
    $global:systempropertiessupported.StorageControllerTypes = $global:vbox.ISystemProperties_getSupportedStorageControllerTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: ChipsetTypes'
    $global:systempropertiessupported.ChipsetTypes = $global:vbox.ISystemProperties_getSupportedChipsetTypes($global:systemproperties)
    Write-Verbose 'Fetching system properties: MinGuestRam'
    $global:systempropertiessupported.MinGuestRam = $global:vbox.ISystemProperties_getMinGuestRAM($global:systemproperties)
    Write-Verbose 'Fetching system properties: MaxGuestRam'
    $global:systempropertiessupported.MaxGuestRam = $global:vbox.ISystemProperties_getMaxGuestRAM($global:systemproperties)
    Write-Verbose 'Fetching system properties: MinGuestVRam'
    $global:systempropertiessupported.MinGuestVRam = $global:vbox.ISystemProperties_getMinGuestVRAM($global:systemproperties)
    Write-Verbose 'Fetching system properties: MaxGuestVRam'
    $global:systempropertiessupported.MaxGuestVRam = $global:vbox.ISystemProperties_getMaxGuestVRAM($global:systemproperties)
    Write-Verbose 'Fetching system properties: MinGuestCPUCount'
    $global:systempropertiessupported.MinGuestCPUCount = $global:vbox.ISystemProperties_getMinGuestCPUCount($global:systemproperties)
    Write-Verbose 'Fetching system properties: MaxGuestCPUCount'
    $global:systempropertiessupported.MaxGuestCPUCount = $global:vbox.ISystemProperties_getMaxGuestCPUCount($global:systemproperties)
   } # Try
   catch {
    Write-Verbose 'Exception fetching supported system properties'
    Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
   } # Catch
  }
 }
 catch {
  Write-Verbose 'Exception creating the VirtualBox Web Service session'
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
Stops the current VirtualBox Web Service session
.DESCRIPTION
Instruct the VirtualBox Web Service to close the current managed object session referenced by $global:ivbox.
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
 if ($global:ivbox) {
  try {
   # tell vboxwebsrv to end the current session
  Write-Verbose 'Closing the VirtualBox Web Service session ($global:ivbox)'
   $global:vbox.IWebsessionManager_logoff($global:ivbox)
   $global:ivbox = $null
  } # end try
  catch {
   Write-Verbose 'Exception closing the VirtualBox Web Service session'
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
Starts the VirtualBox Web Service
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
  Write-Verbose 'Exception starting the VirtualBox Web Service'
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
Stops the VirtualBox Web Service
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
  Write-Verbose 'Exception ending the VirtualBox Web Service'
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
Restarts the VirtualBox Web Service
.DESCRIPTION
Stops then starts the VirtualBox Web Service using schtask.exe.
.EXAMPLE
PS C:\> Restart-VirtualBoxWebSrv
Restarts the VirtualBox Web Service if it is running
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
Gets the updated status of the VirtualBox Web Service
.DESCRIPTION
Gets the updated status of the VirtualBox Web Service using schtask.exe.
.EXAMPLE
PS C:\> Update-VirtualBoxWebSrv
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
  Write-Verbose 'Exception updating the VirtualBox Web Service'
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
"PoweredOff","Running","Saved","Teleported","Aborted","Paused","Stuck","Snapshotting",
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
 try {
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
 } # Try
 catch {
  Write-Verbose 'Exception retreiving machine information'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
 } # Catch
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
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER SkipCheck
A switch to skip service update (for development use).
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
VirtualBoxVM[]:  Array for virtual machine objects
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
     $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession, $global:locktype.ToInt('Shared'))
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
     if ($global:vbox.ISession_getState($imachine.ISession) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
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
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER SkipCheck
A switch to skip service update (for development use).
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
VirtualBoxVM[]:  Array for virtual machine objects
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
     $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession, $global:locktype.ToInt('Shared'))
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
     if ($global:vbox.ISession_getState($imachine.ISession) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
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
Start VirtualBox VMs by machine object, name, or GUID. The default Type is to start them in GUI mode. You can also run them headless mode which will start a new hidden process. If the machine(s) disk(s) are encrypted, you must specify the -Encrypted switch and supply credentials using the -Credential parameter. The username (identifier) is the name of the virtual machine by default, unless it has been otherwise specified.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
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
PS C:\> Start-VirtualBoxVM "2016" -Headless -Encrypted -Credential $diskCredentials
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
VirtualBoxVM[]:  Array for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
PsCredential  :  Credential for virtual machine disks
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
[Parameter(ParameterSetName='Unencrypted',Mandatory=$false,
HelpMessage="Enter the requested start type (Headless or Gui)",Position=1)]
[Parameter(ParameterSetName='Encrypted',Mandatory=$false,
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
     $imachine.IProgress.Id = $global:vbox.IMachine_launchVMProcess($imachine.Id, $imachine.ISession, $Type.ToLower(),$null)
     # collect iprogress data
     Write-Verbose "Fetching IProgress data"
     $imachine.IProgress = $imachine.IProgress.Fetch($imachine.IProgress.Id)
     if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%” -percentComplete ($imachine.IProgress.Percent) -CurrentOperation “Current Operation: $($imachine.IProgress.OperationDescription)” -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)}
     do {
      # get the current machine state
      $machinestate = $global:vbox.IMachine_getState($imachine.Id)
      # update iprogress data
      $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
      if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%” -percentComplete ($imachine.IProgress.Percent) -CurrentOperation “Current Operation: $($imachine.IProgress.OperationDescription)” -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)}
      if ($ProgressBar) {Write-Progress -Activity “$($imachine.IProgress.OperationDescription)” -status “$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%” -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
     } until ($machinestate -eq 'Running') # continue once the vm is running
    } # end if not Encrypted
    elseif ($Encrypted) {
     # start the vm in $Type mode
     Write-Verbose "Starting VM $($imachine.Name) in $Type mode"
     $imachine.IProgress.Id = $global:vbox.IMachine_launchVMProcess($imachine.Id, $imachine.ISession, $Type.ToLower(), $null)
     # collect iprogress data
     Write-Verbose "Fetching IProgress data"
     $imachine.IProgress = $imachine.IProgress.Fetch($imachine.IProgress.Id)
     if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%” -percentComplete ($imachine.IProgress.Percent) -CurrentOperation “Current Operation: $($imachine.IProgress.OperationDescription)” -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)}
     Write-Verbose "Waiting for VM $($imachine.Name) to pause for password"
     do {
      # get the current machine state
      $machinestate = $global:vbox.IMachine_getState($imachine.Id)
      # update iprogress data
      $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
      if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%” -percentComplete ($imachine.IProgress.Percent) -CurrentOperation “Current Operation: $($imachine.IProgress.OperationDescription)” -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)}
      if ($ProgressBar) {Write-Progress -Activity “$($imachine.IProgress.OperationDescription)” -status “$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%” -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
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
     if ($global:vbox.ISession_getState($imachine.ISession) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
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
Stop one or more virtual box machines by powering them off. You may also provide the -Acpi switch to send an ACPI shutdown signal. Alternatively, if a machine will not respond to an ACPI shutdown signal, you may try the -PsShutdown switch which will send a shutdown command via PowerShell. Valid administrator credentials will be required if -PsShutdown is used. You can supply credentials with the -Credential parameter.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Acpi
A switch to send an ACPI shutdown signal to the machine.
.PARAMETER PsShutdown
A switch to send the Stop-Computer PowerShell command to the machine.
.PARAMETER Credential
Administrator credentials for the machine. Required if the PsShutdown switch is used.
.PARAMETER SkipCheck
A switch to skip service update (for development use).
.EXAMPLE
PS C:\> Stop-VirtualBoxVM "Win10"
Stops the VM named Win10
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State 'Running' | Stop-VirtualBoxVM
Stops all running VMs
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State 'Running' | Where-Object {$_.GuestOS -match 'win'} | Stop-VirtualBoxVM -PsShutdown -Credential $domainAdminCredentials
Sends a PowerShell shutdown command to all running Windows VMs using domain administrator credentials stored in $domainAdminCredentials
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
VirtualBoxVM[]:  Array for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
PsCredential  :  Credential for virtual machine guests
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
      $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession, $global:locktype.ToInt('Shared'))
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
      # send a stop-computer -force command to the guest machine
      Write-Verbose 'Sending PowerShell Stop-Computer -Force -Confirm:$false command to guest machine'
      Write-Output (Submit-VirtualBoxVMProcess -Machine $imachine -PathToExecutable 'cmd.exe' -Arguments '/c','powershell.exe','-ExecutionPolicy','Bypass','-Command','Stop-Computer','-Force','-Confirm:$false' -Credential $Credential -Bypass)
      #$iguestprocess = $global:vbox.IGuestSession_processCreate($imachine.IGuestSession, 'C:\\Windows\\System32\\cmd.exe', [array]@('cmd.exe','/c','powershell.exe','-ExecutionPolicy','Bypass','-Command','Stop-Computer','-Force','-Confirm:$false'), [array]@(), 3, 10000)
     }
     else {return "Only machines that are running may be stopped."}
    }
    else {
     Write-Verbose "Power-off requested"
     if ($imachine.State -eq 'Running') {
      Write-verbose "Locking the machine session"
      $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession, $global:locktype.ToInt('Shared'))
      # create iconsole session to vm
      Write-verbose "Creating IConsole session to the machine"
      $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession)
      # Power off the machine
      Write-verbose "Powering off the machine"
      $imachine.IProgress.Id = $global:vbox.IConsole_powerDown($imachine.IConsole)
      # collect iprogress data
      Write-Verbose "Fetching IProgress data"
      $imachine.IProgress = $imachine.IProgress.Fetch($imachine.IProgress.Id)
      if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%” -percentComplete ($imachine.IProgress.Percent) -CurrentOperation “Current Operation: $($imachine.IProgress.OperationDescription)” -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)}
      do {
       # get the current machine state
       $machinestate = $global:vbox.IMachine_getState($imachine.Id)
       # update iprogress data
       $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
       if ($ProgressBar) {Write-Progress -Activity “Starting VM $($imachine.Name) in $Type Mode” -status “$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%” -percentComplete ($imachine.IProgress.Percent) -CurrentOperation “Current Operation: $($imachine.IProgress.OperationDescription)” -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)}
       if ($ProgressBar) {Write-Progress -Activity “$($imachine.IProgress.OperationDescription)” -status “$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%” -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
      } until ($machinestate -eq 'Running') # continue once the vm is running
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
     if ($global:vbox.ISession_getState($imachine.ISession) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession)
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
     Write-verbose "Releasing the IGuestSession object for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IGuestSession)
    } # end if $imachine.IConsole
    if ($imachine.IConsoleGuest) {
     # release the iconsole session
     Write-verbose "Releasing the IConsoleGuest object for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsoleGuest)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
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
Function New-VirtualBoxVM {
<#
.SYNOPSIS
Create a virtual machine
.DESCRIPTION
Creates a new virtual machine. The name provided by the Name parameter must not exist in the VirtualBox inventory, or this command will fail. You can optionally supply custom values using a large number of parameters available to this command. There are too many to fully document in this help text, so tab completion has been added where it is possible. The values provided by tab completion are updated when Start-VirtualBoxSession is successfully run. To force the values to be updated again, use the -Force switch with Start-VirtualBoxSession.
.PARAMETER Name
The name of at least one virtual machine. This is a required parameter.
.PARAMETER OsTypeId
The type ID for the virtual machine guest OS. This is a required parameter.
.PARAMETER AllowTracingToAccessVM
Enable or disable tracing access to the virtual machine.
.PARAMETER AudioControllerType
The audio controller type for the virtual machine.
.PARAMETER AudioDriverType
The audio driver type for the virtual machine.
.PARAMETER AutostartDelay
The auto start delay in seconds for the virtual machine.
.PARAMETER AutostartEnabled
Enable or disable auto start for the virtual machine.
.PARAMETER AutostopType
The auto stop type for the virtual machine.
.PARAMETER ChipsetType
The chipset type for the virtual machine.
.PARAMETER ClipboardFileTransfersEnabled
Enable or disable clipboard file transfers for the virtual machine. Default value is $false.
.PARAMETER ClipboardMode
The clipboard mode for the virtual machine.
.PARAMETER CpuCount
The number of CPUs available to the virtual machine.
.PARAMETER CpuExecutionCap
The CPU execution cap for the virtual machine. Valid range is 1-100. Default value is 100.
.PARAMETER CpuHotPlugEnabled
Enable or disable CPU hotplug for the virtual machine.
.PARAMETER CpuIdPortabilityLevel
The CPUID portability level for the virtual machine. Default value is 0.
.PARAMETER CpuProfile
The CPU profile for the virtual machine.
.PARAMETER Description
The description for the virtual machine.
.PARAMETER DndMode
The drag n' drop mode for the virtual machine.
.PARAMETER EmulatedUsbCardReaderEnabled
Enable or disable emulated USB card reader for the virtual machine.
.PARAMETER FirmwareType
The firmware type for the virtual machine.
.PARAMETER Flags
Optional flags for the virtual machine.
.PARAMETER GraphicsControllerType
The graphics controller type for the virtual machine.
.PARAMETER Group
Optional virtual machine group(s).
.PARAMETER HardwareUuid
The hardware UUID for the virtual machine.
.PARAMETER HpetEnabled
Enable or disable High Precision Event Timer for the virtual machine.
.PARAMETER IoCacheEnabled
The Enable or disable IO cache for the virtual machine.
.PARAMETER IoCacheSize
The IO cache size in MB for the virtual machine.
.PARAMETER KeyboardHidType
The keyboard HID type for the virtual machine.
.PARAMETER MemoryBalloonSize
The memory balloon size in MB for the virtual machine.
.PARAMETER MemorySize
The memory size in MB for the virtual machine.
.PARAMETER NetworkAdapterType
The network adapter type for the virtual machine.
.PARAMETER NetworkAttachmentType
The network attachment type for the virtual machine.
.PARAMETER PageFusionEnabled
Enable or disable page fusion for the virtual machine.
.PARAMETER ParavirtProvider
The paravirtual provider for the virtual machine.
.PARAMETER Path
Optional flags for the virtual machine.
.PARAMETER PointingHidType
The pointing HID type for the virtual machine.
.PARAMETER PortMode
The port mode for the virtual machine.
.PARAMETER RecordingAudioCodec
The recording audio codec for the virtual machine.
.PARAMETER RecordingVideoCodec
The recording video codec for the virtual machine.
.PARAMETER RecordingVrcMode
The recording VRC mode for the virtual machine.
.PARAMETER RecordingVsMethod
The recording VS method for the virtual machine.
.PARAMETER RtcUseUtc
Enable or disable RTC to UTC conversion for the virtual machine.
.PARAMETER StorageBus
The storage bus for the virtual machine.
.PARAMETER StorageControllerType
The storage controller type for the virtual machine.
.PARAMETER TeleporterAddress
The teleporter address for the virtual machine. The default value is a blank string which will force it to listen on all
addresses.
.PARAMETER TeleporterEnabled
Enable or disable teleporter for the virtual machine. The default value is $false which will disable it.
.PARAMETER TeleporterPassword
The teleporter password for the virtual machine.
.PARAMETER TeleporterPort
The teleporter TCP port for the virtual machine. The valid range for this parameter is 0-65535. The default value is 0 which means the port is automatically selected upon power on. 
.PARAMETER TracingConfig
The tracing configuration for the virtual machine.
.PARAMETER TracingEnabled
Enable or disable tracing for the virtual machine.
.PARAMETER UartType
The emulated UART implementation type for the virtual machine.
.PARAMETER UsbControllerType
The USB controller type for the virtual machine.
.PARAMETER VfsType
The Virtual File System type for the virtual machine.
.PARAMETER VmProcPriority
The VM process priority for the virtual machine.
.PARAMETER SkipCheck
A switch to skip service update (for development use).
.EXAMPLE
PS C:\> New-VirtualBoxVM -Name "My New Win10 VM" -OsTypeId Windows10_64
Create a new virtual machine named "My New Win10 VM" with the all the recommended 64bit Windows10 defaults
.NOTES
NAME        :  New-VirtualBoxVM
VERSION     :  1.0
LAST UPDATED:  1/15/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Modify-VirtualBoxVM
.INPUTS
String        :  String for virtual machine name
String        :  String for virtual machine OS Type ID
Other optional input parameters available. Use "Get-Help New-VirtualBoxVM -Full" for a complete list.
.OUTPUTS
None
#>
[CmdletBinding(DefaultParameterSetName='Template')]
Param(
[Parameter(HelpMessage="Enter a virtual machine name",
Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(HelpMessage="Enter the path for the virtual machine",
ParameterSetName='Custom',Mandatory=$false)]
  [string]$Path,
[Parameter(HelpMessage="Enter optional virtual machine group(s)",
ParameterSetName='Custom',Mandatory=$false)]
  [string[]]$Group,
[Parameter(HelpMessage="Enter optional flags for the virtual machine",
ParameterSetName='Custom',Mandatory=$false)]
  [string]$Flags,
[Parameter(HelpMessage="Enter the description for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Description,
[Parameter(HelpMessage="Enter the hardware UUID for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [guid]$HardwareUuid,
[Parameter(HelpMessage="Enable or disable CPU hotplug for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$CpuHotPlugEnabled,
[Parameter(HelpMessage="Enter the CPU execution cap for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
[ValidateRange(1, 100)]
  [uint64]$CpuExecutionCap = 100,
[Parameter(HelpMessage="Enter the CPUID portability level for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
[ValidateRange(0, 3)]
  [uint64]$CpuIdPortabilityLevel = 0,
[Parameter(HelpMessage="Enable or disable page fusion for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$PageFusionEnabled = $false,
[Parameter(HelpMessage="Enable or disable HPET for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$HpetEnabled = $false,
[Parameter(HelpMessage="Enable or disable emulated USB card reader for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$EmulatedUsbCardReaderEnabled = $false,
[Parameter(HelpMessage="Enable or disable clipboard file transfers for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$ClipboardFileTransfersEnabled = $false,
[Parameter(HelpMessage="Enable or disable teleporter for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$TeleporterEnabled = $false,
[Parameter(HelpMessage="Enter the teleporter TCP port for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
[ValidateRange(0, 65535)]
  [uint16]$TeleporterPort = 0,
[Parameter(HelpMessage="Enter the teleporter address for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$TeleporterAddress = '',
[Parameter(HelpMessage="Enter the teleporter password for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [securestring]$TeleporterPassword,
[Parameter(HelpMessage="Enable or disable RTC to UTC conversion for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$RtcUseUtc = $false,
[Parameter(HelpMessage="Enable or disable IO cache for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$IoCacheEnabled = $false,
[Parameter(HelpMessage="Enter the IO cache size in MB for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [uint32]$IoCacheSize,
[Parameter(HelpMessage="Enable or disable tracing for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$TracingEnabled = $false,
[Parameter(HelpMessage="Enter the tracing configuration for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$TracingConfig,
[Parameter(HelpMessage="Enable or disable tracing access to the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$AllowTracingToAccessVM = $false,
[Parameter(HelpMessage="Enable or disable auto start for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [bool]$AutostartEnabled = $false,
[Parameter(HelpMessage="Enter the auto start delay in seconds for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [uint32]$AutostartDelay = 300,
[Parameter(HelpMessage="Enter the CPU profile for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CpuProfile,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
DynamicParam {
 $OsTypeIdAttributes = new-object System.Management.Automation.ParameterAttribute
 $OsTypeIdAttributes.Mandatory = $true
 $OsTypeIdAttributes.Position = 1
 $OsTypeIdAttributes.HelpMessage = 'Enter the type ID for the virtual machine guest OS'
 $OsTypeIdCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $OsTypeIdCollection.Add($OsTypeIdAttributes)
 $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute(@('Other','Other_64','Windows31','Windows95','Windows98','WindowsMe','WindowsNT3x','WindowsNT4','Windows2000','WindowsXP','WindowsXP_64','Windows2003','Windows2003_64','WindowsVista','WindowsVista_64','Windows2008','Windows2008_64','Windows7','Windows7_64','Windows8','Windows8_64','Windows81','Windows81_64','Windows2012_64','Windows10','Windows10_64','Windows2016_64','Windows2019_64','WindowsNT','WindowsNT_64','Linux22','Linux24','Linux24_64','Linux26','Linux26_64','ArchLinux','ArchLinux_64','Debian','Debian_64','Fedora','Fedora_64','Gentoo','Gentoo_64','Mandriva','Mandriva_64','Oracle','Oracle_64','RedHat','RedHat_64','OpenSUSE','OpenSUSE_64','Turbolinux','Turbolinux_64','Ubuntu','Ubuntu_64','Xandros','Xandros_64','Linux','Linux_64','Solaris','Solaris_64','OpenSolaris','OpenSolaris_64','Solaris11_64','FreeBSD','FreeBSD_64','OpenBSD','OpenBSD_64','NetBSD','NetBSD_64','OS2Warp3','OS2Warp4','OS2Warp45','OS2eCS','OS21x','OS2','MacOS','MacOS_64','MacOS106','MacOS106_64','MacOS107_64','MacOS108_64','MacOS109_64','MacOS1010_64','MacOS1011_64','MacOS1012_64','MacOS1013_64','DOS','Netware','L4','QNX','JRockitVE','VBoxBS_64'))
 if ($global:iguestostype.id) {
  $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute($global:iguestostype.id)
 }
 $OsTypeIdCollection.Add($ValidateSetOsTypeId)
 $OsTypeId = new-object -Type System.Management.Automation.RuntimeDefinedParameter("OsTypeId", [string], $OsTypeIdCollection)
 $CustomAttributes = new-object System.Management.Automation.ParameterAttribute
 $CustomAttributes.Mandatory = $false
 $CustomAttributes.ParameterSetName = 'Custom'
 $CustomAttributes.HelpMessage = 'Enter the paravirtual provider for the virtual machine'
 $ParavirtProvidersCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $ParavirtProvidersCollection.Add($CustomAttributes)
 $ValidateSetParavirtProviders = New-Object System.Management.Automation.ValidateSetAttribute(@('None','Default','Legacy','Minimal','HyperV','KVM'))
 if ($global:systempropertiessupported.ParavirtProviders) {
  $ValidateSetParavirtProviders = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.ParavirtProviders)
 }
 $ParavirtProvidersCollection.Add($ValidateSetParavirtProviders)
 $ParavirtProviders = new-object -Type System.Management.Automation.RuntimeDefinedParameter("ParavirtProvider", [string], $ParavirtProvidersCollection)
 $CustomAttributes.HelpMessage = 'Enter the clipboard mode for the virtual machine'
 $ClipboardModesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $ClipboardModesCollection.Add($CustomAttributes)
 $ValidateSetClipboardModes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disabled','HostToGuest','GuestToHost','Bidirectional'))
 if ($global:systempropertiessupported.ClipboardModes) {
  $ValidateSetClipboardModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.ClipboardModes)
 }
 $ClipboardModesCollection.Add($ValidateSetClipboardModes)
 $ClipboardModes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("ClipboardMode", [string], $ClipboardModesCollection)
 $CustomAttributes.HelpMessage = "Enter the drag n' drop mode for the virtual machine"
 $DndModesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $DndModesCollection.Add($CustomAttributes)
 $ValidateSetDndModes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disabled','HostToGuest','GuestToHost','Bidirectional'))
 if ($global:systempropertiessupported.DndModes) {
  $ValidateSetDndModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.DndModes)
 }
 $DndModesCollection.Add($ValidateSetDndModes)
 $DndModes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("DndMode", [string], $DndModesCollection)
 $CustomAttributes.HelpMessage = 'Enter the firmware type for the virtual machine'
 $FirmwareTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $FirmwareTypesCollection.Add($CustomAttributes)
 $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('BIOS','EFI','EFI32','EFI64','EFIDUAL'))
 if ($global:systempropertiessupported.FirmwareTypes) {
  $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.FirmwareTypes)
 }
 $FirmwareTypesCollection.Add($ValidateSetFirmwareTypes)
 $FirmwareTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("FirmwareType", [string], $FirmwareTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the pointing HID type for the virtual machine'
 $PointingHidTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $PointingHidTypesCollection.Add($CustomAttributes)
 $ValidateSetPointingHidTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('PS2Mouse','USBTablet','USBMultiTouch'))
 if ($global:systempropertiessupported.PointingHidTypes) {
  $ValidateSetPointingHidTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.PointingHidTypes)
 }
 $PointingHidTypesCollection.Add($ValidateSetPointingHidTypes)
 $PointingHidTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("PointingHidType", [string], $PointingHidTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the keyboard HID type for the virtual machine'
 $KeyboardHidTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $KeyboardHidTypesCollection.Add($CustomAttributes)
 $ValidateSetKeyboardHidTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('PS2Keyboard','USBKeyboard'))
 if ($global:systempropertiessupported.KeyboardHidTypes) {
  $ValidateSetKeyboardHidTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.KeyboardHidTypes)
 }
 $KeyboardHidTypesCollection.Add($ValidateSetKeyboardHidTypes)
 $KeyboardHidTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("KeyboardHidType", [string], $KeyboardHidTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the Virtual File System type for the virtual machine'
 $VfsTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $VfsTypesCollection.Add($CustomAttributes)
 $ValidateSetVfsTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('File','Cloud','S3'))
 if ($global:systempropertiessupported.VfsTypes) {
  $ValidateSetVfsTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.VfsTypes)
 }
 $VfsTypesCollection.Add($ValidateSetVfsTypes)
 $VfsTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("VfsType", [string], $VfsTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording audio codec for the virtual machine'
 $RecordingAudioCodecsCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingAudioCodecsCollection.Add($CustomAttributes)
 $ValidateSetRecordingAudioCodecs = New-Object System.Management.Automation.ValidateSetAttribute(@('Opus'))
 if ($global:systempropertiessupported.RecordingAudioCodecs) {
  $ValidateSetRecordingAudioCodecs = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingAudioCodecs)
 }
 $RecordingAudioCodecsCollection.Add($ValidateSetRecordingAudioCodecs)
 $RecordingAudioCodecs = new-object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingAudioCodec", [string], $RecordingAudioCodecsCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording video codec for the virtual machine'
 $RecordingVideoCodecsCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingVideoCodecsCollection.Add($CustomAttributes)
 $ValidateSetRecordingVideoCodecs = New-Object System.Management.Automation.ValidateSetAttribute(@('VP8'))
 if ($global:systempropertiessupported.RecordingVideoCodecs) {
  $ValidateSetRecordingVideoCodecs = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingVideoCodecs)
 }
 $RecordingVideoCodecsCollection.Add($ValidateSetRecordingVideoCodecs)
 $RecordingVideoCodecs = new-object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingVideoCodec", [string], $RecordingVideoCodecsCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording VS codec for the virtual machine'
 $RecordingVsMethodsCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingVsMethodsCollection.Add($CustomAttributes)
 $ValidateSetRecordingVsMethods = New-Object System.Management.Automation.ValidateSetAttribute(@('None'))
 if ($global:systempropertiessupported.RecordingVsMethods) {
  $ValidateSetRecordingVsMethods = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingVsMethods)
 }
 $RecordingVsMethodsCollection.Add($ValidateSetRecordingVsMethods)
 $RecordingVsMethods = new-object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingVsMethod", [string], $RecordingVsMethodsCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording VRC mode for the virtual machine'
 $RecordingVrcModesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingVrcModesCollection.Add($CustomAttributes)
 $ValidateSetRecordingVrcModes = New-Object System.Management.Automation.ValidateSetAttribute(@('CBR'))
 if ($global:systempropertiessupported.RecordingVrcModes) {
  $ValidateSetRecordingVrcModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingVrcModes)
 }
 $RecordingVrcModesCollection.Add($ValidateSetRecordingVrcModes)
 $RecordingVrcModes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingVrcMode", [string], $RecordingVrcModesCollection)
 $CustomAttributes.HelpMessage = 'Enter the graphics controller type for the virtual machine'
 $GraphicsControllerTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $GraphicsControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetGraphicsControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('VBoxVGA','VMSVGA','VBoxSVGA','Null'))
 if ($global:systempropertiessupported.GraphicsControllerTypes) {
  $ValidateSetGraphicsControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.GraphicsControllerTypes)
 }
 $GraphicsControllerTypesCollection.Add($ValidateSetGraphicsControllerTypes)
 $GraphicsControllerTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("GraphicsControllerType", [string], $GraphicsControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the auto stop type for the virtual machine'
 $AutostopTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $AutostopTypesCollection.Add($CustomAttributes)
 $ValidateSetAutostopTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disabled','SaveState','PowerOff','AcpiShutdown'))
 if ($global:systempropertiessupported.AutostopTypes) {
  $ValidateSetAutostopTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.AutostopTypes)
 }
 $AutostopTypesCollection.Add($ValidateSetAutostopTypes)
 $AutostopTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("AutostopType", [string], $AutostopTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the VM process priority for the virtual machine'
 $VmProcPrioritiesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $VmProcPrioritiesCollection.Add($CustomAttributes)
 $ValidateSetVmProcPriorities = New-Object System.Management.Automation.ValidateSetAttribute(@('Default','Flat','Low','Normal','High'))
 if ($global:systempropertiessupported.VmProcPriorities) {
  $ValidateSetVmProcPriorities = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.VmProcPriorities)
 }
 $VmProcPrioritiesCollection.Add($ValidateSetVmProcPriorities)
 $VmProcPriorities = new-object -Type System.Management.Automation.RuntimeDefinedParameter("VmProcPriority", [string], $VmProcPrioritiesCollection)
 $CustomAttributes.HelpMessage = 'Enter the network attachment type for the virtual machine'
 $NetworkAttachmentTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $NetworkAttachmentTypesCollection.Add($CustomAttributes)
 $ValidateSetNetworkAttachmentTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('NAT','Bridged','Internal','HostOnly','Generic','NATNetwork','Null'))
 if ($global:systempropertiessupported.NetworkAttachmentTypes) {
  $ValidateSetNetworkAttachmentTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.NetworkAttachmentTypes)
 }
 $NetworkAttachmentTypesCollection.Add($ValidateSetNetworkAttachmentTypes)
 $NetworkAttachmentTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("NetworkAttachmentType", [string], $NetworkAttachmentTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the network adapter type for the virtual machine'
 $NetworkAdapterTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $NetworkAdapterTypesCollection.Add($CustomAttributes)
 $ValidateSetNetworkAdapterTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('Am79C970A','Am79C973','I82540EM','I82543GC','I82545EM','Virtio','Am79C960'))
 if ($global:systempropertiessupported.NetworkAdapterTypes) {
  $ValidateSetNetworkAdapterTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.NetworkAdapterTypes)
 }
 $NetworkAdapterTypesCollection.Add($ValidateSetNetworkAdapterTypes)
 $NetworkAdapterTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("NetworkAdapterType", [string], $NetworkAdapterTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the port mode for the virtual machine'
 $PortModesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $PortModesCollection.Add($CustomAttributes)
 $ValidateSetPortModes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disconnected','HostPipe','HostDevice','RawFile','TCP'))
 if ($global:systempropertiessupported.PortModes) {
  $ValidateSetPortModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.PortModes)
 }
 $PortModesCollection.Add($ValidateSetPortModes)
 $PortModes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("PortMode", [string], $PortModesCollection)
 $CustomAttributes.HelpMessage = 'Enter the emulated UART implementation type for the virtual machine'
 $UartTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $UartTypesCollection.Add($CustomAttributes)
 $ValidateSetUartTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('U16450','U16550A','U16750'))
 if ($global:systempropertiessupported.UartTypes) {
  $ValidateSetUartTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.UartTypes)
 }
 $UartTypesCollection.Add($ValidateSetUartTypes)
 $UartTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("UartType", [string], $UartTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the USB controller type for the virtual machine'
 $UsbControllerTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $UsbControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetUsbControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('OHCI','EHCI','XHCI'))
 if ($global:systempropertiessupported.UsbControllerTypes) {
  $ValidateSetUsbControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.UsbControllerTypes)
 }
 $UsbControllerTypesCollection.Add($ValidateSetUsbControllerTypes)
 $UsbControllerTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("UsbControllerType", [string], $UsbControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the audio driver type for the virtual machine'
 $AudioDriverTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $AudioDriverTypesCollection.Add($CustomAttributes)
 $ValidateSetAudioDriverTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('DirectSound','Null'))
 if ($global:systempropertiessupported.AudioDriverTypes) {
  $ValidateSetAudioDriverTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.AudioDriverTypes)
 }
 $AudioDriverTypesCollection.Add($ValidateSetAudioDriverTypes)
 $AudioDriverTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("AudioDriverType", [string], $AudioDriverTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the audio controller type for the virtual machine'
 $AudioControllerTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $AudioControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetAudioControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('AC97','SB16','HDA'))
 if ($global:systempropertiessupported.AudioControllerTypes) {
  $ValidateSetAudioControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.AudioControllerTypes)
 }
 $AudioControllerTypesCollection.Add($ValidateSetAudioControllerTypes)
 $AudioControllerTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("AudioControllerType", [string], $AudioControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the storage bus for the virtual machine'
 $StorageBusesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $StorageBusesCollection.Add($CustomAttributes)
 $ValidateSetStorageBuses = New-Object System.Management.Automation.ValidateSetAttribute(@('SATA','IDE','SCSI','Floppy','SAS','USB','PCIe','VirtioSCSI'))
 if ($global:systempropertiessupported.StorageBuses) {
  $ValidateSetStorageBuses = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.StorageBuses)
 }
 $StorageBusesCollection.Add($ValidateSetStorageBuses)
 $StorageBuses = new-object -Type System.Management.Automation.RuntimeDefinedParameter("StorageBus", [string], $StorageBusesCollection)
 $CustomAttributes.HelpMessage = 'Enter the storage controller type for the virtual machine'
 $StorageControllerTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $StorageControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetStorageControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('IntelAhci','PIIX4','PIIX3','ICH6','LsiLogic','BusLogic','I82078','LsiLogicSas','USB','NVMe','VirtioSCSI'))
 if ($global:systempropertiessupported.StorageControllerTypes) {
  $ValidateSetStorageControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.StorageControllerTypes)
 }
 $StorageControllerTypesCollection.Add($ValidateSetStorageControllerTypes)
 $StorageControllerTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("StorageControllerType", [string], $StorageControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the chipset type for the virtual machine'
 $ChipsetTypesCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $ChipsetTypesCollection.Add($CustomAttributes)
 $ValidateSetChipsetTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('PIIX3','ICH9'))
 if ($global:systempropertiessupported.ChipsetTypes) {
  $ValidateSetChipsetTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.ChipsetTypes)
 }
 $ChipsetTypesCollection.Add($ValidateSetChipsetTypes)
 $ChipsetTypes = new-object -Type System.Management.Automation.RuntimeDefinedParameter("ChipsetType", [string], $ChipsetTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the number of CPUs available to the virtual machine'
 $CpuCountCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $CpuCountCollection.Add($CustomAttributes)
 $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute(1, 32)
 if ($global:systempropertiessupported.MinGuestCPUCount -and $global:systempropertiessupported.MaxGuestCPUCount) {
  $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestCPUCount, $global:systempropertiessupported.MaxGuestCPUCount)
 }
 $CpuCountCollection.Add($ValidateSetCpuCount)
 $CpuCount = new-object -Type System.Management.Automation.RuntimeDefinedParameter("CpuCount", [uint64], $CpuCountCollection)
 $CustomAttributes.HelpMessage = 'Enter the memory size in MB for the virtual machine'
 $MemorySizeCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $MemorySizeCollection.Add($CustomAttributes)
 $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute(4, 2097152)
 if ($global:systempropertiessupported.MinGuestRam -and $global:systempropertiessupported.MaxGuestRam) {
  $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestRam, $global:systempropertiessupported.MaxGuestRam)
 }
 $MemorySizeCollection.Add($ValidateSetMemorySize)
 $MemorySize = new-object -Type System.Management.Automation.RuntimeDefinedParameter("MemorySize", [uint64], $MemorySizeCollection)
 $CustomAttributes.HelpMessage = 'Enter the memory balloon size in MB for the virtual machine'
 $MemoryBalloonSizeCollection = new-object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $MemoryBalloonSizeCollection.Add($CustomAttributes)
 $ValidateSetMemoryBalloonSize = New-Object System.Management.Automation.ValidateRangeAttribute(4, 2097152)
 if ($global:systempropertiessupported.MinGuestRam -and $global:systempropertiessupported.MaxGuestRam) {
  $ValidateSetMemoryBalloonSize = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestRam, $global:systempropertiessupported.MaxGuestRam)
 }
 $MemoryBalloonSizeCollection.Add($ValidateSetMemoryBalloonSize)
 $MemoryBalloonSize = new-object -Type System.Management.Automation.RuntimeDefinedParameter("MemoryBalloonSize", [uint64], $MemoryBalloonSizeCollection)
 $paramDictionary = new-object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
 $paramDictionary.Add("OsTypeId", $OsTypeId)
 $paramDictionary.Add("ParavirtProvider", $ParavirtProviders)
 $paramDictionary.Add("ClipboardMode", $ClipboardModes)
 $paramDictionary.Add("DndMode", $DndModes)
 $paramDictionary.Add("FirmwareType", $FirmwareTypes)
 $paramDictionary.Add("PointingHidType", $PointingHidTypes)
 $paramDictionary.Add("KeyboardHidType", $KeyboardHidTypes)
 $paramDictionary.Add("VfsType", $VfsTypes)
 $paramDictionary.Add("RecordingAudioCodec", $RecordingAudioCodecs)
 $paramDictionary.Add("RecordingVideoCodec", $RecordingVideoCodecs)
 $paramDictionary.Add("RecordingVsMethod", $RecordingVsMethods)
 $paramDictionary.Add("RecordingVrcMode", $RecordingVrcModes)
 $paramDictionary.Add("GraphicsControllerType", $GraphicsControllerTypes)
 $paramDictionary.Add("AutostopType", $AutostopTypes)
 $paramDictionary.Add("VmProcPriority", $VmProcPriorities)
 $paramDictionary.Add("NetworkAttachmentType", $NetworkAttachmentTypes)
 $paramDictionary.Add("NetworkAdapterType", $NetworkAdapterTypes)
 $paramDictionary.Add("PortMode", $PortModes)
 $paramDictionary.Add("UartType", $UartTypes)
 $paramDictionary.Add("UsbControllerType", $UsbControllerTypes)
 $paramDictionary.Add("AudioDriverType", $AudioDriverTypes)
 $paramDictionary.Add("AudioControllerType", $AudioControllerTypes)
 $paramDictionary.Add("StorageBus", $StorageBuses)
 $paramDictionary.Add("StorageControllerType", $StorageControllerTypes)
 $paramDictionary.Add("ChipsetType", $ChipsetTypes)
 $paramDictionary.Add("CpuCount", $CpuCount)
 $paramDictionary.Add("MemorySize", $MemorySize)
 $paramDictionary.Add("MemoryBalloonSize", $MemoryBalloonSize)
 return $paramDictionary
}
Begin {
 Write-Verbose "Ending $($myinvocation.mycommand)"
 #get global vbox variable or create it if it doesn't exist create it
 if (-Not $global:vbox) {$global:vbox = Get-VirtualBox}
 # refresh vboxwebsrv variable
 if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
 # start the websrvtask if it's not running
 if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
 $OsTypeId = $PSBoundParameters['OsTypeId']
 $ParavirtProvider = $PSBoundParameters['ParavirtProvider']
 $ClipboardMode = $PSBoundParameters['ClipboardMode']
 $DndMode = $PSBoundParameters['DndMode']
 $FirmwareType = $PSBoundParameters['FirmwareType']
 $PointingHidType = $PSBoundParameters['PointingHidType']
 $KeyboardHidType = $PSBoundParameters['KeyboardHidType']
 $VfsType = $PSBoundParameters['VfsType']
 $RecordingAudioCodec = $PSBoundParameters['RecordingAudioCodec']
 $RecordingVideoCodec = $PSBoundParameters['RecordingVideoCodec']
 $RecordingVsMethod = $PSBoundParameters['RecordingVsMethod']
 $RecordingVrcMode = $PSBoundParameters['RecordingVrcMode']
 $GraphicsControllerType = $PSBoundParameters['GraphicsControllerType']
 $AutostopType = $PSBoundParameters['AutostopType']
 $VmProcPriority = $PSBoundParameters['VmProcPriority']
 $NetworkAttachmentType = $PSBoundParameters['NetworkAttachmentType']
 $NetworkAdapterType = $PSBoundParameters['NetworkAdapterType']
 $PortMode = $PSBoundParameters['PortMode']
 $UartType = $PSBoundParameters['UartType']
 $UsbControllerType = $PSBoundParameters['UsbControllerType']
 $AudioDriverType = $PSBoundParameters['AudioDriverType']
 $AudioControllerType = $PSBoundParameters['AudioControllerType']
 $StorageBus = $PSBoundParameters['StorageBus']
 $StorageControllerType = $PSBoundParameters['StorageControllerType']
 $ChipsetType = $PSBoundParameters['ChipsetType']
 $CpuCount = $PSBoundParameters['CpuCount']
 $MemorySize = $PSBoundParameters['MemorySize']
 $MemoryBalloonSize = $PSBoundParameters['MemoryBalloonSize']
} # Begin
Process {
 if (!$global:iguestostype) {throw "Could not find guest defaults. Run Start-VirtualBoxSession with the -Force switch and try again."}
 $defaultsettings = $global:iguestostype | Where-Object {$_.id -eq $OsTypeId}
 try {
  # create a reference object for the new machine
  Write-Verbose "Creating reference object for $Name"
  $imachine = New-Object VirtualBoxVM
  $imachine.Id = $global:vbox.IVirtualBox_createMachine($global:ivbox, $Path, $Name, $Group, $OsTypeId, $Flags)
  $global:vbox.IMachine_applyDefaults($imachine.Id, $null)
  if ($PsCmdlet.ParameterSetName -eq 'Custom') {
   try {
    if ($Description) {$global:vbox.IMachine_setDescription($imachine.Id, $Description)}
    if ($HardwareUuid) {$global:vbox.IMachine_setHardwareUUID($imachine.Id, $HardwareUuid)}
    if ($CpuCount) {$global:vbox.IMachine_setCPUCount($imachine.Id, $CpuCount)}
    if ($CpuHotPlugEnabled) {$global:vbox.IMachine_setCPUHotPlugEnabled($imachine.Id, $CpuHotPlugEnabled)}
    if ($CpuExecutionCap) {$global:vbox.IMachine_setCPUExecutionCap($imachine.Id, $CpuExecutionCap)}
    if ($CpuIdPortabilityLevel) {$global:vbox.IMachine_setCPUIDPortabilityLevel($imachine.Id, $CpuIdPortabilityLevel)}
    if ($MemorySize) {$global:vbox.IMachine_setMemorySize($imachine.Id, $MemorySize)}
    if ($MemoryBalloonSize) {$global:vbox.IMachine_setMemoryBalloonSize($imachine.Id, $MemoryBalloonSize)}
    if ($PageFusionEnabled) {$global:vbox.IMachine_setPageFusionEnabled($imachine.Id, $PageFusionEnabled)}
    if ($FirmwareType) {$global:vbox.IMachine_setFirmwareType($imachine.Id, $FirmwareType)}
    if ($PointingHidType) {$global:vbox.IMachine_setPointingHIDType($imachine.Id, $PointingHidType)}
    if ($KeyboardHidType) {$global:vbox.IMachine_setKeyboardHIDType($imachine.Id, $KeyboardHidType)}
    if ($HpetEnabled) {$global:vbox.IMachine_setHPETEnabled($imachine.Id, $HpetEnabled)}
    if ($ChipsetType) {$global:vbox.IMachine_setChipsetType($imachine.Id, $ChipsetType)}
    if ($EmulatedUsbCardReaderEnabled) {$global:vbox.IMachine_setEmulatedUSBCardReaderEnabled($imachine.Id, $EmulatedUsbCardReaderEnabled)}
    if ($ClipboardMode) {$global:vbox.IMachine_setClipboardMode($imachine.Id, $ClipboardMode)}
    if ($ClipboardFileTransfersEnabled) {$global:vbox.IMachine_setClipboardFileTransfersEnabled($imachine.Id, $ClipboardFileTransfersEnabled)}
    if ($DndMode) {$global:vbox.IMachine_setDnDMode($imachine.Id, $DndMode)}
    if ($TeleporterEnabled) {$global:vbox.IMachine_setTeleporterEnabled($imachine.Id, $TeleporterEnabled)}
    if ($TeleporterPort) {$global:vbox.IMachine_setTeleporterPort($imachine.Id, $TeleporterPort)}
    if ($TeleporterAddress) {$global:vbox.IMachine_setTeleporterAddress($imachine.Id, $TeleporterAddress)}
    if ($TeleporterPassword) {$global:vbox.IMachine_setTeleporterPassword($imachine.Id, [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($TeleporterPassword)))}
    if ($ParavirtProvider) {$global:vbox.IMachine_setParavirtProvider($imachine.Id, $ParavirtProvider)}
    if ($RtcUseUtc) {$global:vbox.IMachine_setRTCUseUTC($imachine.Id, $RtcUseUtc)}
    if ($IoCacheEnabled) {$global:vbox.IMachine_setIOCacheEnabled($imachine.Id, $IoCacheEnabled)}
    if ($IoCacheSize) {$global:vbox.IMachine_setIOCacheSize($imachine.Id, $IoCacheSize)}
    if ($TracingEnabled) {$global:vbox.IMachine_setTracingEnabled($imachine.Id, $TracingEnabled)}
    if ($TracingConfig) {$global:vbox.IMachine_setTracingConfig($imachine.Id, $TracingConfig)}
    if ($AllowTracingToAccessVM) {$global:vbox.IMachine_setAllowTracingToAccessVM($imachine.Id, $AllowTracingToAccessVM)}
    if ($AutostartEnabled) {$global:vbox.IMachine_setAutostartEnabled($imachine.Id, $AutostartEnabled)}
    if ($AutostartDelay) {$global:vbox.IMachine_setAutostartDelay($imachine.Id, $AutostartDelay)}
    if ($AutostopType) {$global:vbox.IMachine_setAutostopType($imachine.Id, $AutostopType)}
    if ($CPUProfile) {$global:vbox.IMachine_setCPUProfile($imachine.Id, $CPUProfile)}
   }
   catch {
    Write-Verbose 'Exception applying custom parameters to machine'
    Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
   }
  }
  $global:vbox.IMachine_saveSettings($imachine.Id)
  $global:vbox.IVirtualBox_registerMachine($global:ivbox, $imachine.Id)
 } # Try
 catch {
  Write-Verbose 'Exception creating machine'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession) {
     if ($global:vbox.ISession_getState($imachine.ISession) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
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
Function Get-VirtualBoxDisks {
<#
.SYNOPSIS
Get VirtualBox disk information
.DESCRIPTION
Retrieve VirtualBox disks by machine object, machine name, machine GUID, or all.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER MachineName
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER MachineGuid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER SkipCheck
A switch to skip service update (for development use).
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
VirtualBoxVM[]:  Array for virtual machine objects
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
 try {
 # get virtual machine disk inventory
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
 }
 catch {
  Write-Verbose 'Exception retrieving virtual disk information'
  Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
 } # Catch
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
.PARAMETER SkipCheck
A switch to skip service update (for development use).
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
[Parameter(HelpMessage="Use this switch ONLY if you send a shutdown command")]
  [switch]$Bypass,
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
 $command = "$($PathToExecutable) -- $($Arguments)"
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
    $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession, $global:locktype.ToInt('Shared'))
    # create iconsole session to vm
    Write-verbose "Creating IConsole session to the machine"
    $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession)
    # create iconsole guest session to vm
    Write-verbose "Creating IConsole guest session to the machine"
    $imachine.IConsoleGuest = $global:vbox.IConsole_getGuest($imachine.IConsole)
    # create a guest session
    Write-Verbose "Creating a guest console session"
    $imachine.IGuestSession = $global:vbox.IGuest_createSession($imachine.IConsoleGuest,$Credential.GetNetworkCredential().UserName,$Credential.GetNetworkCredential().Password,$Credential.GetNetworkCredential().Domain,"PsLaunchProcess_$($imachine.IConsoleGuest)")
    # wait 10 seconds for the session to be created successfully - this needs to be merged with the previous call
    Write-Verbose "Waiting for guest console to establish successfully (timeout: 10s)"
    $iguestsessionstatus = $global:vbox.IGuestSession_waitFor($imachine.IGuestSession, $global:guestsessionwaitforflag.ToULong('Start'), 10000)
    Write-Verbose "Guest console status: $iguestsessionstatus"
    # create the process in the guest machine and send it a list of arguments
    Write-Verbose "Sending `"$command`" command (timeout: 10s)"
    $iguestprocess = $global:vbox.IGuestSession_processCreate($imachine.IGuestSession, $PathToExecutable, $Arguments, [array]@(), $global:processcreateflag.ToInt('Hidden'), 10000)
    if (!$Bypass) {
     # create event source
     Write-Verbose "Creating event source"
     $ieventsource = $global:vbox.IConsole_getEventSource($imachine.IConsole)
     # create event listener
     Write-Verbose "Creating event listener"
     $ieventlistener = $global:vbox.IEventSource_createListener($ieventsource)
     # register event listener
     Write-Verbose "Registering event listener"
     $global:vbox.IEventSource_registerListener($ieventsource, $ieventlistener, $global:vboxeventtype.ToInt('Any'), $false)
     try {
      # wait for process creation
      Write-Verbose "Waiting for guest process to be created (timeout: 10s)"
      $processwaitresult = $global:vbox.IProcess_waitFor($iguestprocess, $global:processwaitforflag.ToULong('Start'), 10000)
      Write-Verbose "Process wait result: $($processwaitresult)"
      $ieventsublistener = $null
      do {
       # get new events
       $ievent = $global:vbox.IEventSource_getEvent($ieventsource, $ieventlistener, 200)
       if ($ievent -ne '') {
        # process new event
        Write-Verbose "Encountered event ID: $($ievent)"
        $ieventtype = $global:vbox.IEvent_getType($ievent)
        Write-Verbose "Event type: $($ieventtype)"
        if ($ieventtype -eq 'OnEventSourceChanged') {
         # new event source... let's listen
         $ieventsublistener = $global:vbox.IEventSourceChangedEvent_getListener($ievent)
         Write-Verbose "New event listener object found: $($ieventsublistener)"
        } # end if event source changed
        if ($ieventtype -eq 'OnGuestPropertyChanged') {
         $guestpropertyname = $global:vbox.IGuestPropertyChangedEvent_getName($ievent)
         $guestpropertyvalue = $global:vbox.IGuestPropertyChangedEvent_getValue($ievent)
         $guestpropertyflags = $global:vbox.IGuestPropertyChangedEvent_getFlags($ievent)
         $guestpropertytimestamp = $global:vbox.IMachine_getGuestPropertyTimestamp($imachine.Id,$guestpropertyname)
         Write-Verbose "Guest property name: $($guestpropertyname)"
         Write-Verbose "Guest property value: $($guestpropertyvalue)"
         Write-Verbose "Guest property flags: $($guestpropertyflags)"
         Write-Verbose "Guest property timestamp: $($guestpropertytimestamp)"
        }
        $global:vbox.IEventSource_eventProcessed($ieventsource, $ieventlistener, $ievent)
       } # end if $ievent -ne ''
       if ($ieventsublistener -ne $null) {$isubevent = $global:vbox.IEventSource_getEvent($ieventsource, $ieventsublistener, 200)}
       if ($isubevent -ne '') {
        Write-Verbose "Encountered sub event ID: $($isubevent)"
        $isubeventtype = $global:vbox.IEvent_getType($isubevent)
        Write-Verbose "Sub event type: $($ieventtype)"
        if ($isubeventtype -eq 'OnGuestPropertyChanged') {
         $guestpropertyname = $global:vbox.IGuestPropertyChangedEvent_getName($isubevent)
         $guestpropertyvalue = $global:vbox.IGuestPropertyChangedEvent_getValue($isubevent)
         $guestpropertyflags = $global:vbox.IGuestPropertyChangedEvent_getFlags($isubevent)
         $guestpropertytimestamp = $global:vbox.IMachine_getGuestPropertyTimestamp($imachine.Id,$guestpropertyname)
         Write-Verbose "Guest property name: $($guestpropertyname)"
         Write-Verbose "Guest property value: $($guestpropertyvalue)"
         Write-Verbose "Guest property flags: $($guestpropertyflags)"
         Write-Verbose "Guest property timestamp: $($guestpropertytimestamp)"
        }
        $global:vbox.IEventSource_eventProcessed($ieventsource, $ieventsublistener, $isubevent)
       } # end if $isubevent -ne ''
       # this is returning WaitFlagNotSupported - waiting for Stdout is not currently implemented - leaving this for when it does work since it steps over anyway
       $processwaitresult = $global:vbox.IProcess_waitForArray($iguestprocess, @($global:processwaitforflag.ToULong('StdOut'),$global:processwaitforflag.ToULong('Terminate')), 200)
       #Write-Verbose "[DEBUG] Process wait result: $($processwaitresult)"
       # read guest process stdout
       [char[]]$stdout = $global:vbox.IProcess_read($iguestprocess, $global:handle.ToULong('StdOut'), 64, 0)
       Write-Verbose "[DEBUG] StdOut: $($stdout)"
       # this should be removed after debugging $stdout
       if ($stdout -ne $null) {Write-Verbose "[DEBUG] StdOut Type: $($stdout.GetType())"}
       if ($stdout) {
        # write stdout to pipeline
        Write-Verbose "Writing StdOut to pipeline"
        Write-Output ($stdout -join '')
       } # end if $stdout
       # read guest process stderr
       $stderr = $global:vbox.IProcess_read($iguestprocess, $global:handle.ToULong('StdErr'), 64, 0)
       Write-Debug "[DEBUG] StdErr: $($stdout)"
       # write stderr to the host as error text if it contains anything
       if ($stderr) {Write-Host ($stderr -join '') -ForegroundColor Red -BackgroundColor Black}
       $iprocessstatus = $global:vbox.IProcess_getStatus($iguestprocess)
       # note the process status to look for abnormal return
       if ($iprocessstatus -notmatch 'Start') {
        if ($iprocessstatus -eq 'TerminatedNormally') {Write-Verbose 'Process terminated normally'}
        else {Write-Debug "Process status: $($iprocessstatus)"}
       } # end if $iprocessstatus -notmatch 'Start'
       $keeplooping = !$iprocessstatus.toString().contains('Terminated')
      } until (!$keeplooping)
     } # Try
     catch {
      Write-Verbose 'Exception while running process in guest machine'
      Write-Host $_.Exception -ForegroundColor Red -BackgroundColor Black
      Write-Host ' '
      Write-Host
     } # Catch
     finally {
      # unregister listener object
      Write-Verbose 'Unregistering listener'
      $global:vbox.IEventSource_unregisterListener($ieventsource, $ieventlistener)
      if (!($global:vbox.IProcess_getStatus($iguestprocess)).toString().contains('Terminated')) {
       # kill guest process if it hasn't ended yet
       Write-Verbose 'Terminating guest process'
       $global:vbox.IProcess_terminate($iguestprocess)
      } # end if process hasn't terminated
     } # Finally
    } # end if not bypass
   } # foreach $imachine in $imachines
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
     if ($global:vbox.ISession_getState($imachine.ISession) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession)
     } # end if session state not unlocked
    } # end if $imachine.ISession
    if ($imachine.IConsole) {
     # release the iconsole session object
     Write-verbose "Releasing the IConsole session object for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    # next 2 ifs only for in-guest sessions
    if ($imachine.IGuestSession -and !$Bypass) {
     # close the iconsole session
     Write-verbose "Closing the IGuestSession session for VM $($imachine.Name)"
     $global:vbox.IGuestSession_close($imachine.IGuestSession)
     # release the iconsole session
     Write-verbose "Releasing the IGuestSession object for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IGuestSession)
    } # end if $imachine.IConsole and not bypass
    if ($imachine.IConsoleGuest) {
     # release the iconsole session
     Write-verbose "Releasing the IConsoleGuest object for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsoleGuest)
    } # end if $imachine.IConsole
    $imachine.ISession = $null
    $imachine.IConsole = $null
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
.PARAMETER SkipCheck
A switch to skip service update (for development use).
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
  Start-VirtualBoxWebSrv
 }
 else {throw "VBoxWebSrv not found."}
} # end if VBoxWebSrv check
# get the global reference to the virtualbox web service object
Write-Verbose "Initializing VirtualBox environment"
if (!$vbox -or $ivbox) {$vbox = Get-VirtualBox}
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
New-Alias -Name nvboxvm -Value Stop-VirtualBoxVM
New-Alias -Name gvboxd -Value Get-VirtualBoxDisks
New-Alias -Name subvboxvmp -Value Submit-VirtualBoxVMProcess
New-Alias -Name subvboxvmpss -Value Submit-VirtualBoxVMPowerShellScript
# export module members
Export-ModuleMember -Alias * -Function * -Variable @('vbox','vboxwebsrvtask','vboxerror')