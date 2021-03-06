# Requires -version 5.0
<#
VirtualBox API Version: 6.1
TODO:
Standardize data types (Immediate priority) - https://forums.virtualbox.org/viewtopic.php?f=34&t=96465
Add support for importing/exporting encrypted VMs
Write more comprehensive error handling (low priority)
Finish implementing -WhatIf support (Extremely low priority)

Unsupported interface found - IInternalMachineControl - too bad
void IInternalMachineControl::beginPoweringUp()
IMediumAttachment IInternalMachineControl::ejectMedium()

IConsole::display() - No. Too many commands not usable with web service.

Useful methods
IStorageController IMachine::storageControllers
wstring IMedium::getEncryptionSettings
void IMachine::mountMedium()
void IMachine::unmountMedium()
IMediumAttachment[] IMachine::getMediumAttachmentsOfController()
IMediumAttachment IMachine::getMediumAttachment()
void IMachine::attachDeviceWithoutMedium()
uuid IAppliance::getMediumIdsForPasswordId()
wstring[] IAppliance::getPasswordIds()
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
# Module Parameters
Param(
[Parameter(HelpMessage="Allows you to switch between using COM or WebSrv when importing the module",
Mandatory=$false,Position=0)]
[ValidateSet('Com','WebSrv')]
    [string]$ModuleHost = 'Com',
[Parameter(HelpMessage="Enter the credentials used to run the web service",
Mandatory=$false,Position=1)]
    [pscredential]$WebSrvCredential,
[Parameter(HelpMessage="Enter protocol to be used to connect to the web service (Default: http)",
Mandatory=$false,Position=2)]
[ValidateSet("http","https")]
    [string]$WebSrvProtocol = 'http',
[Parameter(HelpMessage="Enter the domain name or IP address running the web service (Default: localhost)",
Mandatory=$false,Position=3)]
    [string]$WebSrvAddress = 'localhost',
[Parameter(HelpMessage="Enter the TCP port the web service is listening on (Default: 18083)",
Mandatory=$false,Position=4)]
    [string]$WebSrvPort = '18083'
) # Param
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
    [System.__ComObject]$Progress
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
			if ($Variable.Completed -eq $true) {$Variable.ResultCode = $global:vbox.IProgress_getResultCode($Id)}
			if ($Variable.Completed -eq $true) {$Variable.ErrorInfo = $global:vbox.IProgress_getErrorInfo($Id)}
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
			$Variable.Description = $global:vbox.IProgress_getDescription($Id)
			$Variable.Initiator = $global:vbox.IProgress_getInitiator($Id)
			$Variable.Percent = $global:vbox.IProgress_getPercent($Id)
			$Variable.TimeRemaining = $global:vbox.IProgress_getTimeRemaining($Id)
			$Variable.Completed = $global:vbox.IProgress_getCompleted($Id)
			$Variable.Canceled = $global:vbox.IProgress_getCanceled($Id)
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
if ($ModuleHost.ToLower() -eq 'websrv') {Update-TypeData -TypeName IProgress -DefaultDisplayProperty GUID -DefaultDisplayPropertySet @("GUID","Description") -Force}
if ($ModuleHost.ToLower() -eq 'com') {Update-TypeData -TypeName IProgress -DefaultDisplayProperty GUID -DefaultDisplayPropertySet @("GUID","Description","Progress") -Force}
class IVrdeServer {
    [ValidateNotNullOrEmpty()]
    [string]$Id
	[bool]$Enabled
	[string]$AuthType
	[uint32]$AuthTimeout
	[bool]$AllowMultiConnection
	[bool]$ReuseSingleConnection
	[string]$VrdeExtPack
	[string]$AuthLibrary
	[string[]]$VrdeProperties
	[uint32]$TcpPort
	[string]$IpAddress
	[bool]$VideoChannelEnabled
	[string]$VideoChannelQuality
	[string]$VideoChannelDownscaleProtection
	[bool]$DisableClientDisplay
	[bool]$DisableClientInput
	[bool]$DisableClientAudio
	[bool]$DisableClientUsb
	[bool]$DisableClientClipboard
	[bool]$DisableClientUpstreamAudio
	[bool]$DisableClientRdpdr
	[bool]$H3dRedirectEnabled
	[string]$SecurityMethod
	[string]$SecurityServerCertificate
	[string]$SecurityServerPrivateKey
	[string]$SecurityCaCertificate
	[string]$AudioRateCorrectionMode
	[string]$AudioLogPath
    [IVrdeServer]Fetch ([string]$IMachine) {
        $Variable = [IVrdeServer]::new()
        if ($Variable){
			$Variable.Id = $global:vbox.IMachine_getVRDEServer($IMachine)
			$Variable.Enabled = $global:vbox.IVRDEServer_getEnabled($Variable.Id)
			if ($Variable.Enabled) {$Variable.AuthType = $global:vbox.IVRDEServer_getAuthType($Variable.Id)}
			if ($Variable.Enabled) {$Variable.AuthTimeout = $global:vbox.IVRDEServer_getAuthTimeout($Variable.Id)}
			if ($Variable.Enabled) {$Variable.AllowMultiConnection = $global:vbox.IVRDEServer_getAllowMultiConnection($Variable.Id)}
			if ($Variable.Enabled) {$Variable.ReuseSingleConnection = $global:vbox.IVRDEServer_getReuseSingleConnection($Variable.Id)}
			if ($Variable.Enabled) {$Variable.VrdeExtPack = $global:vbox.IVRDEServer_getVRDEExtPack($Variable.Id)}
			if ($Variable.Enabled) {$Variable.AuthLibrary = $global:vbox.IVRDEServer_getAuthLibrary($Variable.Id)}
			if ($Variable.Enabled) {$Variable.VrdeProperties = $global:vbox.IVRDEServer_getVRDEProperties($Variable.Id)}
			if ($Variable.Enabled) {$Variable.TcpPort = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'TCP/Ports')}
			if ($Variable.Enabled) {$Variable.IpAddress = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'TCP/Address')}
			if ($Variable.Enabled) {$Variable.VideoChannelEnabled = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'VideoChannel/Enabled')}
			if ($Variable.Enabled) {$Variable.VideoChannelQuality = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'VideoChannel/Quality')}
			if ($Variable.Enabled) {$Variable.VideoChannelDownscaleProtection = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'VideoChannel/DownscaleProtection')}
			if ($Variable.Enabled) {$Variable.DisableClientDisplay = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Client/DisableDisplay')}
			if ($Variable.Enabled) {$Variable.DisableClientInput = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Client/DisableInput')}
			if ($Variable.Enabled) {$Variable.DisableClientAudio = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Client/DisableAudio')}
			if ($Variable.Enabled) {$Variable.DisableClientUsb = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Client/DisableUSB')}
			if ($Variable.Enabled) {$Variable.DisableClientClipboard = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Client/DisableClipboard')}
			if ($Variable.Enabled) {$Variable.DisableClientUpstreamAudio = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Client/DisableUpstreamAudio')}
			if ($Variable.Enabled) {$Variable.DisableClientRdpdr = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Client/DisableRDPDR')}
			if ($Variable.Enabled) {$Variable.H3dRedirectEnabled = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'H3DRedirect/Enabled')}
			if ($Variable.Enabled) {$Variable.SecurityMethod = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Security/Method')}
			if ($Variable.Enabled) {$Variable.SecurityServerCertificate = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Security/ServerCertificate')}
			if ($Variable.Enabled) {$Variable.SecurityServerPrivateKey = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Security/ServerPrivateKey')}
			if ($Variable.Enabled) {$Variable.SecurityCaCertificate = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Security/CACertificate')}
			if ($Variable.Enabled) {$Variable.AudioRateCorrectionMode = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Audio/RateCorrectionMode')}
			if ($Variable.Enabled) {$Variable.AudioLogPath = $global:vbox.IVRDEServer_getVRDEProperty($Variable.Id, 'Audio/LogPath')}
            return $Variable
        }
        else {return $null}
    }
    Update () {
		$this.Enabled = $global:vbox.IVRDEServer_getEnabled($this.Id)
		if ($this.Enabled) {$this.AuthType = $global:vbox.IVRDEServer_getAuthType($this.Id)}
		if ($this.Enabled) {$this.AuthTimeout = $global:vbox.IVRDEServer_getAuthTimeout($this.Id)}
		if ($this.Enabled) {$this.AllowMultiConnection = $global:vbox.IVRDEServer_getAllowMultiConnection($this.Id)}
		if ($this.Enabled) {$this.ReuseSingleConnection = $global:vbox.IVRDEServer_getReuseSingleConnection($this.Id)}
		if ($this.Enabled) {$this.VrdeExtPack = $global:vbox.IVRDEServer_getVRDEExtPack($this.Id)}
		if ($this.Enabled) {$this.AuthLibrary = $global:vbox.IVRDEServer_getAuthLibrary($this.Id)}
		if ($this.Enabled) {$this.VrdeProperties = $global:vbox.IVRDEServer_getVRDEProperties($this.Id)}
		if ($this.Enabled) {$this.TcpPort = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'TCP/Ports')}
		if ($this.Enabled) {$this.IpAddress = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'TCP/Address')}
		if ($this.Enabled) {$this.VideoChannelEnabled = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'VideoChannel/Enabled')}
		if ($this.Enabled) {$this.VideoChannelQuality = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'VideoChannel/Quality')}
		if ($this.Enabled) {$this.VideoChannelDownscaleProtection = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'VideoChannel/DownscaleProtection')}
		if ($this.Enabled) {$this.DisableClientDisplay = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Client/DisableDisplay')}
		if ($this.Enabled) {$this.DisableClientInput = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Client/DisableInput')}
		if ($this.Enabled) {$this.DisableClientAudio = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Client/DisableAudio')}
		if ($this.Enabled) {$this.DisableClientUsb = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Client/DisableUSB')}
		if ($this.Enabled) {$this.DisableClientClipboard = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Client/DisableClipboard')}
		if ($this.Enabled) {$this.DisableClientUpstreamAudio = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Client/DisableUpstreamAudio')}
		if ($this.Enabled) {$this.DisableClientRdpdr = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Client/DisableRDPDR')}
		if ($this.Enabled) {$this.H3dRedirectEnabled = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'H3DRedirect/Enabled')}
		if ($this.Enabled) {$this.SecurityMethod = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Security/Method')}
		if ($this.Enabled) {$this.SecurityServerCertificate = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Security/ServerCertificate')}
		if ($this.Enabled) {$this.SecurityServerPrivateKey = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Security/ServerPrivateKey')}
		if ($this.Enabled) {$this.SecurityCaCertificate = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Security/CACertificate')}
		if ($this.Enabled) {$this.AudioRateCorrectionMode = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Audio/RateCorrectionMode')}
		if ($this.Enabled) {$this.AudioLogPath = $global:vbox.IVRDEServer_getVRDEProperty($this.Id, 'Audio/LogPath')}
    }
}
Update-TypeData -TypeName IVrdeServer -DefaultDisplayPropertySet @("Enabled","AuthType","AuthTimeout","AuthLibrary") -Force
class GuestProperties {
    [string]$Name
	[string]$Value
	[string]$Flag
	[uint64]$Timestamp
	[DateTimeOffset]$DateTimeOffset
    [array]Enumerate ($IMachine, $ModuleHost) {
        [string[]]$Names = @()
        [string[]]$Values = @()
        [int64[]]$Timestamps = @()
        [string[]]$Flags = @()
        if ($ModuleHost.ToLower() -eq 'websrv') {
            $Names = $global:vbox.IMachine_enumerateGuestProperties($IMachine, $null, [ref]$Values, [ref]$Timestamps, [ref]$Flags)
        }
        elseif ($ModuleHost.ToLower() -eq 'com') {
            $IMachine.EnumerateGuestProperties($null, [ref]$Names, [ref]$Values, [ref]$Timestamps, [ref]$Flags)
        }
        $Variable = [GuestProperties]::new()
        [array]$ret = @()
        for ($i=0;$i-lt($Names | Measure-Object).Count;$i++) {
			$ret += [GuestProperties]@{Name=$Names[$i];Value=$Values[$i];Flag=$Flags[$i];Timestamp=$Timestamps[$i];DateTimeOffset=[DateTimeOffset]::FromUnixTimeMilliseconds($Timestamps[$i] / 1000000)}
        }
        $ret = $ret | Where-Object {$_.Name -ne $null}
        $ret = $ret | Sort-Object Name
        return $ret
    }
}
Update-TypeData -TypeName GuestProperties -DefaultDisplayPropertySet @("Name","Value","Flag") -Force
class ISession {
    [string]$Id
    [System.__ComObject]$Session
}
Update-TypeData -TypeName ISession -DefaultDisplayPropertySet @("Id","Session") -Force
class IStorageControllers {
    [ValidateNotNullOrEmpty()]
    [string]$Name
    [string]$Id
    [uint32]$MaxDevicesPerPortCount
    [uint32]$MinPortCount
    [uint32]$MaxPortCount
    [uint32]$Instance
    [uint32]$PortCount
    [string]$Bus
    [string]$ControllerType
    [bool]$UseHostIOCache
    [bool]$Bootable
    [System.__ComObject]$ComObject
    [IStorageControllers[]]Fetch ([string]$IMachine) {
        $Variable = [IStorageControllers]::new()
        [string[]]$istoragecontrollers = $global:vbox.IMachine_getStorageControllers($IMachine)
        foreach ($istoragecontroller in $istoragecontrollers) {
            $Variable.Id = $istoragecontroller
            $Variable.Name = $global:vbox.IStorageController_getName($istoragecontroller)
            $Variable.MaxDevicesPerPortCount = $global:vbox.IStorageController_getMaxDevicesPerPortCount($istoragecontroller)
            $Variable.MinPortCount = $global:vbox.IStorageController_getMinPortCount($istoragecontroller)
            $Variable.MaxPortCount = $global:vbox.IStorageController_getMaxPortCount($istoragecontroller)
            $Variable.Instance = $global:vbox.IStorageController_getInstance($istoragecontroller)
            $Variable.PortCount = $global:vbox.IStorageController_getPortCount($istoragecontroller)
            $Variable.Bus = $global:vbox.IStorageController_getBus($istoragecontroller)
            $Variable.ControllerType = $global:vbox.IStorageController_getControllerType($istoragecontroller)
            $Variable.UseHostIOCache = $global:vbox.IStorageController_getUseHostIOCache($istoragecontroller)
            $Variable.Bootable = $global:vbox.IStorageController_getBootable($istoragecontroller)
            [IStorageControllers[]]$ret += [IStorageControllers]@{Id=$Variable.Id;Name=$Variable.Name;MaxDevicesPerPortCount=$Variable.MaxDevicesPerPortCount;MinPortCount=$Variable.MinPortCount;MaxPortCount=$Variable.MaxPortCount;Instance=$Variable.Instance;PortCount=$Variable.PortCount;Bus=$Variable.Bus;ControllerType=$Variable.ControllerType;UseHostIOCache=$Variable.UseHostIOCache;Bootable=$Variable.Bootable}
        }
        return $ret
    }
}
if ($ModuleHost.ToLower() -eq 'websrv') {Update-TypeData -TypeName IStorageControllers -DefaultDisplayProperty Name -DefaultDisplayPropertySet @("Name","Bus","ControllerType","UseHostIOCache","Bootable") -Force}
if ($ModuleHost.ToLower() -eq 'com') {Update-TypeData -TypeName IStorageControllers -DefaultDisplayProperty Name -DefaultDisplayPropertySet @("Name","Bus","ControllerType","UseHostIOCache","Bootable","ComObject") -Force}
class IMediumAttachments {
    [VirtualBoxVHD]$IMedium = [VirtualBoxVHD]::new()
    [string]$Controller
    [uint32]$Port
    [uint32]$Device
    [string]$Type
    [bool]$Passthrough
    [bool]$TemporaryEject
    [bool]$IsEjected
    [bool]$NonRotational
    [bool]$Discard
    [bool]$HotPluggable
    [string]$BandwidthGroup
    [string]$MachineId
    [string]$MediumId
    [System.__ComObject]$ComObject
    [IMediumAttachments[]]Fetch ([string]$IMachine) {
        $Variable = [IMediumAttachments]::new()
        $medatts = $global:vbox.IMachine_getMediumAttachments($IMachine)
        foreach ($medatt in $medatts){
			$Variable.MachineId = $medatt.machine
			$Variable.MediumId = $medatt.medium
			$Variable.Controller = $medatt.controller
			$Variable.Port = $medatt.port
			$Variable.Device = $medatt.device
			$Variable.Type = $medatt.type
			$Variable.Passthrough = $medatt.passthrough
			$Variable.TemporaryEject = $medatt.temporaryEject
			$Variable.IsEjected = $medatt.isEjected
			$Variable.NonRotational = $medatt.nonRotational
			$Variable.Discard = $medatt.discard
			$Variable.HotPluggable = $medatt.hotPluggable
			$Variable.BandwidthGroup = $medatt.bandwidthGroup
            [IMediumAttachments[]]$ret += [IMediumAttachments]@{MachineId=$Variable.MachineId;MediumId=$Variable.MediumId;Controller=$Variable.Controller;Port=$Variable.Port;Device=$Variable.Device;Type=$Variable.Type;Passthrough=$Variable.Passthrough;TemporaryEject=$Variable.TemporaryEject;IsEjected=$Variable.IsEjected;NonRotational=$Variable.NonRotational;Discard=$Variable.Discard;HotPluggable=$Variable.HotPluggable;BandwidthGroup=$Variable.BandwidthGroup}
        }
        return $ret
    }
}
if ($ModuleHost.ToLower() -eq 'websrv') {Update-TypeData -TypeName IMediumAttachments -DefaultDisplayProperty Controller -DefaultDisplayPropertySet @("IMedium","Controller","Port","Device","Type") -Force}
if ($ModuleHost.ToLower() -eq 'com') {Update-TypeData -TypeName IMediumAttachments -DefaultDisplayProperty Controller -DefaultDisplayPropertySet @("IMedium","Controller","Port","Device","Type","ComObject") -Force}
class IBiosSettings {
    [bool]$LogoFadeIn
    [bool]$LogoFadeOut
    [uint32]$LogoDisplayTime
    [string]$LogoImagePath
    [string]$BootMenuMode
    [bool]$AcpiEnabled
    [bool]$IoApicEnabled
    [bool]$ApicMode
    [int64]$TimeOffset
    [bool]$PxeDebugEnabled
    [string]$NonVolatileStorageFile
    [bool]$SmBiosUuidLittleEndian
    [string]$Id
    [System.__ComObject]$ComObject
}
if ($ModuleHost.ToLower() -eq 'websrv') {Update-TypeData -TypeName IBiosSettings -DefaultDisplayProperty BootMenuMode -DefaultDisplayPropertySet @("BootMenuMode","AcpiEnabled","IoApicEnabled","ApicMode","TimeOffset") -Force}
if ($ModuleHost.ToLower() -eq 'com') {Update-TypeData -TypeName IBiosSettings -DefaultDisplayProperty BootMenuMode -DefaultDisplayPropertySet @("BootMenuMode","AcpiEnabled","IoApicEnabled","ApicMode","TimeOffset","ComObject") -Force}
# property classes
class VirtualBoxVM {
    [ValidateNotNullOrEmpty()]
    [string]$Name
    [string]$Id
    [string]$MMachine
    [bool]$Accessible
    [string]$AccessError
    [string]$Groups
    [string]$OsTypeId
    [string]$HardwareVersion
    [guid]$HardwareUuid
    [uint32]$CpuCount
    [bool]$CpuHotPlugEnabled
    [ValidateRange(1, 100)]
    [uint32]$CpuExecutionCap
    [uint32]$CpuIdPortabilityLevel
    [uint32]$MemorySize
    [uint32]$MemoryBalloonSize
    [bool]$PageFusionEnabled
    [string]$PointingHidType
    [string]$KeyboardHidType
    [bool]$HpetEnabled
    [string]$ChipsetType
    [string]$SnapshotFolder
    [bool]$EmulatedUsbCardReaderEnabled
    [string]$SettingsFilePath
    [string]$SessionState
    [string]$SessionName
    [uint32]$SessionPid
    [string]$State
    [uint64]$LastStateChange
    [string]$StateFilePath
    [string]$LogFolder
    [uint32]$SnapshotCount
    [bool]$CurrentStateModified
    [string]$ClipboardMode
    [bool]$ClipboardFileTransfersEnabled
    [string]$DnDMode
    [bool]$TeleporterEnabled
    [uint32]$TeleporterPort
    [string]$TeleporterAddress
    [securestring]$TeleporterPassword
    [string]$ParavirtProvider
    [bool]$RtcUseUtc
    [bool]$IoCacheEnabled
    [uint32]$IoCacheSize
    [bool]$TracingEnabled
    [string]$TracingConfig
    [bool]$AllowTracingToAccessVm
    [bool]$AutostartEnabled
    [uint32]$AutostartDelay
    [string]$AutostopType
    [string]$DefaultFrontend
    [bool]$UsbProxyAvailable
    [string]$VmProcessPriority
    [string]$ParavirtDebug
    [string]$CpuProfile
    [guid]$Guid
    [string]$Description
    [string]$MemoryMB
    [bool]$Running
    [string]$Info
    [string]$GuestOS
    [ISession]$ISession = [ISession]::new()
    [string]$MSession
    [string]$IConsole
    [string]$MConsole
    [IProgress]$IProgress = [IProgress]::new()
    [string]$IConsoleGuest
    [string]$IGuestSession
    [IStorageControllers[]]$IStorageControllers = [IStorageControllers]::new()
    [IVrdeServer]$IVrdeServer = [IVrdeServer]::new()
    [array]$GuestProperties
    [IMediumAttachments[]]$IMediumAttachments = [IMediumAttachments]::new()
    [IBiosSettings]$IBiosSettings = [IBiosSettings]::new()
    [System.__ComObject]$ComObject
}
if ($ModuleHost.ToLower() -eq 'websrv') {Update-TypeData -TypeName VirtualBoxVM -DefaultDisplayProperty Name -DefaultDisplayPropertySet @("GUID","Name","MemoryMB","Description","State","GuestOS") -Force}
if ($ModuleHost.ToLower() -eq 'com') {Update-TypeData -TypeName VirtualBoxVM -DefaultDisplayProperty Name -DefaultDisplayPropertySet @("GUID","Name","MemoryMB","Description","State","GuestOS","ComObject") -Force}
class VirtualBoxVHD {
    [string]$Name
    [guid]$Guid
    [string]$Description
    [string]$Format
    [string]$Size
    [string]$LogicalSize
    [string[]]$MachineGuid
    [string[]]$MachineName
    [string]$State
    [string[]]$Variant
    [string]$Location
    [bool]$HostDrive
    [string]$MediumFormat
    [string]$Type
    [string]$Parent
    [string[]]$Children
    [string]$Id
    [IProgress]$IProgress = [IProgress]::new()
    [bool]$ReadOnly
    [bool]$AutoReset
    [string]$LastAccessError
    [System.__ComObject]$ComObject
    static [array]op_Addition($A,$B) {
        [array]$C = $null
        $C += [VirtualBoxVHD]@{Name=$A.Name;Description=$A.Description;Format=$A.Format;Size=$A.Size;LogicalSize=$A.LogicalSize;MachineGuid=$A.MachineGuid;MachineName=$A.MachineName;State=$A.State;Variant=$A.Variant;Location=$A.Location;HostDrive=$A.HostDrive;MediumFormat=$A.MediumFormat;Type=$A.Type;Parent=$A.Parent;Children=$A.Children;Id=$A.Id;ReadOnly=$A.ReadOnly;AutoReset=$A.AutoReset;LastAccessError=$A.LastAccessError}
        $C += [VirtualBoxVHD]@{Name=$B.Name;Description=$B.Description;Format=$B.Format;Size=$B.Size;LogicalSize=$B.LogicalSize;MachineGuid=$B.MachineGuid;MachineName=$B.MachineName;State=$B.State;Variant=$B.Variant;Location=$B.Location;HostDrive=$B.HostDrive;MediumFormat=$B.MediumFormat;Type=$B.Type;Parent=$B.Parent;Children=$B.Children;Id=$B.Id;ReadOnly=$B.ReadOnly;AutoReset=$B.AutoReset;LastAccessError=$B.LastAccessError}
        return $C
    }
}
if ($ModuleHost.ToLower() -eq 'websrv') {Update-TypeData -TypeName VirtualBoxVHD -DefaultDisplayProperty Name -DefaultDisplayPropertySet @("Name","Description","Format","Size","LogicalSize","MachineGuid","MachineName") -Force}
if ($ModuleHost.ToLower() -eq 'com') {Update-TypeData -TypeName VirtualBoxVHD -DefaultDisplayProperty Name -DefaultDisplayPropertySet @("Name","Description","Format","Size","LogicalSize","MachineGuid","MachineName","ComObject") -Force}
class VirtualBoxWebSrvTask {
    [string]$Name
    [string]$Path
    [string]$Status
}
Update-TypeData -TypeName VirtualBoxWebSrvTask -DefaultDisplayPropertySet @("Name","Path","Status") -Force
class MediumFormats {
    [string[]]$Name
    [string[]]$Extensions
    [string[]]$Capabilities
    [string[]]$Id
    Fetch () {
        $Ids = $global:vbox.ISystemProperties_getMediumFormats($global:isystemproperties)
        $this.Id = $Ids
        foreach ($Id in $Ids) {
           	$devicetypevar = New-Object VirtualBox.DeviceType
           	$this.Name += $global:vbox.IMediumFormat_getName($Id)
           	$this.Capabilities += $global:vbox.IMediumFormat_getCapabilities($Id)
           	$this.Extensions += $global:vbox.IMediumFormat_describeFileExtensions($Id, [ref]$devicetypevar)
        }
    }
    FetchCom () {
        foreach ($mediumformat in $global:vbox.SystemProperties.MediumFormats) {
           	$devicetypevar = New-Object DeviceTypeEnum
            $extensionsvar = ''
           	$this.Name += $mediumformat.Name
           	$this.Capabilities += $mediumformat.Capabilities
           	$mediumformat.DescribeFileExtensions([ref]$extensionsvar, [ref]$devicetypevar)
            $this.Extensions += $extensionsvar
        }
    }
    [array]FetchObject ([string[]]$Ids) {
        $ret = New-Object MediumFormats
        foreach ($Id in $Ids) {
            $somevar = New-Object MediumFormats
            $devicetypevar = New-Object VirtualBox.DeviceType
            $somevar.Id = $Id
		    $somevar.Name = $global:vbox.IMediumFormat_getName($Id)
		    $somevar.Capabilities = $global:vbox.IMediumFormat_getCapabilities($Id)
            $somevar.Extensions = @($global:vbox.IMediumFormat_describeFileExtensions($Id, [ref]$devicetypevar))
            [array]$ret += [MediumFormats]@{Id=$somevar.Id;Name=$somevar.Name;Capabilities=$somevar.Capabilities;Extensions=$somevar.Extensions}
        }
        return $ret
    }
    [array]FetchComObject () {
        $ret = New-Object MediumFormats
        foreach ($mediumformat in $global:vbox.SystemProperties.MediumFormats) {
            $somevar = New-Object MediumFormats
            $devicetypevar = New-Object DeviceTypeEnum
            $extensionsvar = ''
		    $somevar.Name = $mediumformat.Name
		    $somevar.Capabilities = $mediumformat.Capabilities
            $mediumformat.DescribeFileExtensions([ref]$extensionsvar, [ref]$devicetypevar)
            $somevar.Extensions = @($extensionsvar)
            [array]$ret += [MediumFormats]@{Name=$somevar.Name;Capabilities=$somevar.Capabilities;Extensions=$somevar.Extensions}
        }
        return $ret
    }
}
Update-TypeData -TypeName MediumFormats -DefaultDisplayPropertySet @("Name","Extensions") -Force
class IVirtualSystemDescription {
	[string[]]$TypeNames = @('OS','Name','Product','Vendor','Version','ProductUrl','VendorUrl','Description','License','Miscellaneous','CPU','Memory','HardDiskControllerIDE','HardDiskControllerSATA','HardDiskControllerSCSI','HardDiskControllerSAS','HardDiskImage','Floppy','CDROM','NetworkAdapter','USBController','SoundCard','CloudInstanceShape','CloudDomain','CloudBootDiskSize','CloudBucket','CloudOCIVCN','CloudPublicIP','CloudProfileName','CloudOCISubnet','CloudKeepObject','CloudLaunchInstance','CloudInstanceId','CloudImageId','CloudInstanceState','CloudImageState','CloudInstanceDisplayName','CloudImageDisplayName','CloudOCILaunchMode','CloudPrivateIP','CloudBootVolumeId','CloudOCIVCNCompartment','CloudOCISubnetCompartment','CloudPublicSSHKey','BootingFirmware',
	'SettingsFile', # Optional, may be unset by the API caller. If this is changed by the API caller it defines the absolute path of the VM settings file and therefore also the VM folder with highest priority.
	'BaseFolder', # Optional, may be unset by the API caller. If set (and SettingsFile is not changed), defines the VM base folder (taking the primary group into account if also set).
	'PrimaryGroup' # Optional, empty by default and may be unset by the API caller. Defines the primary group of the VM after import. May influence the selection of the VM folder. Additional groups may be configured later using IMachine::groups[], after importing.
    )
    [string]$Types
    [string]$Refs
    [string]$OVFValues
    [string]$VBoxValues
    [string]$ExtraConfigValues
    [bool]$Options
    [array]Fetch ([string]$VirtualSystemDescription) {
        $ret = New-Object IVirtualSystemDescription
        [string[]]$outTypes = $()
        [string[]]$outRefs = @()
        [string[]]$outOVFValues = @()
        [string[]]$outVBoxValues = @()
        [string[]]$outExtraConfigValues = @()
        $outTypes = $global:vbox.IVirtualSystemDescription_getDescription($VirtualSystemDescription, [ref]$outRefs, [ref]$outOVFValues, [ref]$outVBoxValues, [ref]$outExtraConfigValues)
        for ($i=0;$i-lt($outTypes | Measure-Object).Count;$i++) {
		    [array]$ret += [IVirtualSystemDescription]@{Types=$outTypes[$i];Refs=$outRefs[$i];OVFValues=$outOVFValues[$i];VBoxValues=$outVBoxValues[$i];ExtraConfigValues=$outExtraConfigValues[$i];Options=$true}
        }
        $ret.TypeNames | Get-Unique | ForEach-Object -Process {
            if ($ret.Types -notcontains $_) {
                $global:vbox.IVirtualSystemDescription_addDescription($VirtualSystemDescription, [string[]]$_, '', '')
                [array]$ret += [IVirtualSystemDescription]@{Types=$_;Refs='';OVFValues='';VBoxValues='';ExtraConfigValues='';Options=$false}
            }
        }
        if ((($ret.Types | Where-Object {$_ -eq 'HardDiskControllerIDE'}) | Measure-Object).Count -eq 1) {
            $global:vbox.IVirtualSystemDescription_addDescription($VirtualSystemDescription, 'HardDiskControllerIDE', '', '')
            [array]$ret += [IVirtualSystemDescription]@{Types='HardDiskControllerIDE';Refs='6';OVFValues='';VBoxValues='';ExtraConfigValues='';Options=$false}
        }
        $ret = $ret | Where-Object {$_.Types -ne $null}
        return $ret
    }
    [array]FetchCom ([System.__ComObject]$VirtualSystemDescription) {
        $ret = New-Object IVirtualSystemDescription
        [int[]]$tempOutTypes = @()
        [string[]]$outRefs = @()
        [string[]]$outOVFValues = @()
        [string[]]$outVBoxValues = @()
        [string[]]$outExtraConfigValues = @()
        $VirtualSystemDescription.GetDescription([ref]$tempOutTypes, [ref]$outRefs, [ref]$outOVFValues, [ref]$outVBoxValues, [ref]$outExtraConfigValues)
        [string[]]$outTypes = @()
        foreach ($outType in $tempOutTypes) {
            $outTypes += [VirtualSystemDescriptionType]::new().ToStr($outType)
        }
        $outTypes = $outTypes | Where-Object {$_ -ne $null}
        for ($i=0;$i-lt($outTypes | Measure-Object).Count;$i++) {
		    [array]$ret += [IVirtualSystemDescription]@{Types=$outTypes[$i];Refs=$outRefs[$i];OVFValues=$outOVFValues[$i];VBoxValues=$outVBoxValues[$i];ExtraConfigValues=$outExtraConfigValues[$i];Options=$true}
        }
        $ret.TypeNames | Get-Unique | ForEach-Object -Process {
            if ($ret.Types -notcontains $_) {
                $VirtualSystemDescription.AddDescription([VirtualSystemDescriptionType]::new().ToInt($_), '', '')
                [array]$ret += [IVirtualSystemDescription]@{Types=$_;Refs='';OVFValues='';VBoxValues='';ExtraConfigValues='';Options=$false}
            }
        }
        if ((($ret.Types | Where-Object {$_ -eq 'HardDiskControllerIDE'}) | Measure-Object).Count -eq 1) {
            $VirtualSystemDescription.AddDescription([VirtualSystemDescriptionType]::new().ToInt('HardDiskControllerIDE'), '', '')
            [array]$ret += [IVirtualSystemDescription]@{Types='HardDiskControllerIDE';Refs='6';OVFValues='';VBoxValues='';ExtraConfigValues='';Options=$false}
        }
        $ret = $ret | Where-Object {$_.Types -ne $null}
        return $ret
    }
}
Update-TypeData -TypeName IVirtualSystemDescription -DefaultDisplayPropertySet @("Types","OVFValues","VBoxValues") -Force
class SystemPropertiesSupported {
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
    Fetch () {
		$this.ParavirtProviders = $global:vbox.ISystemProperties_getSupportedParavirtProviders($global:isystemproperties)
		$this.ClipboardModes = $global:vbox.ISystemProperties_getSupportedClipboardModes($global:isystemproperties)
		$this.DndModes = $global:vbox.ISystemProperties_getSupportedDnDModes($global:isystemproperties)
		$this.FirmwareTypes = $global:vbox.ISystemProperties_getSupportedFirmwareTypes($global:isystemproperties)
		$this.PointingHidTypes = $global:vbox.ISystemProperties_getSupportedPointingHIDTypes($global:isystemproperties)
		$this.KeyboardHidTypes = $global:vbox.ISystemProperties_getSupportedKeyboardHIDTypes($global:isystemproperties)
		$this.VfsTypes = $global:vbox.ISystemProperties_getSupportedVFSTypes($global:isystemproperties)
		$this.ImportOptions = $global:vbox.ISystemProperties_getSupportedImportOptions($global:isystemproperties)
		$this.ExportOptions = $global:vbox.ISystemProperties_getSupportedExportOptions($global:isystemproperties)
		$this.RecordingAudioCodecs = $global:vbox.ISystemProperties_getSupportedRecordingAudioCodecs($global:isystemproperties)
		$this.RecordingVideoCodecs = $global:vbox.ISystemProperties_getSupportedRecordingVideoCodecs($global:isystemproperties)
		$this.RecordingVsMethods = $global:vbox.ISystemProperties_getSupportedRecordingVSMethods($global:isystemproperties)
		$this.RecordingVrcModes = $global:vbox.ISystemProperties_getSupportedRecordingVRCModes($global:isystemproperties)
		$this.GraphicsControllerTypes = $global:vbox.ISystemProperties_getSupportedGraphicsControllerTypes($global:isystemproperties)
		$this.CloneOptions = $global:vbox.ISystemProperties_getSupportedCloneOptions($global:isystemproperties)
		$this.AutostopTypes = $global:vbox.ISystemProperties_getSupportedAutostopTypes($global:isystemproperties)
		$this.VmProcPriorities = $global:vbox.ISystemProperties_getSupportedVMProcPriorities($global:isystemproperties)
		$this.NetworkAttachmentTypes = $global:vbox.ISystemProperties_getSupportedNetworkAttachmentTypes($global:isystemproperties)
		$this.NetworkAdapterTypes = $global:vbox.ISystemProperties_getSupportedNetworkAdapterTypes($global:isystemproperties)
		$this.PortModes = $global:vbox.ISystemProperties_getSupportedPortModes($global:isystemproperties)
		$this.UartTypes = $global:vbox.ISystemProperties_getSupportedUartTypes($global:isystemproperties)
		$this.UsbControllerTypes = $global:vbox.ISystemProperties_getSupportedUSBControllerTypes($global:isystemproperties)
		$this.AudioDriverTypes = $global:vbox.ISystemProperties_getSupportedAudioDriverTypes($global:isystemproperties)
		$this.AudioControllerTypes = $global:vbox.ISystemProperties_getSupportedAudioControllerTypes($global:isystemproperties)
		$this.StorageBuses = $global:vbox.ISystemProperties_getSupportedStorageBuses($global:isystemproperties)
		$this.StorageControllerTypes = $global:vbox.ISystemProperties_getSupportedStorageControllerTypes($global:isystemproperties)
		$this.ChipsetTypes = $global:vbox.ISystemProperties_getSupportedChipsetTypes($global:isystemproperties)
		$this.MinGuestRam = $global:vbox.ISystemProperties_getMinGuestRAM($global:isystemproperties)
		$this.MaxGuestRam = $global:vbox.ISystemProperties_getMaxGuestRAM($global:isystemproperties)
		$this.MinGuestVRam = $global:vbox.ISystemProperties_getMinGuestVRAM($global:isystemproperties)
		$this.MaxGuestVRam = $global:vbox.ISystemProperties_getMaxGuestVRAM($global:isystemproperties)
		$this.MinGuestCPUCount = $global:vbox.ISystemProperties_getMinGuestCPUCount($global:isystemproperties)
		$this.MaxGuestCPUCount = $global:vbox.ISystemProperties_getMaxGuestCPUCount($global:isystemproperties)
    }
    FetchCom () {
		foreach ($SupportedParavirtProvider in $global:vbox.SystemProperties.SupportedParavirtProviders) {$this.ParavirtProviders += [ParavirtProvider]::new().ToStr($SupportedParavirtProvider)}
		foreach ($SupportedClipboardMode in $global:vbox.SystemProperties.SupportedClipboardModes) {$this.ClipboardModes += [ClipboardMode]::new().ToStr($SupportedClipboardMode)}
		foreach ($SupportedDnDMode in $global:vbox.SystemProperties.SupportedDnDModes) {$this.DndModes += [DnDMode]::new().ToStr($SupportedDnDMode)}
		foreach ($SupportedFirmwareType in $global:vbox.SystemProperties.SupportedFirmwareTypes) {$this.FirmwareTypes += [FirmwareType]::new().ToStr($SupportedFirmwareType)}
		foreach ($SupportedPointingHIDType in $global:vbox.SystemProperties.SupportedPointingHIDTypes) {$this.PointingHidTypes += [PointingHIDType]::new().ToStr($SupportedPointingHIDType)}
		foreach ($SupportedKeyboardHIDType in $global:vbox.SystemProperties.SupportedKeyboardHIDTypes) {$this.KeyboardHidTypes += [KeyboardHIDType]::new().ToStr($SupportedKeyboardHIDType)}
		foreach ($SupportedVFSType in $global:vbox.SystemProperties.SupportedVFSTypes) {$this.VfsTypes += [VFSType]::new().ToStr($SupportedVFSType)}
		foreach ($SupportedImportOption in $global:vbox.SystemProperties.SupportedImportOptions) {$this.ImportOptions += [ImportOptions]::new().ToStr($SupportedImportOption)}
		foreach ($SupportedExportOption in $global:vbox.SystemProperties.SupportedExportOptions) {$this.ExportOptions += [ExportOptions]::new().ToStr($SupportedExportOption)}
		foreach ($SupportedRecordingAudioCodec in $global:vbox.SystemProperties.SupportedRecordingAudioCodecs) {$this.RecordingAudioCodecs += [RecordingAudioCodec]::new().ToStr($SupportedRecordingAudioCodec)}
		foreach ($SupportedRecordingVideoCodec in $global:vbox.SystemProperties.SupportedRecordingVideoCodecs) {$this.RecordingVideoCodecs += [RecordingVideoCodec]::new().ToStr($SupportedRecordingVideoCodec)}
		foreach ($SupportedRecordingVSMethod in $global:vbox.SystemProperties.SupportedRecordingVSMethods) {$this.RecordingVsMethods += [RecordingVideoScalingMethod]::new().ToStr($SupportedRecordingVSMethod)}
		foreach ($SupportedRecordingVRCMode in $global:vbox.SystemProperties.SupportedRecordingVRCModes) {$this.RecordingVrcModes += [RecordingVideoRateControlMode]::new().ToStr($SupportedRecordingVRCMode)}
		foreach ($SupportedGraphicsControllerType in $global:vbox.SystemProperties.SupportedGraphicsControllerTypes) {$this.GraphicsControllerTypes += [GraphicsControllerType]::new().ToStr($SupportedGraphicsControllerType)}
		foreach ($SupportedCloneOption in $global:vbox.SystemProperties.SupportedCloneOptions) {$this.CloneOptions += [CloneOptions]::new().ToStr($SupportedCloneOption)}
		foreach ($SupportedAutostopType in $global:vbox.SystemProperties.SupportedAutostopTypes) {$this.AutostopTypes += [AutostopType]::new().ToStr($SupportedAutostopType)}
		foreach ($SupportedVMProcPriority in $global:vbox.SystemProperties.SupportedVMProcPriorities) {$this.VmProcPriorities += [VMProcPriority]::new().ToStr($SupportedVMProcPriority)}
		foreach ($SupportedNetworkAttachmentType in $global:vbox.SystemProperties.SupportedNetworkAttachmentTypes) {$this.NetworkAttachmentTypes += [NetworkAttachmentType]::new().ToStr($SupportedNetworkAttachmentType)}
		foreach ($SupportedNetworkAdapterType in $global:vbox.SystemProperties.SupportedNetworkAdapterTypes) {$this.NetworkAdapterTypes += [NetworkAdapterType]::new().ToStr($SupportedNetworkAdapterType)}
		foreach ($SupportedPortMode in $global:vbox.SystemProperties.SupportedPortModes) {$this.PortModes += [PortMode]::new().ToStr($SupportedPortMode)}
		foreach ($SupportedUartType in $global:vbox.SystemProperties.SupportedUartTypes) {$this.UartTypes += [UartType]::new().ToStr($SupportedUartType)}
		foreach ($SupportedUSBControllerType in $global:vbox.SystemProperties.SupportedUSBControllerTypes) {$this.UsbControllerTypes += [USBControllerType]::new().ToStr($SupportedUSBControllerType)}
		foreach ($SupportedAudioDriverType in $global:vbox.SystemProperties.SupportedAudioDriverTypes) {$this.AudioDriverTypes += [AudioDriverType]::new().ToStr($SupportedAudioDriverType)}
		foreach ($SupportedAudioControllerType in $global:vbox.SystemProperties.SupportedAudioControllerTypes) {$this.AudioControllerTypes += [AudioControllerType]::new().ToStr($SupportedAudioControllerType)}
		foreach ($SupportedStorageBus in $global:vbox.SystemProperties.SupportedStorageBuses) {$this.StorageBuses += [StorageBus]::new().ToStr($SupportedStorageBus)}
		foreach ($SupportedStorageControllerType in $global:vbox.SystemProperties.SupportedStorageControllerTypes) {$this.StorageControllerTypes += [StorageControllerType]::new().ToStr($SupportedStorageControllerType)}
		foreach ($SupportedChipsetType in $global:vbox.SystemProperties.SupportedChipsetTypes) {$this.ChipsetTypes += [ChipsetType]::new().ToStr($SupportedChipsetType)}
		$this.MinGuestRam = $global:vbox.SystemProperties.MinGuestRAM
		$this.MaxGuestRam = $global:vbox.SystemProperties.MaxGuestRAM
		$this.MinGuestVRam = $global:vbox.SystemProperties.MinGuestVRAM
		$this.MaxGuestVRam = $global:vbox.SystemProperties.MaxGuestVRAM
		$this.MinGuestCPUCount = $global:vbox.SystemProperties.MinGuestCPUCount
		$this.MaxGuestCPUCount = $global:vbox.SystemProperties.MaxGuestCPUCount
    }
}
class MediumVariantsSupported {
    [string[]]$Type = @('Standard','VmdkSplit2G','VmdkRawDisk','VmdkStreamOptimized','VmdkESX','VdiZeroExpand')
    [string[]]$Flags = @('Fixed','Diff','Formatted','NoCreateDir')
}
class AccessModesSupported {
    [string[]]$Type = @('ReadOnly','ReadWrite')
}
# method classes - mostly for conversions
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
if (!$__reallyThisIsJustToMakeTheseCollapsableForMe) {
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
    Value is typed �long�, not �result�, to make interface usable from scripting languages.
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
class DeviceType {
    [uint64]ToULong ([string]$FromStr) {
        if ($FromStr){
            $ToULong = $null
            Switch ($FromStr) {
                'Null'         {$ToULong = 0} # Null value, may also mean "no device". ***Note: Not allowed for IConsole_getDeviceActivity()
                'Floppy'       {$ToULong = 1} # Floppy device.
                'DVD'          {$ToULong = 2} # CD/DVD-ROM device.
                'HardDisk'     {$ToULong = 3} # Hard disk device.
                'Network'      {$ToULong = 4} # Network device.
                'USB'          {$ToULong = 5} # USB device.
                'SharedFolder' {$ToULong = 6} # Shared folder device.
                'Graphics3D'   {$ToULong = 7} # Graphics device 3D activity.
                Default        {$ToULong = 0} # Default to 0.
            }
            return [uint64]$ToULong
        }
        else {return $null}
    }
    [string]ToStr ([uint64]$FromLong) {
        if ($FromLong -ne $null){
            $ToStr = $null
            Switch ($FromLong) {
                0       {$ToStr = 'Null'} # Null value, may also mean "no device". ***Note: Not allowed for IConsole_getDeviceActivity()
                1       {$ToStr = 'Floppy'} # Floppy device.
                2       {$ToStr = 'DVD'} # CD/DVD-ROM device.
                3       {$ToStr = 'HardDisk'} # Hard disk device.
                4       {$ToStr = 'Network'} # Network device.
                5       {$ToStr = 'USB'} # USB device.
                6       {$ToStr = 'SharedFolder'} # Shared folder device.
                7       {$ToStr = 'Graphics3D'} # Graphics device 3D activity.
                Default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Unsigned Long
class AccessMode {
    [uint64]ToULong ([string]$FromStr) {
        if ($FromStr){
            $ToULong = $null
            Switch ($FromStr) {
                'ReadOnly'  {$ToULong = 0}
                'ReadWrite' {$ToULong = 1}
                Default     {$ToULong = 0} # Default to 0.
            }
            return [uint64]$ToULong
        }
        else {return $null}
    }
    [string]ToStr ([uint64]$FromLong) {
        if ($FromLong -ne $null){
            $ToStr = $null
            Switch ($FromLong) {
                0       {$ToStr = 'ReadOnly'}
                1       {$ToStr = 'ReadWrite'}
                Default {$ToStr = 'ReadOnly'} # Default to ReadOnly.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Unsigned Long
class CleanupMode {
    [uint64]ToULong ([string]$FromStr) {
        if ($FromStr){
            $ToULong = $null
            Switch ($FromStr) {
                'UnregisterOnly'               {$ToULong = 0} # Unregister only the machine, but neither delete snapshots nor detach media.
                'DetachAllReturnNone'          {$ToULong = 1} # Delete all snapshots and detach all media but return none; this will keep all media registered.
                'DetachAllReturnHardDisksOnly' {$ToULong = 2} # Delete all snapshots, detach all media and return hard disks for closing, but not removable media.
                'Full'                         {$ToULong = 3} # Delete all snapshots, detach all media and return all media for closing.
                Default                        {$ToULong = 0} # Default to 0.
            }
            return [uint64]$ToULong
        }
        else {return $null}
    }
    [string]ToStr ([uint64]$FromLong) {
        if ($FromLong -ne $null){
            $ToStr = $null
            Switch ($FromLong) {
                0       {$ToStr = 'UnregisterOnly'} # Unregister only the machine, but neither delete snapshots nor detach media.
                1       {$ToStr = 'DetachAllReturnNone'} # Delete all snapshots and detach all media but return none; this will keep all media registered.
                2       {$ToStr = 'DetachAllReturnHardDisksOnly'} # Delete all snapshots, detach all media and return hard disks for closing, but not removable media.
                3       {$ToStr = 'Full'} # Delete all snapshots, detach all media and return all media for closing.
                Default {$ToStr = 'UnregisterOnly'} # Default to UnregisterOnly.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Unsigned Long
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
        if ($FromLong -ne $null){
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
class MediumVariant {
    [int32]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToULong = $null
            Switch ($FromStr) {
                'Standard'            {$ToULong = 0} # No particular variant requested, results in using the backend default.
                'VmdkSplit2G'         {$ToULong = 1} # VMDK image split in chunks of less than 2GByte.
                'VmdkRawDisk'         {$ToULong = 2} # VMDK image representing a raw disk.
                'VmdkStreamOptimized' {$ToULong = 3} # VMDK streamOptimized image. Special import/export format which is read-only/append-only.
                'VmdkESX'             {$ToULong = 4} # VMDK format variant used on ESX products.
                'VdiZeroExpand'       {$ToULong = 5} # Fill new blocks with zeroes while expanding image file.
                'Fixed'               {$ToULong = 6} # Fixed image. Only allowed for base images.
                'Diff'                {$ToULong = 7} # Differencing image. Only allowed for child images.
                'Formatted'           {$ToULong = 8} # Special flag which requests formatting the disk image. Right now supported for floppy images only.
                'NoCreateDir'         {$ToULong = 9} # Special flag which suppresses automatic creation of the subdirectory. Only used when passing the medium variant as an input parameter.
                Default               {$ToULong = 0} # Default to 0.
            }
            return [int]$ToULong
        }
        else {return $null}
    }
    [uint64]ToULongCom ([string]$FromStr) {
        if ($FromStr){
            $ToULong = $null
            Switch ($FromStr) {
                'Standard'            {$ToULong = 0} # No particular variant requested, results in using the backend default.
                'VmdkSplit2G'         {$ToULong = 1} # VMDK image split in chunks of less than 2GByte.
                'VmdkRawDisk'         {$ToULong = 2} # VMDK image representing a raw disk.
                'VmdkStreamOptimized' {$ToULong = 4} # VMDK streamOptimized image. Special import/export format which is read-only/append-only.
                'VmdkESX'             {$ToULong = 8} # VMDK format variant used on ESX products.
                'VdiZeroExpand'       {$ToULong = 256} # Fill new blocks with zeroes while expanding image file.
                'Fixed'               {$ToULong = 65536} # Fixed image. Only allowed for base images.
                'Diff'                {$ToULong = 131072} # Differencing image. Only allowed for child images.
                'Formatted'           {$ToULong = 536870912} # Special flag which requests formatting the disk image. Right now supported for floppy images only.
                'NoCreateDir'         {$ToULong = 1073741824} # Special flag which suppresses automatic creation of the subdirectory. Only used when passing the medium variant as an input parameter.
                Default               {$ToULong = 0} # Default to 0.
            }
            return [uint64]$ToULong
        }
        else {return $null}
    }
    [string]ToStr ([int32]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Standard'} # No particular variant requested, results in using the backend default.
                1       {$ToStr = 'VmdkSplit2G'} # VMDK image split in chunks of less than 2GByte.
                2       {$ToStr = 'VmdkRawDisk'} # VMDK image representing a raw disk.
                3       {$ToStr = 'VmdkStreamOptimized'} # VMDK streamOptimized image. Special import/export format which is read-only/append-only.
                4       {$ToStr = 'VmdkESX'} # VMDK format variant used on ESX products.
                5       {$ToStr = 'VdiZeroExpand'} # Fill new blocks with zeroes while expanding image file.
                6       {$ToStr = 'Fixed'} # Fixed image. Only allowed for base images.
                7       {$ToStr = 'Diff'} # Differencing image. Only allowed for child images.
                8       {$ToStr = 'Formatted'} # Special flag which requests formatting the disk image. Right now supported for floppy images only.
                9       {$ToStr = 'NoCreateDir'} # Special flag which suppresses automatic creation of the subdirectory. Only used when passing the medium variant as an input parameter.
                Default {$ToStr = 'Err_Unknown'} # Default to Err_Unknown.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
    [string]ToStrCom ([uint64]$FromULong) {
        if ($FromULong -ne $null){
            $ToStr = $null
            Switch ($FromULong) {
                0          {$ToStr = 'Standard'} # No particular variant requested, results in using the backend default.
                1          {$ToStr = 'VmdkSplit2G'} # VMDK image split in chunks of less than 2GByte.
                2          {$ToStr = 'VmdkRawDisk'} # VMDK image representing a raw disk.
                4          {$ToStr = 'VmdkStreamOptimized'} # VMDK streamOptimized image. Special import/export format which is read-only/append-only.
                8          {$ToStr = 'VmdkESX'} # VMDK format variant used on ESX products.
                256        {$ToStr = 'VdiZeroExpand'} # Fill new blocks with zeroes while expanding image file.
                65536      {$ToStr = 'Fixed'} # Fixed image. Only allowed for base images.
                131072     {$ToStr = 'Diff'} # Differencing image. Only allowed for child images.
                536870912  {$ToStr = 'Formatted'} # Special flag which requests formatting the disk image. Right now supported for floppy images only.
                1073741824 {$ToStr = 'NoCreateDir'} # Special flag which suppresses automatic creation of the subdirectory. Only used when passing the medium variant as an input parameter.
                Default    {$ToStr = 'Err_Unknown'} # Default to Err_Unknown.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class MediumState {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'NotCreated'   {$ToInt = 0} # Associated medium storage does not exist (either was not created yet or was deleted).
                'Created'      {$ToInt = 1} # Associated storage exists and accessible; this gets set if the accessibility check performed by IMedium::refreshState() was successful.
                'LockedRead'   {$ToInt = 2} # Medium is locked for reading (see IMedium::lockRead()), no data modification is possible.
                'LockedWrite'  {$ToInt = 3} # Medium is locked for writing (see IMedium::lockWrite()), no concurrent data reading or modification is possible.
                'Inaccessible' {$ToInt = 4} # Medium accessibility check (see IMedium::refreshState()) has not yet been performed, or else, associated medium storage is not accessible. In the first case, IMedium::lastAccessError is empty, in the second case, it describes the error that occurred.
                'Creating'     {$ToInt = 5} # Associated medium storage is being created.
                'Deleting'     {$ToInt = 6} # Associated medium storage is being deleted.
                Default        {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'NotCreated'} # Associated medium storage does not exist (either was not created yet or was deleted).
                1       {$ToStr = 'Created'} # Associated storage exists and accessible; this gets set if the accessibility check performed by IMedium::refreshState() was successful.
                2       {$ToStr = 'LockedRead'} # Medium is locked for reading (see IMedium::lockRead()), no data modification is possible.
                3       {$ToStr = 'LockedWrite'} # Medium is locked for writing (see IMedium::lockWrite()), no concurrent data reading or modification is possible.
                4       {$ToStr = 'Inaccessible'} # Medium accessibility check (see IMedium::refreshState()) has not yet been performed, or else, associated medium storage is not accessible. In the first case, IMedium::lastAccessError is empty, in the second case, it describes the error that occurred.
                5       {$ToStr = 'Creating'} # Associated medium storage is being created.
                6       {$ToStr = 'Deleting'} # Associated medium storage is being deleted.
                Default {$ToStr = 'NotCreated'} # Default to NotCreated.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class MediumType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Normal'       {$ToInt = 0} # Normal medium (attached directly or indirectly, preserved when taking snapshots).
                'Immutable'    {$ToInt = 1} # Immutable medium (attached indirectly, changes are wiped out the next time the virtual machine is started).
                'Writethrough' {$ToInt = 2} # Write through medium (attached directly, ignored when taking snapshots).
                'Shareable'    {$ToInt = 3} # Allow using this medium concurrently by several machines. **Note: Present since VirtualBox 3.2.0, and accepted since 3.2.8.
                'Readonly'     {$ToInt = 4} # A readonly medium, which can of course be used by several machines. **Note: Present and accepted since VirtualBox 4.0.
                'MultiAttach'  {$ToInt = 5} # A medium which is indirectly attached, so that one base medium can be used for several VMs which have their own differencing medium to store their modifications. In some sense a variant of Immutable with unset AutoReset flag in each differencing medium. **Note: Present and accepted since VirtualBox 4.0.
                Default        {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Normal'} # Normal medium (attached directly or indirectly, preserved when taking snapshots).
                1       {$ToStr = 'Immutable'} # Immutable medium (attached indirectly, changes are wiped out the next time the virtual machine is started).
                2       {$ToStr = 'Writethrough'} # Write through medium (attached directly, ignored when taking snapshots).
                3       {$ToStr = 'Shareable'} # Allow using this medium concurrently by several machines. **Note: Present since VirtualBox 3.2.0, and accepted since 3.2.8.
                4       {$ToStr = 'Readonly'} # A readonly medium, which can of course be used by several machines. **Note: Present and accepted since VirtualBox 4.0.
                5       {$ToStr = 'MultiAttach'} # A medium which is indirectly attached, so that one base medium can be used for several VMs which have their own differencing medium to store their modifications. In some sense a variant of Immutable with unset AutoReset flag in each differencing medium. **Note: Present and accepted since VirtualBox 4.0.
                Default {$ToStr = 'Normal'} # Default to Normal.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
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
        if ($FromInt -ne $null){
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
class AuthType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'     {$ToInt = 0} # Null value, also means �no authentication�.
                'External' {$ToInt = 1}
                'Guest'    {$ToInt = 2}
                Default    {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value, also means �no authentication�.
                1       {$ToStr = 'External'}
                2       {$ToStr = 'Guest'}
                Default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class ImportOptions {
    [int[]]ToInt ([string[]]$FromStrs) {
        if ($FromStrs){
            $ToInts = @()
            foreach ($FromStr in $FromStrs) {
                $ToInt = $null
                Switch ($FromStr) {
                    'KeepAllMACs' {$ToInt = 0} # Don�t generate new MAC addresses of the attached network adapters.
                    'KeepNATMACs' {$ToInt = 1} # Don�t generate new MAC addresses of the attached network adapters when they are using NAT.
                    'ImportToVDI' {$ToInt = 2} # Import all disks to VDI format
                    Default       {$ToInt = 3} # Default to 3.
                }
                $ToInts += $ToInt
            }
            $ToInts = $ToInts | Where-Object {$_ -ne $null}
            return [int[]]$ToInts
        }
        else {return $null}
    }
    [string[]]ToStr ([int[]]$FromInts) {
        if ($FromInts){
            $ToStrs = @()
            foreach ($FromInt in $FromInts) {
                $ToStr = $null
                Switch ($FromInt) {
                    0       {$ToStr = 'KeepAllMACs'} # Don�t generate new MAC addresses of the attached network adapters.
                    1       {$ToStr = 'KeepNATMACs'} # Don�t generate new MAC addresses of the attached network adapters when they are using NAT.
                    2       {$ToStr = 'ImportToVDI'} # Import all disks to VDI format
                    Default {$ToStr = $null} # Default to Null.
                }
                $ToStrs += $ToStr
            }
            $ToStrs = $ToStrs | Where-Object {$_ -ne $null}
            return [string[]]$ToStrs
        }
        else {return $null}
    }
} # Int[]
class DnDMode {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Disabled'      {$ToInt = 0}
                'HostToGuest'   {$ToInt = 1}
                'GuestToHost'   {$ToInt = 2}
                'Bidirectional' {$ToInt = 3}
                Default         {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Disabled'}
                1       {$ToStr = 'HostToGuest'}
                2       {$ToStr = 'GuestToHost'}
                3       {$ToStr = 'Bidirectional'}
                default {$ToStr = 'Disabled'} # Default to Disabled.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class ExportOptions {
    [int[]]ToInt ([string[]]$FromStrs) {
        if ($FromStrs){
            $ToInts = @()
            foreach ($FromStr in $FromStrs) {
                $ToInt = $null
                Switch ($FromStr) {
                    'CreateManifest'     {$ToInt = 0} # Write the optional manifest file (.mf) which is used for integrity checks prior import.
                    'ExportDVDImages'    {$ToInt = 1} # Export DVD images. Default is not to export them as it is rarely needed for typical VMs.
                    'StripAllMACs'       {$ToInt = 2} # Do not export any MAC address information. Default is to keep them to avoid losing information which can cause trouble after import, at the price of risking duplicate MAC addresses, if the import options are used to keep them.
                    'StripAllNonNATMACs' {$ToInt = 3} # Do not export any MAC address information, except for adapters using NAT. Default is to keep them to avoid losing information which can cause trouble after import, at the price of risking duplicate MAC addresses, if the import options are used to keep them.
                    Default              {$ToInt = $null} # Default to Null.
                }
                $ToInts += $ToInt
            }
            $ToInts = $ToInts | Where-Object {$_ -ne $null}
            return [int[]]$ToInts
        }
        else {return $null}
    }
    [string[]]ToStr ([int[]]$FromInts) {
        if ($FromInts){
            $ToStrs = @()
            foreach ($FromInt in $FromInts) {
                $ToStr = $null
                Switch ($FromInt) {
                    0       {$ToStr = 'CreateManifest'} # Write the optional manifest file (.mf) which is used for integrity checks prior import.
                    1       {$ToStr = 'ExportDVDImages'} # Export DVD images. Default is not to export them as it is rarely needed for typical VMs.
                    2       {$ToStr = 'StripAllMACs'} # Do not export any MAC address information. Default is to keep them to avoid losing information which can cause trouble after import, at the price of risking duplicate MAC addresses, if the import options are used to keep them.
                    3       {$ToStr = 'StripAllNonNATMACs'} # Do not export any MAC address information, except for adapters using NAT. Default is to keep them to avoid losing information which can cause trouble after import, at the price of risking duplicate MAC addresses, if the import options are used to keep them.
                    Default {$ToStr = $null} # Default to Null.
                }
                $ToStrs += $ToStr
            }
            $ToStrs = $ToStrs | Where-Object {$_ -ne $null}
            return [string[]]$ToStrs
        }
        else {return $null}
    }
} # Int[]
class ProcessCreateFlag {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'None'                    {$ToInt = 0} # No flag set
                'WaitForProcessStartOnly' {$ToInt = 1} # Only use the specified timeout value to wait for starting the guest process - the guest process itself then uses an infinite timeout.
                'IgnoreOrphanedProcesses' {$ToInt = 2} # Do not report an error when executed processes are still alive when VBoxService or the guest OS is shutting down.
                'Hidden'                  {$ToInt = 3} # Do not show the started process according to the guest OS guidelines.
                'Profile'                 {$ToInt = 4} # Utilize the user�s profile data when exeuting a process. Only available for Windows guests at the moment.
                'WaitForStdOut'           {$ToInt = 5} # The guest process waits until all data from stdout is read out.
                'WaitForStdErr'           {$ToInt = 6} # The guest process waits until all data from stderr is read out.
                'ExpandArguments'         {$ToInt = 7} # Expands environment variables in process arguments. ***Note: This is not yet implemented and is currently silently ignored. We will document the protocolVersion number for this feature once it appears, so don�t use it till then.
                'UnquotedArguments'       {$ToInt = 8} # Work around for Windows and OS/2 applications not following normal argument quoting and escaping rules. The arguments are passed to the application without any extra quoting, just a single space between each. ***Note: Present since VirtualBox 4.3.28 and 5.0 beta 3.
                Default                   {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [int]ToIntCom ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'None'                    {$ToInt = 0} # No flag set
                'WaitForProcessStartOnly' {$ToInt = 1} # Only use the specified timeout value to wait for starting the guest process - the guest process itself then uses an infinite timeout.
                'IgnoreOrphanedProcesses' {$ToInt = 2} # Do not report an error when executed processes are still alive when VBoxService or the guest OS is shutting down.
                'Hidden'                  {$ToInt = 4} # Do not show the started process according to the guest OS guidelines.
                'Profile'                 {$ToInt = 8} # Utilize the user�s profile data when exeuting a process. Only available for Windows guests at the moment.
                'WaitForStdOut'           {$ToInt = 16} # The guest process waits until all data from stdout is read out.
                'WaitForStdErr'           {$ToInt = 32} # The guest process waits until all data from stderr is read out.
                'ExpandArguments'         {$ToInt = 64} # Expands environment variables in process arguments. ***Note: This is not yet implemented and is currently silently ignored. We will document the protocolVersion number for this feature once it appears, so don�t use it till then.
                'UnquotedArguments'       {$ToInt = 128} # Work around for Windows and OS/2 applications not following normal argument quoting and escaping rules. The arguments are passed to the application without any extra quoting, just a single space between each. ***Note: Present since VirtualBox 4.3.28 and 5.0 beta 3.
                Default                   {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'None'} # No flag set
                1       {$ToStr = 'WaitForProcessStartOnly'} # Only use the specified timeout value to wait for starting the guest process - the guest process itself then uses an infinite timeout.
                2       {$ToStr = 'IgnoreOrphanedProcesses'} # Do not report an error when executed processes are still alive when VBoxService or the guest OS is shutting down.
                3       {$ToStr = 'Hidden'} # Do not show the started process according to the guest OS guidelines.
                4       {$ToStr = 'Profile'} # Utilize the user�s profile data when exeuting a process. Only available for Windows guests at the moment.
                5       {$ToStr = 'WaitForStdOut'} # The guest process waits until all data from stdout is read out.
                6       {$ToStr = 'WaitForStdErr'} # The guest process waits until all data from stderr is read out.
                7       {$ToStr = 'ExpandArguments'} # Expands environment variables in process arguments. ***Note: This is not yet implemented and is currently silently ignored. We will document the protocolVersion number for this feature once it appears, so don�t use it till then.
                8       {$ToStr = 'UnquotedArguments'} # Work around for Windows and OS/2 applications not following normal argument quoting and escaping rules. The arguments are passed to the application without any extra quoting, just a single space between each. ***Note: Present since VirtualBox 4.3.28 and 5.0 beta 3.
                Default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
    [string]ToStrCom ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'None'} # No flag set
                1       {$ToStr = 'WaitForProcessStartOnly'} # Only use the specified timeout value to wait for starting the guest process - the guest process itself then uses an infinite timeout.
                2       {$ToStr = 'IgnoreOrphanedProcesses'} # Do not report an error when executed processes are still alive when VBoxService or the guest OS is shutting down.
                4       {$ToStr = 'Hidden'} # Do not show the started process according to the guest OS guidelines.
                8       {$ToStr = 'Profile'} # Utilize the user�s profile data when exeuting a process. Only available for Windows guests at the moment.
                16      {$ToStr = 'WaitForStdOut'} # The guest process waits until all data from stdout is read out.
                32      {$ToStr = 'WaitForStdErr'} # The guest process waits until all data from stderr is read out.
                64      {$ToStr = 'ExpandArguments'} # Expands environment variables in process arguments. ***Note: This is not yet implemented and is currently silently ignored. We will document the protocolVersion number for this feature once it appears, so don�t use it till then.
                128     {$ToStr = 'UnquotedArguments'} # Work around for Windows and OS/2 applications not following normal argument quoting and escaping rules. The arguments are passed to the application without any extra quoting, just a single space between each. ***Note: Present since VirtualBox 4.3.28 and 5.0 beta 3.
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
        if ($FromLong -ne $null){
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
class StorageBus {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'       {$ToInt = 0} # Null value. Never used by the API.
                'IDE'        {$ToInt = 1}
                'SATA'       {$ToInt = 2}
                'SCSI'       {$ToInt = 3}
                'Floppy'     {$ToInt = 4}
                'SAS'        {$ToInt = 5}
                'USB'        {$ToInt = 6}
                'PCIe'       {$ToInt = 7}
                'VirtioSCSI' {$ToInt = 8}
                Default      {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value. Never used by the API.
                1       {$ToStr = 'IDE'}
                2       {$ToStr = 'SATA'}
                3       {$ToStr = 'SCSI'}
                4       {$ToStr = 'Floppy'}
                5       {$ToStr = 'SAS'}
                6       {$ToStr = 'USB'}
                7       {$ToStr = 'PCIe'}
                8       {$ToStr = 'VirtioSCSI'}
                Default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class StorageControllerType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'        {$ToInt = 0} # Null value. Never used by the API.
                'LsiLogic'    {$ToInt = 1} # A SCSI controller of the LsiLogic variant.
                'BusLogic'    {$ToInt = 2} # A SCSI controller of the BusLogic variant.
                'IntelAhci'   {$ToInt = 3} # An Intel AHCI SATA controller; this is the only variant for SATA.
                'PIIX3'       {$ToInt = 4} # An IDE controller of the PIIX3 variant.
                'PIIX4'       {$ToInt = 5} # An IDE controller of the PIIX4 variant.
                'ICH6'        {$ToInt = 6} # An IDE controller of the ICH6 variant.
                'I82078'      {$ToInt = 7} # A floppy disk controller; this is the only variant for floppy drives.
                'LsiLogicSas' {$ToInt = 8} # A variant of the LsiLogic controller using SAS.
                'USB'         {$ToInt = 9} # Special USB based storage controller.
                'NVMe'        {$ToInt = 10} # An NVMe storage controller.
                'VirtioSCSI'  {$ToInt = 11} # Virtio SCSI storage controller.
                Default       {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value. Never used by the API.
                1       {$ToStr = 'LsiLogic'} # A SCSI controller of the LsiLogic variant.
                2       {$ToStr = 'BusLogic'} # A SCSI controller of the BusLogic variant.
                3       {$ToStr = 'IntelAhci'} # An Intel AHCI SATA controller; this is the only variant for SATA.
                4       {$ToStr = 'PIIX3'} # An IDE controller of the PIIX3 variant.
                5       {$ToStr = 'PIIX4'} # An IDE controller of the PIIX4 variant.
                6       {$ToStr = 'ICH6'} # An IDE controller of the ICH6 variant.
                7       {$ToStr = 'I82078'} # A floppy disk controller; this is the only variant for floppy drives.
                8       {$ToStr = 'LsiLogicSas'} # A variant of the LsiLogic controller using SAS.
                9       {$ToStr = 'USB'} # Special USB based storage controller.
                10      {$ToStr = 'NVMe'} # An NVMe storage controller.
                11      {$ToStr = 'VirtioSCSI'} # Virtio SCSI storage controller.
                Default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
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
        if ($FromInt -ne $null){
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
        if ($FromLong -ne $null){
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
class VirtualSystemDescriptionType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Ignore'                    {$ToInt = 0}
                'OS'                        {$ToInt = 1}
                'Name'                      {$ToInt = 2}
                'Product'                   {$ToInt = 3}
                'Vendor'                    {$ToInt = 4}
                'Version'                   {$ToInt = 5}
                'ProductUrl'                {$ToInt = 6}
                'VendorUrl'                 {$ToInt = 7}
                'Description'               {$ToInt = 8}
                'License'                   {$ToInt = 9}
                'Miscellaneous'             {$ToInt = 10}
                'CPU'                       {$ToInt = 11}
                'Memory'                    {$ToInt = 12}
                'HardDiskControllerIDE'     {$ToInt = 13}
                'HardDiskControllerSATA'    {$ToInt = 14}
                'HardDiskControllerSCSI'    {$ToInt = 15}
                'HardDiskControllerSAS'     {$ToInt = 16}
                'HardDiskImage'             {$ToInt = 17}
                'Floppy'                    {$ToInt = 18}
                'CDROM'                     {$ToInt = 19}
                'NetworkAdapter'            {$ToInt = 20}
                'USBController'             {$ToInt = 21}
                'SoundCard'                 {$ToInt = 22}
                'SettingsFile'              {$ToInt = 23} # Optional, may be unset by the API caller. If this is changed by the API caller it defines the absolute path of the VM settings file and therefore also the VM folder with highest priority.
                'BaseFolder'                {$ToInt = 24} # Optional, may be unset by the API caller. If set (and SettingsFile is not changed), defines the VM base folder (taking the primary group into account if also set).
                'PrimaryGroup'              {$ToInt = 25} # Optional, empty by default and may be unset by the API caller. Defines the primary group of the VM after import. May influence the selection of the VM folder. Additional groups may be configured later using IMachine::groups[], after importing.
                'CloudInstanceShape'        {$ToInt = 26}
                'CloudDomain'               {$ToInt = 27}
                'CloudBootDiskSize'         {$ToInt = 28}
                'CloudBucket'               {$ToInt = 29}
                'CloudOCIVCN'               {$ToInt = 30}
                'CloudPublicIP'             {$ToInt = 31}
                'CloudProfileName'          {$ToInt = 32}
                'CloudOCISubnet'            {$ToInt = 33}
                'CloudKeepObject'           {$ToInt = 34}
                'CloudLaunchInstance'       {$ToInt = 35}
                'CloudInstanceId'           {$ToInt = 36}
                'CloudImageId'              {$ToInt = 37}
                'CloudInstanceState'        {$ToInt = 38}
                'CloudImageState'           {$ToInt = 39}
                'CloudInstanceDisplayName'  {$ToInt = 40}
                'CloudImageDisplayName'     {$ToInt = 41}
                'CloudOCILaunchMode'        {$ToInt = 42}
                'CloudPrivateIP'            {$ToInt = 43}
                'CloudBootVolumeId'         {$ToInt = 44}
                'CloudOCIVCNCompartment'    {$ToInt = 45}
                'CloudOCISubnetCompartment' {$ToInt = 46}
                'CloudPublicSSHKey'         {$ToInt = 47}
                'BootingFirmware'           {$ToInt = 48}
                Default                     {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Ignore'}
                1       {$ToStr = 'OS'}
                2       {$ToStr = 'Name'}
                3       {$ToStr = 'Product'}
                4       {$ToStr = 'Vendor'}
                5       {$ToStr = 'Version'}
                6       {$ToStr = 'ProductUrl'}
                7       {$ToStr = 'VendorUrl'}
                8       {$ToStr = 'Description'}
                9       {$ToStr = 'License'}
                10      {$ToStr = 'Miscellaneous'}
                11      {$ToStr = 'CPU'}
                12      {$ToStr = 'Memory'}
                13      {$ToStr = 'HardDiskControllerIDE'}
                14      {$ToStr = 'HardDiskControllerSATA'}
                15      {$ToStr = 'HardDiskControllerSCSI'}
                16      {$ToStr = 'HardDiskControllerSAS'}
                17      {$ToStr = 'HardDiskImage'}
                18      {$ToStr = 'Floppy'}
                19      {$ToStr = 'CDROM'}
                20      {$ToStr = 'NetworkAdapter'}
                21      {$ToStr = 'USBController'}
                22      {$ToStr = 'SoundCard'}
                23      {$ToStr = 'SettingsFile'} # Optional, may be unset by the API caller. If this is changed by the API caller it defines the absolute path of the VM settings file and therefore also the VM folder with highest priority.
                24      {$ToStr = 'BaseFolder'} # Optional, may be unset by the API caller. If set (and SettingsFile is not changed), defines the VM base folder (taking the primary group into account if also set).
                25      {$ToStr = 'PrimaryGroup'} # Optional, empty by default and may be unset by the API caller. Defines the primary group of the VM after import. May influence the selection of the VM folder. Additional groups may be configured later using IMachine::groups[], after importing.
                26      {$ToStr = 'CloudInstanceShape'}
                27      {$ToStr = 'CloudDomain'}
                28      {$ToStr = 'CloudBootDiskSize'}
                29      {$ToStr = 'CloudBucket'}
                30      {$ToStr = 'CloudOCIVCN'}
                31      {$ToStr = 'CloudPublicIP'}
                32      {$ToStr = 'CloudProfileName'}
                33      {$ToStr = 'CloudOCISubnet'}
                34      {$ToStr = 'CloudKeepObject'}
                35      {$ToStr = 'CloudLaunchInstance'}
                36      {$ToStr = 'CloudInstanceId'}
                37      {$ToStr = 'CloudImageId'}
                38      {$ToStr = 'CloudInstanceState'}
                39      {$ToStr = 'CloudImageState'}
                40      {$ToStr = 'CloudInstanceDisplayName'}
                41      {$ToStr = 'CloudImageDisplayName'}
                42      {$ToStr = 'CloudOCILaunchMode'}
                43      {$ToStr = 'CloudPrivateIP'}
                44      {$ToStr = 'CloudBootVolumeId'}
                45      {$ToStr = 'CloudOCIVCNCompartment'}
                46      {$ToStr = 'CloudOCISubnetCompartment'}
                47      {$ToStr = 'CloudPublicSSHKey'}
                48      {$ToStr = 'BootingFirmware'}
                default {$ToStr = 'Ignore'} # Default to Invalid.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class ParavirtProvider {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'None'    {$ToInt = 0} # No provider is used.
                'Default' {$ToInt = 1} # A default provider is automatically chosen according to the guest OS type.
                'Legacy'  {$ToInt = 2} # Used for VMs which didn�t used to have any provider settings. Usually interpreted as None for most VMs.
                'Minimal' {$ToInt = 3} # A minimal set of features to expose to the paravirtualized guest.
                'HyperV'  {$ToInt = 4} # Microsoft Hyper-V.
                'KVM'     {$ToInt = 5} # Linux KVM.
                Default   {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'None'} # No provider is used.
                1       {$ToStr = 'Default'} # A default provider is automatically chosen according to the guest OS type.
                2       {$ToStr = 'Legacy'} # Used for VMs which didn�t used to have any provider settings. Usually interpreted as None for most VMs.
                3       {$ToStr = 'Minimal'} # A minimal set of features to expose to the paravirtualized guest.
                4       {$ToStr = 'HyperV'} # Microsoft Hyper-V.
                5       {$ToStr = 'KVM'} # Linux KVM.
                default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class ClipboardMode {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Disabled'      {$ToInt = 0}
                'HostToGuest'   {$ToInt = 1}
                'GuestToHost'   {$ToInt = 2}
                'Bidirectional' {$ToInt = 3}
                Default         {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Disabled'}
                1       {$ToStr = 'HostToGuest'}
                2       {$ToStr = 'GuestToHost'}
                3       {$ToStr = 'Bidirectional'}
                default {$ToStr = 'Disabled'} # Default to Disabled.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class FirmwareType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'BIOS'    {$ToInt = 0} # BIOS Firmware.
                'EFI'     {$ToInt = 1} # EFI Firmware, bitness detected basing on OS type.
                'EFI32'   {$ToInt = 2} # EFI firmware, 32-bit.
                'EFI64'   {$ToInt = 3} # EFI firmware, 64-bit.
                'EFIDUAL' {$ToInt = 4} # EFI firmware, combined 32 and 64-bit.
                Default   {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'BIOS'} # BIOS Firmware.
                1       {$ToStr = 'EFI'} # EFI Firmware, bitness detected basing on OS type.
                2       {$ToStr = 'EFI32'} # EFI firmware, 32-bit.
                3       {$ToStr = 'EFI64'} # EFI firmware, 64-bit.
                4       {$ToStr = 'EFIDUAL'} # EFI firmware, combined 32 and 64-bit.
                default {$ToStr = 'BIOS'} # Default to BIOS.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class PointingHIDType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'None'          {$ToInt = 0} # No mouse.
                'PS2Mouse'      {$ToInt = 1} # PS/2 auxiliary device, a.k.a. mouse.
                'USBMouse'      {$ToInt = 2} # USB mouse (relative pointer).
                'USBTablet'     {$ToInt = 3} # USB tablet (absolute pointer). Also enables a relative USB mouse in addition.
                'ComboMouse'    {$ToInt = 4} # Combined device, working as PS/2 or USB mouse, depending on guest behavior. Using this device can have negative performance implications.
                'USBMultiTouch' {$ToInt = 5} # USB multi-touch device. Also enables the USB tablet and mouse devices.
                Default         {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'None'} # No mouse.
                1       {$ToStr = 'PS2Mouse'} # PS/2 auxiliary device, a.k.a. mouse.
                2       {$ToStr = 'USBMouse'} # USB mouse (relative pointer).
                3       {$ToStr = 'USBTablet'} # USB tablet (absolute pointer). Also enables a relative USB mouse in addition.
                4       {$ToStr = 'ComboMouse'} # Combined device, working as PS/2 or USB mouse, depending on guest behavior. Using this device can have negative performance implications.
                5       {$ToStr = 'USBMultiTouch'} # USB multi-touch device. Also enables the USB tablet and mouse devices.
                default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class KeyboardHIDType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'None'          {$ToInt = 0} # No keyboard.
                'PS2Keyboard'   {$ToInt = 1} # PS/2 keyboard.
                'USBKeyboard'   {$ToInt = 2} # USB keyboard.
                'ComboKeyboard' {$ToInt = 3} # Combined device, working as PS/2 or USB keyboard, depending on guest behavior. Using of such device can have negative performance implications.
                Default         {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'None'} # No keyboard.
                1       {$ToStr = 'PS2Keyboard'} # PS/2 keyboard.
                2       {$ToStr = 'USBKeyboard'} # USB keyboard.
                3       {$ToStr = 'ComboKeyboard'} # Combined device, working as PS/2 or USB keyboard, depending on guest behavior. Using of such device can have negative performance implications.
                default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class VFSType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'File'   {$ToInt = 0}
                'Cloud'  {$ToInt = 1}
                'S3'     {$ToInt = 2}
                'WebDav' {$ToInt = 3}
                Default  {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'File'}
                1       {$ToStr = 'Cloud'}
                2       {$ToStr = 'S3'}
                3       {$ToStr = 'WebDav'}
                default {$ToStr = 'File'} # Default to File.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class RecordingAudioCodec {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'None'   {$ToInt = 0} # No codec set.
                'WavPCM' {$ToInt = 1} # WAV format, linear PCM, uncompressed. Not implemented yet.
                'Opus'   {$ToInt = 2} # Opus Audio.
                Default  {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'None'} # No codec set.
                1       {$ToStr = 'WavPCM'} # WAV format, linear PCM, uncompressed. Not implemented yet.
                2       {$ToStr = 'Opus'} # Opus Audio.
                default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class RecordingVideoCodec {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'None'  {$ToInt = 0} # No codec set.
                'VP8'   {$ToInt = 1} # VP8 codec.
                'VP9'   {$ToInt = 2} # VP9 codec. Not implemented yet.
                'AV1'   {$ToInt = 3} # AV1 codec. Not implemented yet.
                Default {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'None'} # No codec set.
                1       {$ToStr = 'VP8'} # VP8 codec.
                2       {$ToStr = 'VP9'} # VP9 codec. Not implemented yet.
                3       {$ToStr = 'AV1'} # AV1 codec. Not implemented yet.
                default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class RecordingVideoScalingMethod {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'None'            {$ToInt = 0} # No scaling performed.
                'NearestNeighbor' {$ToInt = 1} # Performs scaling via nearest-neighbor interpolation. Not yet implemented.
                'Bilinear'        {$ToInt = 2} # Performs scaling via bilinear interpolation. Not yet implemented.
                'Bicubic'         {$ToInt = 3} # Performs scaling via bicubic interpolation. Not yet implemented.
                Default           {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'None'} # No scaling performed.
                1       {$ToStr = 'NearestNeighbor'} # Performs scaling via nearest-neighbor interpolation. Not yet implemented.
                2       {$ToStr = 'Bilinear'} # Performs scaling via bilinear interpolation. Not yet implemented.
                3       {$ToStr = 'Bicubic'} # Performs scaling via bicubic interpolation. Not yet implemented.
                default {$ToStr = 'None'} # Default to None.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class RecordingVideoRateControlMode {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'CBR'   {$ToInt = 0} # Constant bit rate (CBR).
                'VBR'   {$ToInt = 1} # Variable bit rate (VBR). Not yet implemented.
                Default {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'CBR'} # Constant bit rate (CBR).
                1       {$ToStr = 'VBR'} # Variable bit rate (VBR). Not yet implemented.
                default {$ToStr = 'CBR'} # Default to CBR.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class GraphicsControllerType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'     {$ToInt = 0} # Reserved value, invalid.
                'VBoxVGA'  {$ToInt = 1} # VirtualBox VGA device.
                'VMSVGA'   {$ToInt = 2} # VMware SVGA II device.
                'VBoxSVGA' {$ToInt = 3} # VirtualBox VGA device with VMware SVGA II extensions.
                Default    {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Reserved value, invalid.
                1       {$ToStr = 'VBoxVGA'} # VirtualBox VGA device.
                2       {$ToStr = 'VMSVGA'} # VMware SVGA II device.
                3       {$ToStr = 'VBoxSVGA'} # VirtualBox VGA device with VMware SVGA II extensions.
                default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class CloneOptions {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Link'          {$ToInt = 0} # Create a clone VM where all virtual disks are linked to the original VM.
                'KeepAllMACs'   {$ToInt = 1} # Don�t generate new MAC addresses of the attached network adapters.
                'KeepNATMACs'   {$ToInt = 2} # Don�t generate new MAC addresses of the attached network adapters when they are using NAT.
                'KeepDiskNames' {$ToInt = 3} # Don�t change the disk names.
                'KeepHwUUIDs'   {$ToInt = 4} # Don�t change UUID of the machine hardware.
                Default         {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Link'} # Create a clone VM where all virtual disks are linked to the original VM.
                1       {$ToStr = 'KeepAllMACs'} # Don�t generate new MAC addresses of the attached network adapters.
                2       {$ToStr = 'KeepNATMACs'} # Don�t generate new MAC addresses of the attached network adapters when they are using NAT.
                3       {$ToStr = 'KeepDiskNames'} # Don�t change the disk names.
                4       {$ToStr = 'KeepHwUUIDs'} # Don�t change UUID of the machine hardware.
                default {$ToStr = 'Link'} # Default to Link.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class AutostopType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Disabled'     {$ToInt = 0} # Stopping the VM during system shutdown is disabled.
                'SaveState'    {$ToInt = 1} # The state of the VM will be saved when the system shuts down.
                'PowerOff'     {$ToInt = 2} # The VM is powered off when the system shuts down.
                'AcpiShutdown' {$ToInt = 3} # An ACPI shutdown event is generated.
                Default        {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Disabled'} # Stopping the VM during system shutdown is disabled.
                1       {$ToStr = 'SaveState'} # The state of the VM will be saved when the system shuts down.
                2       {$ToStr = 'PowerOff'} # The VM is powered off when the system shuts down.
                3       {$ToStr = 'AcpiShutdown'} # An ACPI shutdown event is generated.
                default {$ToStr = 'Disabled'} # Default to Disabled.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class VMProcPriority {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Invalid' {$ToInt = 0} # Invalid priority, do not use.
                'Default' {$ToInt = 1} # Default process priority determined by the OS.
                'Flat'    {$ToInt = 2} # Assumes a scheduling policy which puts the process at the default priority and with all thread at the same priority.
                'Low'     {$ToInt = 3} # Assumes a scheduling policy which puts the process mostly below the default priority of the host OS.
                'Normal'  {$ToInt = 4} # Assume a scheduling policy which shares the CPU resources fairly with other processes running with the default priority of the host OS.
                'High'    {$ToInt = 5} # Assumes a scheduling policy which puts the task above the default priority of the host OS. This policy might easily cause other tasks in the system to starve.
                Default   {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Invalid'} # Invalid priority, do not use.
                1       {$ToStr = 'Default'} # Default process priority determined by the OS.
                2       {$ToStr = 'Flat'} # Assumes a scheduling policy which puts the process at the default priority and with all thread at the same priority.
                3       {$ToStr = 'Low'} # Assumes a scheduling policy which puts the process mostly below the default priority of the host OS.
                4       {$ToStr = 'Normal'} # Assume a scheduling policy which shares the CPU resources fairly with other processes running with the default priority of the host OS.
                5       {$ToStr = 'High'} # Assumes a scheduling policy which puts the task above the default priority of the host OS. This policy might easily cause other tasks in the system to starve.
                default {$ToStr = 'Invalid'} # Default to Invalid.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class NetworkAttachmentType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'       {$ToInt = 0} # Null value, also means �not attached�.
                'NAT'        {$ToInt = 1}
                'Bridged'    {$ToInt = 2}
                'Internal'   {$ToInt = 3}
                'HostOnly'   {$ToInt = 4}
                'Generic'    {$ToInt = 5}
                'NATNetwork' {$ToInt = 6}
                'Cloud'      {$ToInt = 7}
                Default      {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value, also means �not attached�.
                1       {$ToStr = 'NAT'}
                2       {$ToStr = 'Bridged'}
                3       {$ToStr = 'Internal'}
                4       {$ToStr = 'HostOnly'}
                5       {$ToStr = 'Generic'}
                6       {$ToStr = 'NATNetwork'}
                7       {$ToStr = 'Cloud'}
                default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class NetworkAdapterType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'      {$ToInt = 0} # Null value (never used by the API).
                'Am79C970A' {$ToInt = 1} # AMD PCNet-PCI II network card (Am79C970A).
                'Am79C973'  {$ToInt = 2} # AMD PCNet-FAST III network card (Am79C973).
                'I82540EM'  {$ToInt = 3} # Intel PRO/1000 MT Desktop network card (82540EM).
                'I82543GC'  {$ToInt = 4} # Intel PRO/1000 T Server network card (82543GC).
                'I82545EM'  {$ToInt = 5} # Intel PRO/1000 MT Server network card (82545EM).
                'Virtio'    {$ToInt = 6} # Virtio network device.
                'Am79C960'  {$ToInt = 7} # AMD PCnet-ISA/NE2100 network card (Am79C960).
                Default     {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value (never used by the API).
                1       {$ToStr = 'Am79C970A'} # AMD PCNet-PCI II network card (Am79C970A).
                2       {$ToStr = 'Am79C973'} # AMD PCNet-FAST III network card (Am79C973).
                3       {$ToStr = 'I82540EM'} # Intel PRO/1000 MT Desktop network card (82540EM).
                4       {$ToStr = 'I82543GC'} # Intel PRO/1000 T Server network card (82543GC).
                5       {$ToStr = 'I82545EM'} # Intel PRO/1000 MT Server network card (82545EM).
                6       {$ToStr = 'Virtio'} # Virtio network device.
                7       {$ToStr = 'Am79C960'} # AMD PCnet-ISA/NE2100 network card (Am79C960).
                default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class PortMode {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Disconnected' {$ToInt = 0} # Virtual device is not attached to any real host device.
                'HostPipe'     {$ToInt = 1} # Virtual device is attached to a host pipe.
                'HostDevice'   {$ToInt = 2} # Virtual device is attached to a host device.
                'RawFile'      {$ToInt = 3} # Virtual device is attached to a raw file.
                'TCP'          {$ToInt = 4} # Virtual device is attached to a TCP socket.
                Default        {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Disconnected'} # Virtual device is not attached to any real host device.
                1       {$ToStr = 'HostPipe'} # Virtual device is attached to a host pipe.
                2       {$ToStr = 'HostDevice'} # Virtual device is attached to a host device.
                3       {$ToStr = 'RawFile'} # Virtual device is attached to a raw file.
                4       {$ToStr = 'TCP'} # Virtual device is attached to a TCP socket.
                default {$ToStr = 'Disconnected'} # Default to Disconnected.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class UartType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'U16450'  {$ToInt = 0} # The most basic emulated UART which doesn�t support FIFO operation.
                'U16550A' {$ToInt = 1} # The successor of the 16450 UART introducing a 16 byte FIFO to reduce operational overhead.
                'U16750'  {$ToInt = 2} # This UART developed by Texas Instruments introduced a 64 byte FIFO and hardware flow control.
                Default   {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'U16450'} # The most basic emulated UART which doesn�t support FIFO operation.
                1       {$ToStr = 'U16550A'} # The successor of the 16450 UART introducing a 16 byte FIFO to reduce operational overhead.
                2       {$ToStr = 'U16750'} # This UART developed by Texas Instruments introduced a 64 byte FIFO and hardware flow control.
                default {$ToStr = 'U16450'} # Default to U16450.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class USBControllerType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'  {$ToInt = 0} # Null value. Never used by the API.
                'OHCI'  {$ToInt = 1}
                'EHCI'  {$ToInt = 2}
                'XHCI'  {$ToInt = 3}
                'Last'  {$ToInt = 4} # Last element (invalid). Used for parameter checks.
                Default {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value. Never used by the API.
                1       {$ToStr = 'OHCI'}
                2       {$ToStr = 'EHCI'}
                3       {$ToStr = 'XHCI'}
                4       {$ToStr = 'Last'} # Last element (invalid). Used for parameter checks.
                default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class AudioDriverType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'        {$ToInt = 0} # Null value, also means �dummy audio driver�.
                'WinMM'       {$ToInt = 1} # Windows multimedia (Windows hosts only, not supported at the moment).
                'OSS'         {$ToInt = 2} # Open Sound System (Linux / Unix hosts only).
                'ALSA'        {$ToInt = 3} # Advanced Linux Sound Architecture (Linux hosts only).
                'DirectSound' {$ToInt = 4} # DirectSound (Windows hosts only).
                'CoreAudio'   {$ToInt = 5} # CoreAudio (Mac hosts only).
                'MMPM'        {$ToInt = 6} # Reserved for historical reasons.
                'Pulse'       {$ToInt = 7} # PulseAudio (Linux hosts only).
                'SolAudio'    {$ToInt = 8} # Solaris audio (Solaris hosts only, not supported at the moment).
                Default       {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value, also means �dummy audio driver�.
                1       {$ToStr = 'WinMM'} # Windows multimedia (Windows hosts only, not supported at the moment).
                2       {$ToStr = 'OSS'} # Open Sound System (Linux / Unix hosts only).
                3       {$ToStr = 'ALSA'} # Advanced Linux Sound Architecture (Linux hosts only).
                4       {$ToStr = 'DirectSound'} # DirectSound (Windows hosts only).
                5       {$ToStr = 'CoreAudio'} # CoreAudio (Mac hosts only).
                6       {$ToStr = 'MMPM'} # Reserved for historical reasons.
                7       {$ToStr = 'Pulse'} # PulseAudio (Linux hosts only).
                8       {$ToStr = 'SolAudio'} # Solaris audio (Solaris hosts only, not supported at the moment).
                default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class AudioControllerType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'AC97'  {$ToInt = 0}
                'SB16'  {$ToInt = 1}
                'HDA'   {$ToInt = 2}
                Default {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'AC97'}
                1       {$ToStr = 'SB16'}
                2       {$ToStr = 'HDA'}
                default {$ToStr = 'AC97'} # Default to AC97.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class ChipsetType {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'  {$ToInt = 0} # Null value. Never used by the API.
                'PIIX3' {$ToInt = 1} # A PIIX3 (PCI IDE ISA Xcelerator) chipset.
                'ICH9'  {$ToInt = 2} # A ICH9 (I/O Controller Hub) chipset.
                Default {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value. Never used by the API.
                1       {$ToStr = 'PIIX3'} # A PIIX3 (PCI IDE ISA Xcelerator) chipset.
                2       {$ToStr = 'ICH9'} # A ICH9 (I/O Controller Hub) chipset.
                default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class BIOSBootMenuMode {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Disabled'       {$ToInt = 0}
                'MenuOnly'       {$ToInt = 1}
                'MessageAndMenu' {$ToInt = 2}
                Default          {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Disabled'}
                1       {$ToStr = 'MenuOnly'}
                2       {$ToStr = 'MessageAndMenu'}
                default {$ToStr = 'Disabled'} # Default to Disabled.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class SessionState {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'      {$ToInt = 0} # Null value. Never used by the API.
                'Unlocked'  {$ToInt = 1} # In IMachine::sessionState, this means that the machine is not locked for any sessions. In ISession::state, this means that no machine is currently locked for this session.
                'Locked'    {$ToInt = 2} # In IMachine::sessionState, this means that the machine is currently locked for a session, whose process identifier can then be found in the IMachine::sessionPID attribute. In ISession::state, this means that a machine is currently locked for this session, and the mutable machine object can be found in the ISession::machine attribute (see IMachine::lockMachine() for details).
                'Spawning'  {$ToInt = 3} # A new process is being spawned for the machine as a result of IMachine::launchVMProcess() call. This state also occurs as a short transient state during an IMachine::lockMachine() call.
                'Unlocking' {$ToInt = 4} # The session is being unlocked.
                Default     {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value. Never used by the API.
                1       {$ToStr = 'Unlocked'} # In IMachine::sessionState, this means that the machine is not locked for any sessions. In ISession::state, this means that no machine is currently locked for this session.
                2       {$ToStr = 'Locked'} # In IMachine::sessionState, this means that the machine is currently locked for a session, whose process identifier can then be found in the IMachine::sessionPID attribute. In ISession::state, this means that a machine is currently locked for this session, and the mutable machine object can be found in the ISession::machine attribute (see IMachine::lockMachine() for details).
                3       {$ToStr = 'Spawning'} # A new process is being spawned for the machine as a result of IMachine::launchVMProcess() call. This state also occurs as a short transient state during an IMachine::lockMachine() call.
                4       {$ToStr = 'Unlocking'} # The session is being unlocked.
                default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
class MachineState {
    [int]ToInt ([string]$FromStr) {
        if ($FromStr){
            $ToInt = $null
            Switch ($FromStr) {
                'Null'                   {$ToInt = 0} # Null value. Never used by the API.
                'PoweredOff'             {$ToInt = 1} # The machine is not running and has no saved execution state; it has either never been started or been shut down successfully.
                'Saved'                  {$ToInt = 2} # The machine is not currently running, but the execution state of the machine has been saved to an external file when it was running, from where it can be resumed.
                'Teleported'             {$ToInt = 3} # The machine was teleported to a different host (or process) and then powered off. Take care when powering it on again may corrupt resources it shares with the teleportation target (e.g. disk and network).
                'Aborted'                {$ToInt = 4} # The process running the machine has terminated abnormally. This may indicate a crash of the VM process in host execution context, or the VM process has been terminated externally.
                'Running'                {$ToInt = 5} # The machine is currently being executed.
                'Paused'                 {$ToInt = 6} # Execution of the machine has been paused.
                'Stuck'                  {$ToInt = 7} # Execution of the machine has reached the �Guru Meditation� condition. This indicates a severe error in the hypervisor itself.
                'Teleporting'            {$ToInt = 8} # The machine is about to be teleported to a different host or process. It is possible to pause a machine in this state, but it will go to the TeleportingPausedVM state and it will not be possible to resume it again unless the teleportation fails.
                'LiveSnapshotting'       {$ToInt = 9} # A live snapshot is being taken. The machine is running normally, but some of the runtime configuration options are inaccessible. Also, if paused while in this state it will transition to OnlineSnapshotting and it will not be resume the execution until the snapshot operation has completed.
                'Starting'               {$ToInt = 10} # Machine is being started after powering it on from a zero execution state.
                'Stopping'               {$ToInt = 11} # Machine is being normally stopped powering it off, or after the guest OS has initiated a shutdown sequence.
                'Saving'                 {$ToInt = 12} # Machine is saving its execution state to a file.
                'Restoring'              {$ToInt = 13} # Execution state of the machine is being restored from a file after powering it on from the saved execution state.
                'TeleportingPausedVM'    {$ToInt = 14} # The machine is being teleported to another host or process, but it is not running. This is the paused variant of the Teleporting state.
                'TeleportingIn'          {$ToInt = 15} # Teleporting the machine state in from another host or process.
                'DeletingSnapshotOnline' {$ToInt = 16} # Like DeletingSnapshot, but the merging of media is ongoing in the background while the machine is running.
                'DeletingSnapshotPaused' {$ToInt = 17} # Like DeletingSnapshotOnline, but the machine was paused when the merging of differencing media was started.
                'OnlineSnapshotting'     {$ToInt = 18} # Like LiveSnapshotting, but the machine was paused when the merging of differencing media was started.
                'RestoringSnapshot'      {$ToInt = 19} # A machine snapshot is being restored; this typically does not take long.
                'DeletingSnapshot'       {$ToInt = 20} # A machine snapshot is being deleted; this can take a long time since this may require merging differencing media. This value indicates that the machine is not running while the snapshot is being deleted.
                'SettingUp'              {$ToInt = 21} # Lengthy setup operation is in progress.
                'Snapshotting'           {$ToInt = 22} # Taking an (offline) snapshot.
                'FirstOnline'            {$ToInt = 23} # Pseudo-state: first online state (for use in relational expressions).
                'LastOnline'             {$ToInt = 24} # Pseudo-state: last online state (for use in relational expressions).
                'FirstTransient'         {$ToInt = 25} # Pseudo-state: first transient state (for use in relational expressions).
                'LastTransient'          {$ToInt = 26} # Pseudo-state: last transient state (for use in relational expressions).
                Default                  {$ToInt = 0} # Default to 0.
            }
            return [int]$ToInt
        }
        else {return $null}
    }
    [string]ToStr ([int]$FromInt) {
        if ($FromInt -ne $null){
            $ToStr = $null
            Switch ($FromInt) {
                0       {$ToStr = 'Null'} # Null value. Never used by the API.
                1       {$ToStr = 'PoweredOff'} # The machine is not running and has no saved execution state; it has either never been started or been shut down successfully.
                2       {$ToStr = 'Saved'} # The machine is not currently running, but the execution state of the machine has been saved to an external file when it was running, from where it can be resumed.
                3       {$ToStr = 'Teleported'} # The machine was teleported to a different host (or process) and then powered off. Take care when powering it on again may corrupt resources it shares with the teleportation target (e.g. disk and network).
                4       {$ToStr = 'Aborted'} # The process running the machine has terminated abnormally. This may indicate a crash of the VM process in host execution context, or the VM process has been terminated externally.
                5       {$ToStr = 'Running'} # The machine is currently being executed.
                6       {$ToStr = 'Paused'} # Execution of the machine has been paused.
                7       {$ToStr = 'Stuck'} # Execution of the machine has reached the �Guru Meditation� condition. This indicates a severe error in the hypervisor itself.
                8       {$ToStr = 'Teleporting'} # The machine is about to be teleported to a different host or process. It is possible to pause a machine in this state, but it will go to the TeleportingPausedVM state and it will not be possible to resume it again unless the teleportation fails.
                9       {$ToStr = 'LiveSnapshotting'} # A live snapshot is being taken. The machine is running normally, but some of the runtime configuration options are inaccessible. Also, if paused while in this state it will transition to OnlineSnapshotting and it will not be resume the execution until the snapshot operation has completed.
                10      {$ToStr = 'Starting'} # Machine is being started after powering it on from a zero execution state.
                11      {$ToStr = 'Stopping'} # Machine is being normally stopped powering it off, or after the guest OS has initiated a shutdown sequence.
                12      {$ToStr = 'Saving'} # Machine is saving its execution state to a file.
                13      {$ToStr = 'Restoring'} # Execution state of the machine is being restored from a file after powering it on from the saved execution state.
                14      {$ToStr = 'TeleportingPausedVM'} # The machine is being teleported to another host or process, but it is not running. This is the paused variant of the Teleporting state.
                15      {$ToStr = 'TeleportingIn'} # Teleporting the machine state in from another host or process.
                16      {$ToStr = 'DeletingSnapshotOnline'} # Like DeletingSnapshot, but the merging of media is ongoing in the background while the machine is running.
                17      {$ToStr = 'DeletingSnapshotPaused'} # Like DeletingSnapshotOnline, but the machine was paused when the merging of differencing media was started.
                18      {$ToStr = 'OnlineSnapshotting'} # Like LiveSnapshotting, but the machine was paused when the merging of differencing media was started.
                19      {$ToStr = 'RestoringSnapshot'} # A machine snapshot is being restored; this typically does not take long.
                20      {$ToStr = 'DeletingSnapshot'} # A machine snapshot is being deleted; this can take a long time since this may require merging differencing media. This value indicates that the machine is not running while the snapshot is being deleted.
                21      {$ToStr = 'SettingUp'} # Lengthy setup operation is in progress.
                22      {$ToStr = 'Snapshotting'} # Taking an (offline) snapshot.
                23      {$ToStr = 'FirstOnline'} # Pseudo-state: first online state (for use in relational expressions).
                24      {$ToStr = 'LastOnline'} # Pseudo-state: last online state (for use in relational expressions).
                25      {$ToStr = 'FirstTransient'} # Pseudo-state: first transient state (for use in relational expressions).
                26      {$ToStr = 'LastTransient'} # Pseudo-state: last transient state (for use in relational expressions).
                default {$ToStr = 'Null'} # Default to Null.
            }
            return [string]$ToStr
        }
        else {return $null}
    }
} # Int
}
#########################################################################################
# Variable Declarations
$authtype = "VBoxAuth"
$vboxwebsrvtask = New-Object VirtualBoxWebSrvTask
# probably going to drop this in a future version - see the IVirtualBoxErrorInfo class for replacement
$vboxerror = New-Object VirtualBoxError
# unseal globals
if ($mediumformats) {if ((Get-Variable -Name mediumformats -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name mediumformats -Option None -Scope Global -Force}}
if ($mediumformatspso) {if ((Get-Variable -Name mediumformatspso -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name mediumformatspso -Option None -Scope Global -Force}}
if ($systempropertiessupported) {if ((Get-Variable -Name systempropertiessupported -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name systempropertiessupported -Option None -Scope Global -Force}}
if ($mediumvariantssupported) {if ((Get-Variable -Name mediumvariantssupported -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name mediumvariantssupported -Option None -Scope Global -Force}}
if ($accessmodessupported) {if ((Get-Variable -Name accessmodessupported -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name accessmodessupported -Option None -Scope Global -Force}}
# initialize globals
$global:mediumformats = New-Object MediumFormats
$global:mediumformatspso = New-Object MediumFormats
$global:systempropertiessupported = New-Object SystemPropertiesSupported
$global:mediumvariantssupported = New-Object MediumVariantsSupported
$global:accessmodessupported = New-Object AccessModesSupported
# seal globals
if ((Get-Variable -Name mediumformats -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name mediumformats -Description "A list of medium formats." -Option ReadOnly -Scope Global -Visibility Public -Force}
if ((Get-Variable -Name mediumformatspso -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name mediumformatspso -Description "A list of medium formats in PowerShell-readable format." -Option ReadOnly -Scope Global -Visibility Public -Force}
if ((Get-Variable -Name systempropertiessupported -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name systempropertiessupported -Description "A list of VirtualBox properties supported by your system." -Option ReadOnly -Scope Global -Visibility Public -Force}
if ((Get-Variable -Name mediumvariantssupported -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name mediumvariantssupported -Description "A list of medium variants supported by your system." -Option ReadOnly -Scope Global -Visibility Public -Force}
if ((Get-Variable -Name accessmodessupported -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name accessmodessupported -Description "A list of access modes supported by your system." -Option ReadOnly -Scope Global -Visibility Public -Force}
# global automatic variables for conversion
if ($ivirtualboxerrorinfo) {if ((Get-Variable -Name ivirtualboxerrorinfo -Scope Global).Options -match'ReadOnly') {Set-Variable -Name ivirtualboxerrorinfo -Description "Automatic variable used to get VirtualBox error information." -Option None -Scope Global -Force}}
$global:ivirtualboxerrorinfo = New-Object IVirtualBoxErrorInfo
if ((Get-Variable -Name ivirtualboxerrorinfo -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name ivirtualboxerrorinfo -Description "Automatic variable used to get VirtualBox error information." -Option ReadOnly -Scope Global -Force}
#########################################################################################
# Includes
# N/A
#########################################################################################
# Function Definitions
Function Get-VirtualBoxVM {
<#
.SYNOPSIS
Get VirtualBox virtual machine information
.DESCRIPTION
Retrieves any or all VirtualBox virtual machines by name/GUID, state, or all. The default usage, without any parameters is to display all virtual machines.
.PARAMETER Name
The name of at lease one virtual machine.
.PARAMETER Guid
The GUID of at lease one virtual machine.
.PARAMETER State
Return virtual machines based on their state. Valid values are:
"PoweredOff","Running","Saved","Teleported","Aborted","Paused","Stuck","Snapshotting",
"Starting","Stopping","Restoring","TeleportingPausedVM","TeleportingIn","FaultTolerantSync",
"DeletingSnapshotOnline","DeletingSnapshot", and "SettingUp"
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
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
VERSION     :  1.4
LAST UPDATED:  2/17/2020
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
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 if (!$Name) {$All = $true}
} # Begin
Process {
 Write-Verbose "Getting virtual machine inventory"
 # initialize array object to hold virtual machine values
 $vminventory = @()
 try {
  # get virtual machine inventory
  if ($ModuleHost.ToLower() -eq 'websrv') {
   foreach ($vmid in ($global:vbox.IVirtualBox_getMachines($global:ivbox))) {
     $guestprops = New-Object GuestProperties
     $storagecontrollers = New-Object IStorageControllers
     $mediumattachments = New-Object IMediumAttachments
     $tempobj = New-Object VirtualBoxVM
     $tempobj.Name = $global:vbox.IMachine_getName($vmid)
     $tempobj.Description = $global:vbox.IMachine_getDescription($vmid)
     $tempobj.State = $global:vbox.IMachine_getState($vmid)
     $tempobj.GuestOS = $global:vbox.IMachine_getOSTypeId($vmid)
     $tempobj.MemoryMb = $global:vbox.IMachine_getMemorySize($vmid)
     $tempobj.Accessible = $global:vbox.IMachine_getAccessible($vmid)
     $tempobj.AccessError = $global:vbox.IMachine_getAccessError($vmid)
     $tempobj.Groups = $global:vbox.IMachine_getGroups($vmid)
     $tempobj.OsTypeId = $global:vbox.IMachine_getOSTypeId($vmid)
     $tempobj.HardwareVersion = $global:vbox.IMachine_getHardwareVersion($vmid)
     $tempobj.HardwareUuid = $global:vbox.IMachine_getHardwareUUID($vmid)
     $tempobj.CpuCount = $global:vbox.IMachine_getCPUCount($vmid)
     $tempobj.CpuHotPlugEnabled = $global:vbox.IMachine_getCPUHotPlugEnabled($vmid)
     $tempobj.CpuExecutionCap = $global:vbox.IMachine_getCPUExecutionCap($vmid)
     $tempobj.CpuIdPortabilityLevel = $global:vbox.IMachine_getCPUIDPortabilityLevel($vmid)
     $tempobj.MemorySize = $global:vbox.IMachine_getMemorySize($vmid)
     $tempobj.MemoryBalloonSize = $global:vbox.IMachine_getMemoryBalloonSize($vmid)
     $tempobj.PageFusionEnabled = $global:vbox.IMachine_getPageFusionEnabled($vmid)
     $tempobj.PointingHidType = $global:vbox.IMachine_getPointingHIDType($vmid)
     $tempobj.KeyboardHidType = $global:vbox.IMachine_getKeyboardHIDType($vmid)
     $tempobj.HpetEnabled = $global:vbox.IMachine_getHPETEnabled($vmid)
     $tempobj.ChipsetType = $global:vbox.IMachine_getChipsetType($vmid)
     $tempobj.SnapshotFolder = $global:vbox.IMachine_getSnapshotFolder($vmid)
     $tempobj.EmulatedUsbCardReaderEnabled = $global:vbox.IMachine_getEmulatedUSBCardReaderEnabled($vmid)
     $tempobj.SettingsFilePath = $global:vbox.IMachine_getSettingsFilePath($vmid)
     $tempobj.SessionState = $global:vbox.IMachine_getSessionState($vmid)
     $tempobj.LastStateChange = $global:vbox.IMachine_getLastStateChange($vmid)
     $tempobj.StateFilePath = $global:vbox.IMachine_getStateFilePath($vmid)
     $tempobj.LogFolder = $global:vbox.IMachine_getLogFolder($vmid)
     $tempobj.SnapshotCount = $global:vbox.IMachine_getSnapshotCount($vmid)
     $tempobj.CurrentStateModified = $global:vbox.IMachine_getCurrentStateModified($vmid)
     $tempobj.ClipboardMode = $global:vbox.IMachine_getClipboardMode($vmid)
     $tempobj.ClipboardFileTransfersEnabled = $global:vbox.IMachine_getClipboardFileTransfersEnabled($vmid)
     $tempobj.DnDMode = $global:vbox.IMachine_getDnDMode($vmid)
     $tempobj.TeleporterEnabled = $global:vbox.IMachine_getTeleporterEnabled($vmid)
     $tempobj.TeleporterPort = $global:vbox.IMachine_getTeleporterPort($vmid)
     $tempobj.TeleporterAddress = $global:vbox.IMachine_getTeleporterAddress($vmid)
     if ($global:vbox.IMachine_getTeleporterPassword($vmid) -ne '' -and $global:vbox.IMachine_getTeleporterPassword($vmid) -ne $null) {$tempobj.TeleporterPassword = ConvertTo-SecureString -String ($global:vbox.IMachine_getTeleporterPassword($vmid)) -AsPlainText -Force}
     $tempobj.ParavirtProvider = $global:vbox.IMachine_getParavirtProvider($vmid)
     $tempobj.RtcUseUtc = $global:vbox.IMachine_getRTCUseUTC($vmid)
     $tempobj.IoCacheEnabled = $global:vbox.IMachine_getIOCacheEnabled($vmid)
     $tempobj.IoCacheSize = $global:vbox.IMachine_getIOCacheSize($vmid)
     $tempobj.TracingEnabled = $global:vbox.IMachine_getTracingEnabled($vmid)
     $tempobj.TracingConfig = $global:vbox.IMachine_getTracingConfig($vmid)
     $tempobj.AllowTracingToAccessVm = $global:vbox.IMachine_getAllowTracingToAccessVM($vmid)
     $tempobj.AutostartEnabled = $global:vbox.IMachine_getAutostartEnabled($vmid)
     $tempobj.AutostartDelay = $global:vbox.IMachine_getAutostartDelay($vmid)
     $tempobj.AutostopType = $global:vbox.IMachine_getAutostopType($vmid)
     $tempobj.DefaultFrontend = $global:vbox.IMachine_getDefaultFrontend($vmid)
     $tempobj.UsbProxyAvailable = $global:vbox.IMachine_getUSBProxyAvailable($vmid)
     $tempobj.VmProcessPriority = $global:vbox.IMachine_getVMProcessPriority($vmid)
     $tempobj.ParavirtDebug = $global:vbox.IMachine_getParavirtDebug($vmid)
     $tempobj.CpuProfile = $global:vbox.IMachine_getCpuProfile($vmid)
     $tempobj.Id = $vmid
     $tempobj.Guid = $global:vbox.IMachine_getId($vmid)
     $tempobj.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($vmid)
     $tempobj.IVrdeServer = $tempobj.IVrdeServer.Fetch($tempobj.Id)
     $tempobj.GuestProperties = $guestprops.Enumerate($tempobj.Id, $ModuleHost)
     $tempobj.IStorageControllers = $storagecontrollers.Fetch($tempobj.Id)
     $tempobj.IMediumAttachments = $mediumattachments.Fetch($tempobj.Id)
     foreach ($mediumattachment in $tempobj.IMediumAttachments) {
      if ($mediumattachment.Type -eq 'HardDisk') {
       foreach ($imediumid in ($global:vbox.IVirtualBox_getHardDisks($global:ivbox))) {
        if ($imediumid -eq $mediumattachment.MediumId) {
         Write-Verbose "Getting attached disk: $($global:vbox.IMedium_getName($imediumid))"
         $mediumattachment.IMedium.Name = $global:vbox.IMedium_getName($imediumid)
         $mediumattachment.IMedium.GUID = $global:vbox.IMedium_getId($imediumid)
         $mediumattachment.IMedium.Description = $global:vbox.IMedium_getDescription($imediumid)
         $mediumattachment.IMedium.Format = $global:vbox.IMedium_getFormat($imediumid)
         $mediumattachment.IMedium.Size = $global:vbox.IMedium_getSize($imediumid)
         $mediumattachment.IMedium.LogicalSize = $global:vbox.IMedium_getLogicalSize($imediumid)
         $mediumattachment.IMedium.MachineGuid = $global:vbox.IMedium_getMachineIds($imediumid)
         foreach ($machineid in $mediumattachment.IMedium.MachineGuid) {
          foreach ($imachine in ($global:vbox.IVirtualBox_getMachines($global:ivbox))) {
           if (($global:vbox.IMachine_getId($imachine)) -eq $machineid) {
            $mediumattachment.IMedium.MachineName += $global:vbox.IMachine_getName($imachine)
           } # end if ($global:vbox.IMachine_getId($imachine)) -eq $machineid
           $mediumattachment.IMedium.MachineName = $mediumattachment.IMedium.MachineName | Where-Object {$_ -ne $null}
          } # foreach $imachine in ($global:vbox.IVirtualBox_getMachines($global:ivbox))
         } # foreach $machineid in $mediumattachment.IMedium.MachineGuid
         $mediumattachment.IMedium.State = $global:vbox.IMedium_getState($imediumid)
         $mediumattachment.IMedium.Variant = $global:vbox.IMedium_getVariant($imediumid)
         $mediumattachment.IMedium.Location = $global:vbox.IMedium_getLocation($imediumid)
         $mediumattachment.IMedium.HostDrive = $global:vbox.IMedium_getHostDrive($imediumid)
         $mediumattachment.IMedium.MediumFormat = $global:vbox.IMedium_getMediumFormat($imediumid)
         $mediumattachment.IMedium.Type = $global:vbox.IMedium_getType($imediumid)
         $mediumattachment.IMedium.Parent = $global:vbox.IMedium_getParent($imediumid)
         $mediumattachment.IMedium.Children = $global:vbox.IMedium_getChildren($imediumid)
         $mediumattachment.IMedium.Id = $imediumid
         $mediumattachment.IMedium.ReadOnly = $global:vbox.IMedium_getReadOnly($imediumid)
         $mediumattachment.IMedium.AutoReset = $global:vbox.IMedium_getAutoReset($imediumid)
         $mediumattachment.IMedium.LastAccessError = $global:vbox.IMedium_getLastAccessError($imediumid)
        } # end if $imediumid -eq $mediumattachment.MediumId
       } # end foreach loop disk inventory
      } # end if $mediumattachment.Type -match 'HardDisk'
      if ($mediumattachment.Type -eq 'DVD') {
       foreach ($imediumid in ($global:vbox.IVirtualBox_getDVDImages($global:ivbox))) {
        if ($imediumid -eq $mediumattachment.MediumId) {
         Write-Verbose "Getting attached disc: $($global:vbox.IMedium_getName($imediumid))"
         $mediumattachment.IMedium.Name = $global:vbox.IMedium_getName($imediumid)
         $mediumattachment.IMedium.GUID = $global:vbox.IMedium_getId($imediumid)
         $mediumattachment.IMedium.Description = $global:vbox.IMedium_getDescription($imediumid)
         $mediumattachment.IMedium.Format = $global:vbox.IMedium_getFormat($imediumid)
         $mediumattachment.IMedium.Size = $global:vbox.IMedium_getSize($imediumid)
         $mediumattachment.IMedium.LogicalSize = $global:vbox.IMedium_getLogicalSize($imediumid)
         $mediumattachment.IMedium.MachineGuid = $global:vbox.IMedium_getMachineIds($imediumid)
         foreach ($machineid in $mediumattachment.IMedium.MachineGuid) {
          foreach ($imachine in ($global:vbox.IVirtualBox_getMachines($global:ivbox))) {
           if (($global:vbox.IMachine_getId($imachine)) -eq $machineid) {
            $mediumattachment.IMedium.MachineName += $global:vbox.IMachine_getName($imachine)
           } # end if ($global:vbox.IMachine_getId($imachine)) -eq $machineid
           $mediumattachment.IMedium.MachineName = $mediumattachment.IMedium.MachineName | Where-Object {$_ -ne $null}
          } # foreach $imachine in ($global:vbox.IVirtualBox_getMachines($global:ivbox))
         } # foreach $machineid in $mediumattachment.IMedium.MachineGuid
         $mediumattachment.IMedium.State = $global:vbox.IMedium_getState($imediumid)
         $mediumattachment.IMedium.Variant = $global:vbox.IMedium_getVariant($imediumid)
         $mediumattachment.IMedium.Location = $global:vbox.IMedium_getLocation($imediumid)
         $mediumattachment.IMedium.HostDrive = $global:vbox.IMedium_getHostDrive($imediumid)
         $mediumattachment.IMedium.MediumFormat = $global:vbox.IMedium_getMediumFormat($imediumid)
         $mediumattachment.IMedium.Type = $global:vbox.IMedium_getType($imediumid)
         $mediumattachment.IMedium.Parent = $global:vbox.IMedium_getParent($imediumid)
         $mediumattachment.IMedium.Children = $global:vbox.IMedium_getChildren($imediumid)
         $mediumattachment.IMedium.Id = $imediumid
         $mediumattachment.IMedium.ReadOnly = $global:vbox.IMedium_getReadOnly($imediumid)
         $mediumattachment.IMedium.AutoReset = $global:vbox.IMedium_getAutoReset($imediumid)
         $mediumattachment.IMedium.LastAccessError = $global:vbox.IMedium_getLastAccessError($imediumid)
        } # end if $imediumid -eq $mediumattachment.MediumId
       } # end foreach loop disk inventory
      } # end if $mediumattachment.Type -match 'DVD'
      if ($mediumattachment.Type -eq 'Floppy') {
       foreach ($imediumid in ($global:vbox.IVirtualBox_getFloppyImages($global:ivbox))) {
        if ($imediumid -eq $mediumattachment.MediumId) {
         Write-Verbose "Getting attached floppy: $($global:vbox.IMedium_getName($imediumid))"
         $mediumattachment.IMedium.Name = $global:vbox.IMedium_getName($imediumid)
         $mediumattachment.IMedium.GUID = $global:vbox.IMedium_getId($imediumid)
         $mediumattachment.IMedium.Description = $global:vbox.IMedium_getDescription($imediumid)
         $mediumattachment.IMedium.Format = $global:vbox.IMedium_getFormat($imediumid)
         $mediumattachment.IMedium.Size = $global:vbox.IMedium_getSize($imediumid)
         $mediumattachment.IMedium.LogicalSize = $global:vbox.IMedium_getLogicalSize($imediumid)
         $mediumattachment.IMedium.MachineGuid = $global:vbox.IMedium_getMachineIds($imediumid)
         foreach ($machineid in $mediumattachment.IMedium.MachineGuid) {
          foreach ($imachine in ($global:vbox.IVirtualBox_getMachines($global:ivbox))) {
           if (($global:vbox.IMachine_getId($imachine)) -eq $machineid) {
            $mediumattachment.IMedium.MachineName += $global:vbox.IMachine_getName($imachine)
           } # end if ($global:vbox.IMachine_getId($imachine)) -eq $machineid
           $mediumattachment.IMedium.MachineName = $mediumattachment.IMedium.MachineName | Where-Object {$_ -ne $null}
          } # foreach $imachine in ($global:vbox.IVirtualBox_getMachines($global:ivbox))
         } # foreach $machineid in $mediumattachment.IMedium.MachineGuid
         $mediumattachment.IMedium.State = $global:vbox.IMedium_getState($imediumid)
         $mediumattachment.IMedium.Variant = $global:vbox.IMedium_getVariant($imediumid)
         $mediumattachment.IMedium.Location = $global:vbox.IMedium_getLocation($imediumid)
         $mediumattachment.IMedium.HostDrive = $global:vbox.IMedium_getHostDrive($imediumid)
         $mediumattachment.IMedium.MediumFormat = $global:vbox.IMedium_getMediumFormat($imediumid)
         $mediumattachment.IMedium.Type = $global:vbox.IMedium_getType($imediumid)
         $mediumattachment.IMedium.Parent = $global:vbox.IMedium_getParent($imediumid)
         $mediumattachment.IMedium.Children = $global:vbox.IMedium_getChildren($imediumid)
         $mediumattachment.IMedium.Id = $imediumid
         $mediumattachment.IMedium.ReadOnly = $global:vbox.IMedium_getReadOnly($imediumid)
         $mediumattachment.IMedium.AutoReset = $global:vbox.IMedium_getAutoReset($imediumid)
         $mediumattachment.IMedium.LastAccessError = $global:vbox.IMedium_getLastAccessError($imediumid)
        } # end if $imediumid -eq $mediumattachment.MediumId
       } # end foreach loop disk inventory
      } # end if $mediumattachment.Type -match 'Floppy'
     } # foreach $mediumattachment in $tempobj.IMediumAttachments
     $tempobj.IBiosSettings.Id = $global:vbox.IMachine_getBIOSSettings($vmid)
     $tempobj.IBiosSettings.LogoFadeIn = $global:vbox.IBIOSSettings_getLogoFadeIn($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.LogoFadeOut = $global:vbox.IBIOSSettings_getLogoFadeOut($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.LogoDisplayTime = $global:vbox.IBIOSSettings_getLogoDisplayTime($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.LogoImagePath = $global:vbox.IBIOSSettings_getLogoImagePath($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.BootMenuMode = $global:vbox.IBIOSSettings_getBootMenuMode($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.AcpiEnabled = $global:vbox.IBIOSSettings_getACPIEnabled($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.IoApicEnabled = $global:vbox.IBIOSSettings_getIOAPICEnabled($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.ApicMode = $global:vbox.IBIOSSettings_getAPICMode($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.TimeOffset = $global:vbox.IBIOSSettings_getTimeOffset($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.PxeDebugEnabled = $global:vbox.IBIOSSettings_getPXEDebugEnabled($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.NonVolatileStorageFile = $global:vbox.IBIOSSettings_getNonVolatileStorageFile($tempobj.IBiosSettings.Id)
     $tempobj.IBiosSettings.SmBiosUuidLittleEndian = $global:vbox.IBIOSSettings_getSMBIOSUuidLittleEndian($tempobj.IBiosSettings.Id)
     Write-Verbose "Found $($tempobj.Name) and adding to inventory"
     $vminventory += $tempobj
   } # end foreach loop inventory
  } # end if websrv
  elseif ($ModuleHost.ToLower() -eq 'com') {
   foreach ($vm in ($global:vbox.Machines)) {
     $guestprops = New-Object GuestProperties
     $tempobj = New-Object VirtualBoxVM
     $tempobj.Name = $vm.Name
     $tempobj.Description = $vm.Description
     $tempobj.State = [MachineState]::new().ToStr($vm.State)
     $tempobj.GuestOS = $vm.OSTypeID
     $tempobj.MemoryMb = $vm.MemorySize
     $tempobj.Guid = $vm.ID
     $tempobj.Accessible = $vm.Accessible
     $tempobj.AccessError = $vm.AccessError
     $tempobj.Groups = $vm.Groups
     $tempobj.OsTypeId = $vm.OSTypeId
     $tempobj.HardwareVersion = $vm.HardwareVersion
     $tempobj.HardwareUuid = $vm.HardwareUUID
     $tempobj.CpuCount = $vm.CPUCount
     $tempobj.CpuHotPlugEnabled = $vm.CPUHotPlugEnabled
     $tempobj.CpuExecutionCap = $vm.CPUExecutionCap
     $tempobj.CpuIdPortabilityLevel = $vm.CPUIDPortabilityLevel
     $tempobj.MemorySize = $vm.MemorySize
     $tempobj.MemoryBalloonSize = $vm.MemoryBalloonSize
     $tempobj.PageFusionEnabled = $vm.PageFusionEnabled
     $tempobj.PointingHidType = [PointingHIDType]::new().ToStr($vm.PointingHIDType - 1)
     $tempobj.KeyboardHidType = [KeyboardHIDType]::new().ToStr($vm.KeyboardHIDType - 1)
     $tempobj.HpetEnabled = $vm.HPETEnabled
     $tempobj.ChipsetType = [ChipsetType]::new().ToStr($vm.ChipsetType)
     $tempobj.SnapshotFolder = $vm.SnapshotFolder
     $tempobj.EmulatedUsbCardReaderEnabled = $vm.EmulatedUSBCardReaderEnabled
     $tempobj.SettingsFilePath = $vm.SettingsFilePath
     $tempobj.SessionState = [SessionState]::new().ToStr($vm.SessionState)
     $tempobj.LastStateChange = $vm.LastStateChange
     $tempobj.StateFilePath = $vm.StateFilePath
     $tempobj.LogFolder = $vm.LogFolder
     $tempobj.SnapshotCount = $vm.SnapshotCount
     $tempobj.CurrentStateModified = $vm.CurrentStateModified
     $tempobj.ClipboardMode = [ClipboardMode]::new().ToStr($vm.ClipboardMode)
     $tempobj.ClipboardFileTransfersEnabled = $vm.ClipboardFileTransfersEnabled
     $tempobj.DnDMode = [DnDMode]::new().ToStr($vm.DnDMode)
     $tempobj.TeleporterEnabled = $vm.TeleporterEnabled
     $tempobj.TeleporterPort = $vm.TeleporterPort
     $tempobj.TeleporterAddress = $vm.TeleporterAddress
     if ($vm.TeleporterPassword -ne '' -and $vm.TeleporterPassword -ne $null) {$tempobj.TeleporterPassword = ConvertTo-SecureString -String ($vm.TeleporterPassword) -AsPlainText -Force}
     $tempobj.ParavirtProvider = [ParavirtProvider]::new().ToStr($vm.ParavirtProvider)
     $tempobj.RtcUseUtc = $vm.RTCUseUTC
     $tempobj.IoCacheEnabled = $vm.IOCacheEnabled
     $tempobj.IoCacheSize = $vm.IOCacheSize
     $tempobj.TracingEnabled = $vm.TracingEnabled
     $tempobj.TracingConfig = $vm.TracingConfig
     $tempobj.AllowTracingToAccessVm = $vm.AllowTracingToAccessVM
     $tempobj.AutostartEnabled = $vm.AutostartEnabled
     $tempobj.AutostartDelay = $vm.AutostartDelay
     $tempobj.AutostopType = [AutostopType]::new().ToStr($vm.AutostopType - 1)
     $tempobj.DefaultFrontend = $vm.DefaultFrontend
     $tempobj.UsbProxyAvailable = $vm.USBProxyAvailable
     $tempobj.VmProcessPriority = [VMProcPriority]::new().ToStr($vm.VMProcessPriority)
     $tempobj.ParavirtDebug = $vm.ParavirtDebug
     $tempobj.CpuProfile = $vm.CpuProfile
     $tempobj.ISession.Session = New-Object -ComObject VirtualBox.Session
     $tempobj.IVrdeServer = [IVrdeServer]@{Enabled=$vm.VRDEServer.Enabled;AuthType=$vm.VRDEServer.AuthType;AuthTimeout=$vm.VRDEServer.AuthTimeout;AllowMultiConnection=$vm.VRDEServer.AllowMultiConnection;ReuseSingleConnection=$vm.VRDEServer.ReuseSingleConnection;VrdeExtPack=$vm.VRDEServer.VRDEExtPack;AuthLibrary=$vm.VRDEServer.AuthLibrary;VrdeProperties=$vm.VRDEServer.VRDEProperties}
     $tempobj.IVrdeServer.TcpPort = $vm.VRDEServer.GetVRDEProperty('TCP/Ports')
     $tempobj.IVrdeServer.IpAddress = $vm.VRDEServer.GetVRDEProperty('TCP/Address')
     $tempobj.IVrdeServer.VideoChannelEnabled = $vm.VRDEServer.GetVRDEProperty('VideoChannel/Enabled')
     $tempobj.IVrdeServer.VideoChannelQuality = $vm.VRDEServer.GetVRDEProperty('VideoChannel/Quality')
     $tempobj.IVrdeServer.VideoChannelDownscaleProtection = $vm.VRDEServer.GetVRDEProperty('VideoChannel/DownscaleProtection')
     $tempobj.IVrdeServer.DisableClientDisplay = $vm.VRDEServer.GetVRDEProperty('Client/DisableDisplay')
     $tempobj.IVrdeServer.DisableClientInput = $vm.VRDEServer.GetVRDEProperty('Client/DisableInput')
     $tempobj.IVrdeServer.DisableClientAudio = $vm.VRDEServer.GetVRDEProperty('Client/DisableAudio')
     $tempobj.IVrdeServer.DisableClientUsb = $vm.VRDEServer.GetVRDEProperty('Client/DisableUSB')
     $tempobj.IVrdeServer.DisableClientClipboard = $vm.VRDEServer.GetVRDEProperty('Client/DisableClipboard')
     $tempobj.IVrdeServer.DisableClientUpstreamAudio = $vm.VRDEServer.GetVRDEProperty('Client/DisableUpstreamAudio')
     $tempobj.IVrdeServer.DisableClientRdpdr = $vm.VRDEServer.GetVRDEProperty('Client/DisableRDPDR')
     $tempobj.IVrdeServer.H3dRedirectEnabled = $vm.VRDEServer.GetVRDEProperty('H3DRedirect/Enabled')
     $tempobj.IVrdeServer.SecurityMethod = $vm.VRDEServer.GetVRDEProperty('Security/Method')
     $tempobj.IVrdeServer.SecurityServerCertificate = $vm.VRDEServer.GetVRDEProperty('Security/ServerCertificate')
     $tempobj.IVrdeServer.SecurityServerPrivateKey = $vm.VRDEServer.GetVRDEProperty('Security/ServerPrivateKey')
     $tempobj.IVrdeServer.SecurityCaCertificate = $vm.VRDEServer.GetVRDEProperty('Security/CACertificate')
     $tempobj.IVrdeServer.AudioRateCorrectionMode = $vm.VRDEServer.GetVRDEProperty('Audio/RateCorrectionMode')
     $tempobj.IVrdeServer.AudioLogPath = $vm.VRDEServer.GetVRDEProperty('Audio/LogPath')
     foreach ($sc in $vm.StorageControllers) {[IStorageControllers[]]$tempobj.IStorageControllers += [IStorageControllers]@{Name=$sc.Name;MaxDevicesPerPortCount=$sc.MaxDevicesPerPortCount;MinPortCount=$sc.MinPortCount;MaxPortCount=$sc.MaxPortCount;Instance=$sc.Instance;PortCount=$sc.PortCount;Bus=[StorageBus]::new().ToStr($sc.Bus);ControllerType=[StorageControllerType]::new().ToStr($sc.ControllerType);UseHostIOCache=$sc.UseHostIOCache;Bootable=$sc.Bootable;ComObject=$sc}}
     foreach ($ma in $vm.MediumAttachments) {[IMediumAttachments[]]$tempobj.IMediumAttachments += [IMediumAttachments]@{Controller=$ma.Controller;Port=$ma.Port;Device=$ma.Device;Type=[DeviceType]::new().ToStr($ma.Type);Passthrough=$ma.Passthrough;TemporaryEject=$ma.TemporaryEject;IsEjected=$ma.IsEjected;NonRotational=$ma.NonRotational;Discard=$ma.Discard;HotPluggable=$ma.HotPluggable;BandwidthGroup=$ma.BandwidthGroup;ComObject=$ma}}
     $tempobj.IStorageControllers = $tempobj.IStorageControllers | Where-Object {$_.Name -ne $null}
     $tempobj.IMediumAttachments = $tempobj.IMediumAttachments | Where-Object {$_.Controller -ne $null}
     foreach ($mediumattachment in $tempobj.IMediumAttachments) {
      Write-Verbose "Getting attached disk: $($mediumattachment.ComObject.Medium.Name)"
      if ($mediumattachment.ComObject.Medium.Id) {
       $mediumattachment.IMedium.Name = $mediumattachment.ComObject.Medium.Name
       $mediumattachment.IMedium.Guid = $mediumattachment.ComObject.Medium.Id
       $mediumattachment.IMedium.Description = $mediumattachment.ComObject.Medium.Description
       $mediumattachment.IMedium.Format = $mediumattachment.ComObject.Medium.Format
       $mediumattachment.IMedium.Size = $mediumattachment.ComObject.Medium.Size
       $mediumattachment.IMedium.LogicalSize = $mediumattachment.ComObject.Medium.LogicalSize
       $mediumattachment.IMedium.MachineGuid = $mediumattachment.ComObject.Medium.MachineIds
       foreach ($machineid in $mediumattachment.IMedium.MachineGuid) {
        foreach ($imachine in $global:vbox.Machines) {
         if ($imachine.Id -eq $machineid) {
          $mediumattachment.IMedium.MachineName += $imachine.Name
         } # end if $imachine.Id -eq $machineid
         $mediumattachment.IMedium.MachineName = $mediumattachment.IMedium.MachineName | Where-Object {$_ -ne $null}
        } # foreach $imachine in $global:vbox.Machines
       } # foreach $machineid in $mediumattachment.IMedium.MachineGuid
       $mediumattachment.IMedium.State = [MediumState]::new().ToStr($mediumattachment.ComObject.Medium.State)
       foreach ($variant in $mediumattachment.ComObject.Medium.Variant) {
        $mediumattachment.IMedium.Variant = [MediumVariant]::new().ToStrCom($variant)
       } # foreach $variant in $mediumattachment.ComObject.Medium.Variant
       $mediumattachment.IMedium.Location = $mediumattachment.ComObject.Medium.Location
       $mediumattachment.IMedium.HostDrive = $mediumattachment.ComObject.Medium.HostDrive
       $mediumattachment.IMedium.MediumFormat = $mediumattachment.ComObject.Medium.MediumFormat
       $mediumattachment.IMedium.Type = [MediumType]::new().ToStr($mediumattachment.ComObject.Medium.Type)
       $mediumattachment.IMedium.Parent = $mediumattachment.ComObject.Medium.Parent
       $mediumattachment.IMedium.Children = $mediumattachment.ComObject.Medium.Children
       $mediumattachment.IMedium.ReadOnly = $mediumattachment.ComObject.Medium.ReadOnly
       $mediumattachment.IMedium.AutoReset = $mediumattachment.ComObject.Medium.AutoReset
       $mediumattachment.IMedium.LastAccessError = $mediumattachment.ComObject.Medium.LastAccessError
       $mediumattachment.IMedium.ComObject = $mediumattachment.ComObject.Medium
      } # end if $mediumattachment.ComObject.Medium.Id
     } # foreach $mediumattachment in $tempobj.IMediumAttachments
     $tempobj.IBiosSettings.LogoFadeIn = $vm.BIOSSettings.LogoFadeIn
     $tempobj.IBiosSettings.LogoFadeOut = $vm.BIOSSettings.LogoFadeOut
     $tempobj.IBiosSettings.LogoDisplayTime = $vm.BIOSSettings.LogoFadeDisplayTime
     $tempobj.IBiosSettings.LogoImagePath = $vm.BIOSSettings.LogoImagePath
     $tempobj.IBiosSettings.BootMenuMode = [BIOSBootMenuMode]::new().ToStr($vm.BIOSSettings.BootMenuMode)
     $tempobj.IBiosSettings.AcpiEnabled = $vm.BIOSSettings.ACPIEnabled
     $tempobj.IBiosSettings.IoApicEnabled = $vm.BIOSSettings.IOAPICEnabled
     $tempobj.IBiosSettings.ApicMode = $vm.BIOSSettings.APICMode
     $tempobj.IBiosSettings.TimeOffset = $vm.BIOSSettings.TimeOffset
     $tempobj.IBiosSettings.PxeDebugEnabled = $vm.BIOSSettings.PXEDebugEnabled
     $tempobj.IBiosSettings.NonVolatileStorageFile = $vm.BIOSSettings.NonVolatileStorageFile
     $tempobj.IBiosSettings.SmBiosUuidLittleEndian = $vm.BIOSSettings.SMBIOSUuidLittleEndian
     $tempobj.IBiosSettings.ComObject = $vm.BIOSSettings
     $tempobj.GuestProperties = $guestprops.Enumerate($vm, $ModuleHost)
     $tempobj.ComObject = $vm
     Write-Verbose "Found $($tempobj.Name) and adding to inventory"
     $vminventory += $tempobj
   } # end foreach loop inventory
  } # end elseif com
  # filter virtual machines
  if ($Name -and $Name -ne "*") {
   Write-Verbose "Filtering virtual machines by name: $Name"
   foreach ($vm in $vminventory) {
    Write-Verbose "Matching $($vm.Name) to $($Name)"
    if ($vm.Name -match $Name) {
     if ($State -and $vm.State -eq $State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Accessible=$vm.Accessible;AccessError=$vm.AccessError;Groups=$vm.Groups;OsTypeId=$vm.OsTypeId;HardwareVersion=$vm.HardwareVersion;HardwareUuid=$vm.HardwareUuid;CpuCount=$vm.CpuCount;CpuHotPlugEnabled=$vm.CpuHotPlugEnabled;CpuExecutionCap=$vm.CpuExecutionCap;CpuIdPortabilityLevel=$vm.CpuIdPortabilityLevel;MemorySize=$vm.MemorySize;MemoryBalloonSize=$vm.MemoryBalloonSize;PageFusionEnabled=$vm.PageFusionEnabled;PointingHidType=$vm.PointingHidType;KeyboardHidType=$vm.KeyboardHidType;HpetEnabled=$vm.HpetEnabled;ChipsetType=$vm.ChipsetType;SnapshotFolder=$vm.SnapshotFolder;EmulatedUsbCardReaderEnabled=$vm.EmulatedUsbCardReaderEnabled;SettingsFilePath=$vm.SettingsFilePath;SessionState=$vm.SessionState;LastStateChange=$vm.LastStateChange;StateFilePath=$vm.StateFilePath;LogFolder=$vm.LogFolder;SnapshotCount=$vm.SnapshotCount;CurrentStateModified=$vm.CurrentStateModified;ClipboardMode=$vm.ClipboardMode;ClipboardFileTransfersEnabled=$vm.ClipboardFileTransfersEnabled;DnDMode=$vm.DnDMode;TeleporterEnabled=$vm.TeleporterEnabled;TeleporterPort=$vm.TeleporterPort;TeleporterAddress=$vm.TeleporterAddress;TeleporterPassword=$vm.TeleporterPassword;ParavirtProvider=$vm.ParavirtProvider;RtcUseUtc=$vm.RTCUseUTC;IoCacheEnabled=$vm.IOCacheEnabled;IoCacheSize=$vm.IOCacheSize;TracingEnabled=$vm.TracingEnabled;TracingConfig=$vm.TracingConfig;AllowTracingToAccessVm=$vm.AllowTracingToAccessVM;AutostartEnabled=$vm.AutostartEnabled;AutostartDelay=$vm.AutostartDelay;AutostopType=$vm.AutostopType;DefaultFrontend=$vm.DefaultFrontend;UsbProxyAvailable=$vm.USBProxyAvailable;VmProcessPriority=$vm.VMProcessPriority;ParavirtDebug=$vm.ParavirtDebug;CpuProfile=$vm.CpuProfile;Id=$vm.Id;Guid=$vm.Guid;ISession=$vm.ISession;IVrdeServer=$vm.IVrdeServer;GuestProperties=$vm.GuestProperties;IStorageControllers=$vm.IStorageControllers;IMediumAttachments=$vm.IMediumAttachments;IBiosSettings=$vm.IBiosSettings;ComObject=$vm.ComObject}}
     elseif (!$State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Accessible=$vm.Accessible;AccessError=$vm.AccessError;Groups=$vm.Groups;OsTypeId=$vm.OsTypeId;HardwareVersion=$vm.HardwareVersion;HardwareUuid=$vm.HardwareUuid;CpuCount=$vm.CpuCount;CpuHotPlugEnabled=$vm.CpuHotPlugEnabled;CpuExecutionCap=$vm.CpuExecutionCap;CpuIdPortabilityLevel=$vm.CpuIdPortabilityLevel;MemorySize=$vm.MemorySize;MemoryBalloonSize=$vm.MemoryBalloonSize;PageFusionEnabled=$vm.PageFusionEnabled;PointingHidType=$vm.PointingHidType;KeyboardHidType=$vm.KeyboardHidType;HpetEnabled=$vm.HpetEnabled;ChipsetType=$vm.ChipsetType;SnapshotFolder=$vm.SnapshotFolder;EmulatedUsbCardReaderEnabled=$vm.EmulatedUsbCardReaderEnabled;SettingsFilePath=$vm.SettingsFilePath;SessionState=$vm.SessionState;LastStateChange=$vm.LastStateChange;StateFilePath=$vm.StateFilePath;LogFolder=$vm.LogFolder;SnapshotCount=$vm.SnapshotCount;CurrentStateModified=$vm.CurrentStateModified;ClipboardMode=$vm.ClipboardMode;ClipboardFileTransfersEnabled=$vm.ClipboardFileTransfersEnabled;DnDMode=$vm.DnDMode;TeleporterEnabled=$vm.TeleporterEnabled;TeleporterPort=$vm.TeleporterPort;TeleporterAddress=$vm.TeleporterAddress;TeleporterPassword=$vm.TeleporterPassword;ParavirtProvider=$vm.ParavirtProvider;RtcUseUtc=$vm.RTCUseUTC;IoCacheEnabled=$vm.IOCacheEnabled;IoCacheSize=$vm.IOCacheSize;TracingEnabled=$vm.TracingEnabled;TracingConfig=$vm.TracingConfig;AllowTracingToAccessVm=$vm.AllowTracingToAccessVM;AutostartEnabled=$vm.AutostartEnabled;AutostartDelay=$vm.AutostartDelay;AutostopType=$vm.AutostopType;DefaultFrontend=$vm.DefaultFrontend;UsbProxyAvailable=$vm.USBProxyAvailable;VmProcessPriority=$vm.VMProcessPriority;ParavirtDebug=$vm.ParavirtDebug;CpuProfile=$vm.CpuProfile;Id=$vm.Id;Guid=$vm.Guid;ISession=$vm.ISession;IVrdeServer=$vm.IVrdeServer;GuestProperties=$vm.GuestProperties;IStorageControllers=$vm.IStorageControllers;IMediumAttachments=$vm.IMediumAttachments;IBiosSettings=$vm.IBiosSettings;ComObject=$vm.ComObject}}
    } # end if $vm.Name -match $Name
   } # foreach $vm in $vminventory
  } # end if $Name and not *
  elseif ($Guid) {
   Write-Verbose "Filtering virtual machines by GUID: $Guid"
   foreach ($vm in $vminventory) {
    Write-Verbose "Matching $($vm.Guid) to $($Guid)"
    if ($vm.Guid -match $Guid) {
     if ($State -and $vm.State -eq $State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Accessible=$vm.Accessible;AccessError=$vm.AccessError;Groups=$vm.Groups;OsTypeId=$vm.OsTypeId;HardwareVersion=$vm.HardwareVersion;HardwareUuid=$vm.HardwareUuid;CpuCount=$vm.CpuCount;CpuHotPlugEnabled=$vm.CpuHotPlugEnabled;CpuExecutionCap=$vm.CpuExecutionCap;CpuIdPortabilityLevel=$vm.CpuIdPortabilityLevel;MemorySize=$vm.MemorySize;MemoryBalloonSize=$vm.MemoryBalloonSize;PageFusionEnabled=$vm.PageFusionEnabled;PointingHidType=$vm.PointingHidType;KeyboardHidType=$vm.KeyboardHidType;HpetEnabled=$vm.HpetEnabled;ChipsetType=$vm.ChipsetType;SnapshotFolder=$vm.SnapshotFolder;EmulatedUsbCardReaderEnabled=$vm.EmulatedUsbCardReaderEnabled;SettingsFilePath=$vm.SettingsFilePath;SessionState=$vm.SessionState;LastStateChange=$vm.LastStateChange;StateFilePath=$vm.StateFilePath;LogFolder=$vm.LogFolder;SnapshotCount=$vm.SnapshotCount;CurrentStateModified=$vm.CurrentStateModified;ClipboardMode=$vm.ClipboardMode;ClipboardFileTransfersEnabled=$vm.ClipboardFileTransfersEnabled;DnDMode=$vm.DnDMode;TeleporterEnabled=$vm.TeleporterEnabled;TeleporterPort=$vm.TeleporterPort;TeleporterAddress=$vm.TeleporterAddress;TeleporterPassword=$vm.TeleporterPassword;ParavirtProvider=$vm.ParavirtProvider;RtcUseUtc=$vm.RTCUseUTC;IoCacheEnabled=$vm.IOCacheEnabled;IoCacheSize=$vm.IOCacheSize;TracingEnabled=$vm.TracingEnabled;TracingConfig=$vm.TracingConfig;AllowTracingToAccessVm=$vm.AllowTracingToAccessVM;AutostartEnabled=$vm.AutostartEnabled;AutostartDelay=$vm.AutostartDelay;AutostopType=$vm.AutostopType;DefaultFrontend=$vm.DefaultFrontend;UsbProxyAvailable=$vm.USBProxyAvailable;VmProcessPriority=$vm.VMProcessPriority;ParavirtDebug=$vm.ParavirtDebug;CpuProfile=$vm.CpuProfile;Id=$vm.Id;Guid=$vm.Guid;ISession=$vm.ISession;IVrdeServer=$vm.IVrdeServer;GuestProperties=$vm.GuestProperties;IStorageControllers=$vm.IStorageControllers;IMediumAttachments=$vm.IMediumAttachments;IBiosSettings=$vm.IBiosSettings;ComObject=$vm.ComObject}}
     elseif (!$State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Accessible=$vm.Accessible;AccessError=$vm.AccessError;Groups=$vm.Groups;OsTypeId=$vm.OsTypeId;HardwareVersion=$vm.HardwareVersion;HardwareUuid=$vm.HardwareUuid;CpuCount=$vm.CpuCount;CpuHotPlugEnabled=$vm.CpuHotPlugEnabled;CpuExecutionCap=$vm.CpuExecutionCap;CpuIdPortabilityLevel=$vm.CpuIdPortabilityLevel;MemorySize=$vm.MemorySize;MemoryBalloonSize=$vm.MemoryBalloonSize;PageFusionEnabled=$vm.PageFusionEnabled;PointingHidType=$vm.PointingHidType;KeyboardHidType=$vm.KeyboardHidType;HpetEnabled=$vm.HpetEnabled;ChipsetType=$vm.ChipsetType;SnapshotFolder=$vm.SnapshotFolder;EmulatedUsbCardReaderEnabled=$vm.EmulatedUsbCardReaderEnabled;SettingsFilePath=$vm.SettingsFilePath;SessionState=$vm.SessionState;LastStateChange=$vm.LastStateChange;StateFilePath=$vm.StateFilePath;LogFolder=$vm.LogFolder;SnapshotCount=$vm.SnapshotCount;CurrentStateModified=$vm.CurrentStateModified;ClipboardMode=$vm.ClipboardMode;ClipboardFileTransfersEnabled=$vm.ClipboardFileTransfersEnabled;DnDMode=$vm.DnDMode;TeleporterEnabled=$vm.TeleporterEnabled;TeleporterPort=$vm.TeleporterPort;TeleporterAddress=$vm.TeleporterAddress;TeleporterPassword=$vm.TeleporterPassword;ParavirtProvider=$vm.ParavirtProvider;RtcUseUtc=$vm.RTCUseUTC;IoCacheEnabled=$vm.IOCacheEnabled;IoCacheSize=$vm.IOCacheSize;TracingEnabled=$vm.TracingEnabled;TracingConfig=$vm.TracingConfig;AllowTracingToAccessVm=$vm.AllowTracingToAccessVM;AutostartEnabled=$vm.AutostartEnabled;AutostartDelay=$vm.AutostartDelay;AutostopType=$vm.AutostopType;DefaultFrontend=$vm.DefaultFrontend;UsbProxyAvailable=$vm.USBProxyAvailable;VmProcessPriority=$vm.VMProcessPriority;ParavirtDebug=$vm.ParavirtDebug;CpuProfile=$vm.CpuProfile;Id=$vm.Id;Guid=$vm.Guid;ISession=$vm.ISession;IVrdeServer=$vm.IVrdeServer;GuestProperties=$vm.GuestProperties;IStorageControllers=$vm.IStorageControllers;IMediumAttachments=$vm.IMediumAttachments;IBiosSettings=$vm.IBiosSettings;ComObject=$vm.ComObject}}
    } # end if $vm.Guid -match $Guid
   } # foreach $vm in $vminventory
  } # end if $Guid
  elseif ($PSCmdlet.ParameterSetName -eq "All" -or $Name -eq "*") {
   if ($State) {
    Write-Verbose "Filtering all virtual machines by state: $State"
    foreach ($vm in $vminventory) {
     if ($vm.State -eq $State) {[VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Accessible=$vm.Accessible;AccessError=$vm.AccessError;Groups=$vm.Groups;OsTypeId=$vm.OsTypeId;HardwareVersion=$vm.HardwareVersion;HardwareUuid=$vm.HardwareUuid;CpuCount=$vm.CpuCount;CpuHotPlugEnabled=$vm.CpuHotPlugEnabled;CpuExecutionCap=$vm.CpuExecutionCap;CpuIdPortabilityLevel=$vm.CpuIdPortabilityLevel;MemorySize=$vm.MemorySize;MemoryBalloonSize=$vm.MemoryBalloonSize;PageFusionEnabled=$vm.PageFusionEnabled;PointingHidType=$vm.PointingHidType;KeyboardHidType=$vm.KeyboardHidType;HpetEnabled=$vm.HpetEnabled;ChipsetType=$vm.ChipsetType;SnapshotFolder=$vm.SnapshotFolder;EmulatedUsbCardReaderEnabled=$vm.EmulatedUsbCardReaderEnabled;SettingsFilePath=$vm.SettingsFilePath;SessionState=$vm.SessionState;LastStateChange=$vm.LastStateChange;StateFilePath=$vm.StateFilePath;LogFolder=$vm.LogFolder;SnapshotCount=$vm.SnapshotCount;CurrentStateModified=$vm.CurrentStateModified;ClipboardMode=$vm.ClipboardMode;ClipboardFileTransfersEnabled=$vm.ClipboardFileTransfersEnabled;DnDMode=$vm.DnDMode;TeleporterEnabled=$vm.TeleporterEnabled;TeleporterPort=$vm.TeleporterPort;TeleporterAddress=$vm.TeleporterAddress;TeleporterPassword=$vm.TeleporterPassword;ParavirtProvider=$vm.ParavirtProvider;RtcUseUtc=$vm.RTCUseUTC;IoCacheEnabled=$vm.IOCacheEnabled;IoCacheSize=$vm.IOCacheSize;TracingEnabled=$vm.TracingEnabled;TracingConfig=$vm.TracingConfig;AllowTracingToAccessVm=$vm.AllowTracingToAccessVM;AutostartEnabled=$vm.AutostartEnabled;AutostartDelay=$vm.AutostartDelay;AutostopType=$vm.AutostopType;DefaultFrontend=$vm.DefaultFrontend;UsbProxyAvailable=$vm.USBProxyAvailable;VmProcessPriority=$vm.VMProcessPriority;ParavirtDebug=$vm.ParavirtDebug;CpuProfile=$vm.CpuProfile;Id=$vm.Id;Guid=$vm.Guid;ISession=$vm.ISession;IVrdeServer=$vm.IVrdeServer;GuestProperties=$vm.GuestProperties;IStorageControllers=$vm.IStorageControllers;IMediumAttachments=$vm.IMediumAttachments;IBiosSettings=$vm.IBiosSettings;ComObject=$vm.ComObject}}
    } # foreach $vm in $vminventory
   } # end if $State
   else {
    Write-Verbose "Filtering all virtual machines"
    foreach ($vm in $vminventory) {
     [VirtualBoxVM[]]$obj += [VirtualBoxVM]@{Name=$vm.Name;Description=$vm.Description;State=$vm.State;GuestOS=$vm.GuestOS;MemoryMb=$vm.MemoryMb;Accessible=$vm.Accessible;AccessError=$vm.AccessError;Groups=$vm.Groups;OsTypeId=$vm.OsTypeId;HardwareVersion=$vm.HardwareVersion;HardwareUuid=$vm.HardwareUuid;CpuCount=$vm.CpuCount;CpuHotPlugEnabled=$vm.CpuHotPlugEnabled;CpuExecutionCap=$vm.CpuExecutionCap;CpuIdPortabilityLevel=$vm.CpuIdPortabilityLevel;MemorySize=$vm.MemorySize;MemoryBalloonSize=$vm.MemoryBalloonSize;PageFusionEnabled=$vm.PageFusionEnabled;PointingHidType=$vm.PointingHidType;KeyboardHidType=$vm.KeyboardHidType;HpetEnabled=$vm.HpetEnabled;ChipsetType=$vm.ChipsetType;SnapshotFolder=$vm.SnapshotFolder;EmulatedUsbCardReaderEnabled=$vm.EmulatedUsbCardReaderEnabled;SettingsFilePath=$vm.SettingsFilePath;SessionState=$vm.SessionState;LastStateChange=$vm.LastStateChange;StateFilePath=$vm.StateFilePath;LogFolder=$vm.LogFolder;SnapshotCount=$vm.SnapshotCount;CurrentStateModified=$vm.CurrentStateModified;ClipboardMode=$vm.ClipboardMode;ClipboardFileTransfersEnabled=$vm.ClipboardFileTransfersEnabled;DnDMode=$vm.DnDMode;TeleporterEnabled=$vm.TeleporterEnabled;TeleporterPort=$vm.TeleporterPort;TeleporterAddress=$vm.TeleporterAddress;TeleporterPassword=$vm.TeleporterPassword;ParavirtProvider=$vm.ParavirtProvider;RtcUseUtc=$vm.RTCUseUTC;IoCacheEnabled=$vm.IOCacheEnabled;IoCacheSize=$vm.IOCacheSize;TracingEnabled=$vm.TracingEnabled;TracingConfig=$vm.TracingConfig;AllowTracingToAccessVm=$vm.AllowTracingToAccessVM;AutostartEnabled=$vm.AutostartEnabled;AutostartDelay=$vm.AutostartDelay;AutostopType=$vm.AutostopType;DefaultFrontend=$vm.DefaultFrontend;UsbProxyAvailable=$vm.USBProxyAvailable;VmProcessPriority=$vm.VMProcessPriority;ParavirtDebug=$vm.ParavirtDebug;CpuProfile=$vm.CpuProfile;Id=$vm.Id;Guid=$vm.Guid;ISession=$vm.ISession;IVrdeServer=$vm.IVrdeServer;GuestProperties=$vm.GuestProperties;IStorageControllers=$vm.IStorageControllers;IMediumAttachments=$vm.IMediumAttachments;IBiosSettings=$vm.IBiosSettings;ComObject=$vm.ComObject}
    } # foreach $vm in $vminventory
   } # end else
  } # end if All
  Write-Verbose "Found $(($obj | Measure-Object).count) virtual machine(s)"
  if ($obj) {
   Write-Verbose "Found $($obj.Name)"
   # write virtual machines object to the pipeline as an array
   Write-Output ([System.Array]$obj)
  } # end if $obj
  else {Write-Verbose "[Warning] No matching virtual machines found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception retreiving machine information'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
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
A switch to skip web service update. (Only for use between functions - DO NOT USE)
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
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
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
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.State -eq 'Running') {
     if ($ModuleHost.ToLower() -eq 'websrv') {
      Write-Verbose "Suspending $($imachine.Name)"
      # get the machine session
      Write-Verbose "Getting the machine session"
      $imachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($imachine.Id)
      # lock the vm session
      Write-Verbose "Locking the machine session"
      $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Shared'))
      # get the machine IConsole session
      Write-Verbose "Getting the machine IConsole session"
      $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession.Id)
      # suspend the vm
      Write-Verbose "Pausing the virtual machine"
      $global:vbox.IConsole_pause($imachine.IConsole)
     } # end if websrv
     elseif ($ModuleHost.ToLower() -eq 'com') {
      Write-Verbose "Suspending $($imachine.Name)"
      # lock the vm session
      Write-Verbose "Locking the machine session"
      $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Shared'))
      # suspend the vm
      Write-Verbose "Pausing the virtual machine"
      $imachine.ISession.Session.Console.Pause()
     } # end elseif com
    } # end if $imachine.State -eq 'Running'
    else {Write-Verbose "The requested virtual machine `"$($imachine.Name)`" can't be paused because it is not running (State: $($imachine.State))"}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Verbose "[Warning] No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception suspending machine'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
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
A switch to skip web service update. (Only for use between functions - DO NOT USE)
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
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
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
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.State -eq 'Paused') {
     if ($ModuleHost.ToLower() -eq 'websrv') {
      Write-Verbose "Resuming $($imachine.Name)"
      # get the machine session
      Write-Verbose "Getting the machine session"
      $imachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($imachine.Id)
      # lock the vm session
      Write-Verbose "Locking the machine session"
      $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Shared'))
      # get the machine IConsole session
      Write-Verbose "Getting the machine IConsole session"
      $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession.Id)
      # resume the vm
      Write-Verbose "Resuming the virtual machine"
      $global:vbox.IConsole_resume($imachine.IConsole)
     } # end if websrv
     elseif ($ModuleHost.ToLower() -eq 'com') {
      Write-Verbose "Resuming $($imachine.Name)"
      # lock the vm session
      Write-Verbose "Locking the machine session"
      $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Shared'))
      # resume the vm
      Write-Verbose "Resuming the virtual machine"
      $imachine.ISession.Session.Console.Resume()
     } # end elseif com
    } # end if $imachine.State -eq 'Running'
    else {Write-Verbose "The requested virtual machine `"$($imachine.Name)`" can't be resumed because it is not paused (State: $($imachine.State))"}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Host "[Error] No matching virtual machines were found using specified parameters" -ForegroundColor Red -BackgroundColor Black;return}
 } # Try
 catch {
  Write-Verbose 'Exception resuming machine'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Start-VirtualBoxVM {
<#
.SYNOPSIS
Start a virtual machine
.DESCRIPTION
Start VirtualBox VMs by machine object, name, or GUID in the order they are provided. The default Type is to start them in GUI mode. You can also run them in headless mode which will start a new hidden process. If the machine(s) disk(s) are encrypted, you must specify the -Encrypted switch and supply credential(s) using the -Credentials parameter. The username (identifier) is the name of the virtual machine by default, unless it has been otherwise specified.
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
.PARAMETER Credentials
Powershell credentials. Must be provided if the -Encrypted switch is used. The 'UserName' field of the credential must be the disk ID of the desired disk. You can supply multiple credentials in comma-separated format. (i.e. -Credentials $cred1,$cred2) Credentials can be provided in any order
.PARAMETER ProgressBar
A switch to display a progress bar.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Start-VirtualBoxVM "Win10"
Starts the virtual machine called Win10 in GUI mode.
.EXAMPLE
PS C:\> Start-VirtualBoxVM "2016" -Headless -Encrypted -Credentials $diskCredentials
Start the virtual machine called "2016 Core" in headless mode and provides credentials to decrypt the disk(s) on boot.
.EXAMPLE
PS C:\> Start-VirtualBoxVM "2016","Win10" -Headless -Encrypted -Credentials $10Credentials,$2016Credentials
Start the virtual machines called "2016 Core" and "Win10" in headless mode and provides credentials to decrypt the disks on boot.
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
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
PsCredential[]:  Credentials for virtual machine disks
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
  [VirtualBoxVM[]]$Machine,
[Parameter(ParameterSetName='Unencrypted',Mandatory=$false,
ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
[Parameter(ParameterSetName='Encrypted',Mandatory=$false,
ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
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
[ValidateSet("Headless","Gui","Sdl")]
  [string]$Type = 'Gui',
[Parameter(ParameterSetName='Encrypted',Mandatory=$true,
HelpMessage="Use this switch if VM disk(s) are encrypted")]
  [switch]$Encrypted,
[Parameter(ParameterSetName='Encrypted',Mandatory=$true,
HelpMessage="Enter the credentials to unlock the VM disk(s)")]
  [pscredential[]]$Credentials,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
  if ($Encrypted) {$disks = Get-VirtualBoxDisk -Machine $imachines}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
  if ($Encrypted) {$disks = Get-VirtualBoxDisk -Machine $imachines}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
  if ($Encrypted) {$disks = Get-VirtualBoxDisk -Machine $imachines}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
   if ($imachine.State -match 'PoweredOff') {
    if ($ModuleHost.ToLower() -eq 'websrv') {
     if (-not $Encrypted) {
      # start the vm in $Type mode
      Write-Verbose "Starting VM $($imachine.Name) in $Type mode"
      if ($Type -match 'Gui' -or $Type -match 'Sdl') {Write-Host "[Error] Starting VM in GUI or SDL mode is not available for the Web Service. Try again using the -Type Headless parameter and value." -ForegroundColor Red -BackgroundColor Black;return}
      elseif ($Type -match 'Headless') {$imachine.IProgress.Id = $global:vbox.IMachine_launchVMProcess($imachine.Id, $imachine.ISession.Id, $Type.ToLower(), $null)}
      if ($ProgressBar) {
       # collect iprogress data
       Write-Verbose "Fetching IProgress data"
       $imachine.IProgress = $imachine.IProgress.Fetch($imachine.IProgress.Id)
       Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1
      } # end if $ProgressBar
      do {
       # get the current machine state
       $machinestate = $global:vbox.IMachine_getState($imachine.Id)
       # update iprogress data
       if ($ProgressBar) {
        $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
        if ($imachine.IProgress.Percent -lt 20) {
         Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1
         if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
        } # end if $imachine.IProgress.Percent < 20
        else {
         Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)
         if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
        } # end else
       } # end if $ProgressBar
      } until ($machinestate -eq 'Running') # continue once the vm is running
     } # end if not Encrypted
     elseif ($Encrypted) {
      # start the vm in $Type mode
      Write-Verbose "Starting VM $($imachine.Name) in $Type mode"
      if ($Type -match 'Gui' -or $Type -match 'Sdl') {Write-Host "[Error] Starting VM in GUI or SDL mode is not available for the Web Service. Try again using the -Type Headless parameter and value." -ForegroundColor Red -BackgroundColor Black;return}
      elseif ($Type -match 'Headless') {$imachine.IProgress.Id = $global:vbox.IMachine_launchVMProcess($imachine.Id, $imachine.ISession.Id, 'headless', $null)}
      if ($ProgressBar) {
       # collect iprogress data
       Write-Verbose "Fetching IProgress data"
       $imachine.IProgress = $imachine.IProgress.Fetch($imachine.IProgress.Id)
       Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1
      } # end if $ProgressBar
      Write-Verbose "Waiting for VM $($imachine.Name) to pause for password"
      do {
       # get the current machine state
       $machinestate = $global:vbox.IMachine_getState($imachine.Id)
       # update iprogress data
       if ($ProgressBar) {
        $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
        if ($imachine.IProgress.Percent -lt 20) {
         Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1
         if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
        } # end if $imachine.IProgress.Percent < 20
        else {
         Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)
         if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
        } # end else
       } # end if $ProgressBar
      } until ($machinestate -eq 'Paused') # continue once the vm pauses for password
      Write-Verbose "VM $($imachine.Name) paused"
      # create new session object for iconsole
      Write-Verbose "Getting IConsole Session object for VM $($imachine.Name)"
      $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession.Id)
      $exptntrnsltr = New-Object VirtualBoxError
      foreach ($Credential in $Credentials) {
       foreach ($disk in $disks) {
        Write-Verbose "Processing disk $($disk.Name)"
        $diskid = $null
        try {
         # get the disk id
         Write-Verbose "Getting the disk ID"
         $cipher = $global:vbox.IMedium_getEncryptionSettings($disk.Id, [ref]$diskid) # this looks really confusing but it's 2 seperate types of ID
         Write-Verbose "Disk $($disk.Name) encryption cipher: $cipher"
         if ($diskid -eq $Credential.UserName) {
          Write-Verbose "Disk ID `"$($diskid)`" matches the UserName `"$($Credential.UserName)`" of the current credential"
          # check the password against the vm disk
          Write-Verbose "Checking for Password against disk"
          $global:vbox.IMedium_checkEncryptionPassword($disk.Id, $Credential.GetNetworkCredential().Password)
          Write-Verbose  "The image is configured for encryption and the password is correct"
          # pass disk encryption password to the vm console
          Write-Verbose "Sending Identifier: $($imachine.Name) with password: $($Credential.Password)"
          $global:vbox.IConsole_addDiskEncryptionPassword($imachine.IConsole, $diskid, $Credential.GetNetworkCredential().Password, $false)
          Write-Verbose "Disk decryption successful for disk $($disk.Name)"
         } # end if $diskid -eq $Credential.UserName
         else {Write-Verbose "Disk ID `"$($diskid)`" does not match the UserName `"$($Credential.UserName)`" of the current credential"}
        } # Try
        catch {
         Write-Verbose "Exception when sending password for encrypted disk(s)"
         Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
         Write-Verbose $_.Exception.Message
         Write-Host "[Error] Could not decrypt disk $($disk.Name) using provided credential: `"$($exptntrnsltr.Description($_.Exception.Message))`"" -ForegroundColor Red -BackgroundColor Black
        } # Catch
       } # end foreach $disk in $disks
      } # foreach $Credential in $Credentials
      if ($global:vbox.IMachine_getState($imachine.Id) -match 'Paused') {
       Write-Verbose "Decryption unsuccessful for disk $($disk.Name) - powering off machine $($imachine.Name)"
       Write-Host "[Error] Could not decrypt disk $($disk.Name) for machine $($imachine.Name) using any of the provided credentials" -ForegroundColor Red -BackgroundColor Black
       $imachine.IProgress.Id = $global:vbox.IConsole_powerDown($imachine.IConsole)
      }
     } # end elseif Encrypted
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     if (-not $Encrypted) {
      # start the vm in $Type mode
      Write-Verbose "Starting VM $($imachine.Name) in $Type mode"
      if ($Type -match 'Sdl') {Show-VirtualBoxVM -Machine $imachine}
      elseif ($Type -notmatch 'Sdl') {$imachine.IProgress.Progress = $imachine.ComObject.LaunchVMProcess($imachine.ISession.Session, $Type.ToLower(), [string[]]@())}
      if ($ProgressBar) {Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1}
      do {
       # get the current machine state
       $machinestate = $imachine.ComObject.State
       # update iprogress data
       if ($ProgressBar) {
        if ($imachine.IProgress.Progress.Percent -lt 20) {
         Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1
         if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
        } # end if $imachine.IProgress.Progress.Percent < 20
        else {
         Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.Progress.TimeRemaining)
         if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
        } # end else
       } # end if $ProgressBar
      } until ($machinestate -eq 5) # continue once the vm is running
     } # end if not Encrypted
     elseif ($Encrypted) {
      # start the vm in $Type mode
      Write-Verbose "Starting VM $($imachine.Name) in $Type mode"
      if ($Type -match 'Sdl') {Show-VirtualBoxVM -Machine $imachine}
      elseif ($Type -notmatch 'Sdl') {$imachine.IProgress.Progress = $imachine.ComObject.LaunchVMProcess($imachine.ISession.Session, $Type.ToLower(), [string[]]@())}
      if ($ProgressBar) {Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1}
      Write-Verbose "Waiting for VM $($imachine.Name) to pause for password"
      do {
       # get the current machine state
       $machinestate = $imachine.ComObject.State
       # update iprogress data
       if ($ProgressBar) {
        if ($imachine.IProgress.Progress.Percent -lt 20) {
         Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1
         if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
        } # end if $imachine.IProgress.Progress.Percent < 20
        else {
         Write-Progress -Activity "Starting VM $($imachine.Name) in $Type Mode" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.Progress.TimeRemaining)
         if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
        } # end else
       } # end if $ProgressBar
      } until ($machinestate -eq 5) # continue once the vm pauses for password
      Write-Verbose "VM $($imachine.Name) paused"
      if ($Type -match 'Sdl') {
       Write-Verbose "Getting shared lock on machine $($imachine.Name)"
       $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Shared'))
      }
      $exptntrnsltr = New-Object VirtualBoxError
      foreach ($Credential in $Credentials) {
       foreach ($disk in $disks) {
        Write-Verbose "Processing disk $($disk.Name)"
        $cipher = $null
        try {
         # get the disk id
         Write-Verbose "Getting the disk ID"
         $diskid = $disk.ComObject.getEncryptionSettings([string]@([ref]$cipher))
         Write-Verbose "Disk $($disk.Name) encryption cipher: $cipher"
         if ($diskid -eq $Credential.UserName) {
          Write-Verbose "Disk ID `"$($diskid)`" matches the UserName `"$($Credential.UserName)`" of the current credential"
          # check the password against the vm disk
          Write-Verbose "Checking for Password against disk"
          $disk.ComObject.checkEncryptionPassword($Credential.GetNetworkCredential().Password)
          Write-Verbose  "The image is configured for encryption and the password is correct"
          # pass disk encryption password to the vm console
          Write-Verbose "Sending Identifier: $($imachine.Name) with password: $($Credential.Password)"
          $imachine.ISession.Session.Console.AddDiskEncryptionPassword($diskid, $Credential.GetNetworkCredential().Password, 0)
          Write-Verbose "Disk decryption successful for disk $($disk.Name)"
         } # end if $diskid -eq $Credential.UserName
         else {Write-Verbose "Disk ID `"$($diskid)`" does not match the UserName `"$($Credential.UserName)`" of the current credential"}
        } # Try
        catch {
         Write-Verbose "Exception when sending password for encrypted disk(s)"
         Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
         Write-Verbose $_.Exception.Message
         Write-Host "[Error] Could not decrypt disk $($disk.Name) using provided credential: `"$($exptntrnsltr.Description($_.Exception.Message))`"" -ForegroundColor Red -BackgroundColor Black
        } # Catch
       } # end foreach $disk in $disks
      } # foreach $Credential in $Credentials
      if ($imachine.ComObject.State -eq 6) {
       Write-Verbose "Decryption unsuccessful for disk $($disk.Name) - powering off machine $($imachine.Name)"
       Write-Host "[Error] Could not decrypt disk $($disk.Name) for machine $($imachine.Name) using any of the provided credentials" -ForegroundColor Red -BackgroundColor Black
       $imachine.IProgress.Progress = $imachine.ISession.Session.Console.PowerDown()
      }
     } # end elseif Encrypted
    } # end elseif com
   } # end if $machine.State -match 'PoweredOff'
   else {Write-Host "[Error] Only VMs that have been powered off can be started. The state of $($imachine.Name) is $($imachine.State)" -ForegroundColor Red -BackgroundColor Black;return}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Host "[Error] No matching virtual machines were found using specified parameters" -ForegroundColor Red -BackgroundColor Black;return}
 } # Try
 catch {
  Write-Verbose 'Exception starting machine'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
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
A switch to skip web service update. (Only for use between functions - DO NOT USE)
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
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
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
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
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
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    # create Vbox session object
    #Write-Verbose "Creating a session object"
    #$imachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
    if ($Acpi) {
     Write-Verbose "ACPI Shutdown requested"
     if ($imachine.State -eq 'Running') {
      if ($ModuleHost.ToLower() -eq 'websrv') {
       Write-verbose "Locking the machine session"
       $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Shared'))
       # create iconsole session to vm
       Write-verbose "Creating IConsole session to the machine"
       $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession.Id)
       #send ACPI shutdown signal
       Write-verbose "Sending ACPI Shutdown signal to the machine"
       $global:vbox.IConsole_powerButton($imachine.IConsole)
      } # end if websrv
      elseif ($ModuleHost.ToLower() -eq 'com') {
       Write-verbose "Locking the machine session"
       $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Shared'))
       #send ACPI shutdown signal
       Write-verbose "Sending ACPI Shutdown signal to the machine"
       $imachine.ISession.Session.Console.PowerButton()
      } # end elseif com
     }
     else {return "Only machines that are running may be stopped."}
    }
    elseif ($PsShutdown) {
     Write-Verbose "PowerShell Shutdown requested"
     if ($imachine.State -eq 'Running') {
      # send a stop-computer -force command to the guest machine
      Write-Verbose 'Sending PowerShell Stop-Computer -Force -Confirm:$false command to guest machine'
      Write-Output (Submit-VirtualBoxVMProcess -Machine $imachine -PathToExecutable 'cmd.exe' -Arguments '/c','powershell.exe','-ExecutionPolicy','Bypass','-Command','Stop-Computer','-Force','-Confirm:$false' -Credential $Credential -NoWait -SkipCheck)
     }
     else {return "Only machines that are running may be stopped."}
    }
    else {
     Write-Verbose "Power-off requested"
     if ($ModuleHost.ToLower() -eq 'websrv') {
      if ($global:vbox.IMachine_getState($imachine.Id) -ne 'PoweredOff') {
       Write-verbose "Locking the machine session"
       $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Shared'))
       # create iconsole session to vm
       Write-verbose "Creating IConsole session to the machine"
       $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession.Id)
       # Power off the machine
       Write-verbose "Powering off the machine"
       $imachine.IProgress.Id = $global:vbox.IConsole_powerDown($imachine.IConsole)
       if ($ProgressBar) {
        # collect iprogress data
        Write-Verbose "Fetching IProgress data"
        $imachine.IProgress = $imachine.IProgress.Fetch($imachine.IProgress.Id)
        Write-Progress -Activity "Powering VM $($imachine.Name) off" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1
       } # end if $ProgressBar
       do {
        # get the current machine state
        $machinestate = $global:vbox.IMachine_getState($imachine.Id)
        # update iprogress data
        if ($ProgressBar) {
        $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
        if ($imachine.IProgress.Percent -lt 20) {
         Write-Progress -Activity "Powering VM $($imachine.Name) off" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1
         if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
         } # end if $imachine.IProgress.Percent < 20
         else {
          Write-Progress -Activity "Powering VM $($imachine.Name) off" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)
          if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
         } # end else
        } # end if $ProgressBar
       } until ($machinestate -eq 'PoweredOff') # continue once the vm is stopped
      } # end if machine is powered off
      else {return "Only machines that are not powered off may be stopped."}
     } # end if websrv
     elseif ($ModuleHost.ToLower() -eq 'com') {
      if ($imachine.ComObject.State -ne 1) {
       Write-verbose "Locking the machine session"
       $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Shared'))
       # Power off the machine
       Write-verbose "Powering off the machine"
       $imachine.IProgress.Progress = $imachine.ISession.Session.Console.PowerDown()
       # collect iprogress data
       if ($ProgressBar) {Write-Progress -Activity "Powering the VM $($imachine.Name) off" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1}
       do {
        # get the current machine state
        $machinestate = $imachine.ComObject.State
        # update iprogress data
        if ($ProgressBar) {
         if ($imachine.IProgress.Progress.Percent -lt 20) {
          Write-Progress -Activity "Powering the VM $($imachine.Name) off" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1
          if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
         } # end if $imachine.IProgress.Progress.Percent < 20
         else {
          Write-Progress -Activity "Powering the VM $($imachine.Name) off" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.Progress.TimeRemaining)
          if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
         } # end else
        } # end if $ProgressBar
       } until ($machinestate -eq 1) # continue once the vm is stopped
      } # end if $imachine.ComObject.State -ne 1
      else {return "Only machines that are not powered off may be stopped."}
     } # end elseif com
    } # end else
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Host "[Error] No matching virtual machines were found using specified parameters" -ForegroundColor Red -BackgroundColor Black;return}
 } # Try
 catch {
  Write-Verbose 'Exception starting machine'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
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
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
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
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function New-VirtualBoxVM {
<#
.SYNOPSIS
Create a virtual machine
.DESCRIPTION
Creates a new virtual machine. The name provided by the Name parameter must not exist in the VirtualBox inventory, or this command will fail. You can optionally supply custom values using a large number of parameters available to this command. There are too many to fully document in this help text, so tab completion has been added where it is possible. The values provided by tab completion are updated when Start-VirtualBoxSession is successfully run. To force the values to be updated again, use the -Force switch with Start-VirtualBoxSession.
.PARAMETER Name
The name of the virtual machine. This is a required parameter.
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
.PARAMETER Icon
The custom icon for the virtual machine. Must be a valid image file or the command will fail.
.PARAMETER IoCacheEnabled
The Enable or disable IO cache for the virtual machine.
.PARAMETER IoCacheSize
The IO cache size in MB for the virtual machine.
.PARAMETER KeyboardHidType
The keyboard HID type for the virtual machine.
.PARAMETER Location
The location for the virtual machine files.
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
A switch to skip web service update. (Only for use between functions - DO NOT USE)
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
Remove-VirtualBoxVM
Import-VirtualBoxVM
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
  [string]$Location,
[Parameter(HelpMessage="Enter optional virtual machine icon",
ParameterSetName='Custom',Mandatory=$false)]
[ValidateScript({Test-Path $_})]
[ValidateNotNullOrEmpty()]
  [string]$Icon,
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
 $OsTypeIdAttributes = New-Object System.Management.Automation.ParameterAttribute
 $OsTypeIdAttributes.Mandatory = $true
 $OsTypeIdAttributes.Position = 1
 $OsTypeIdAttributes.HelpMessage = 'Enter the type ID for the virtual machine guest OS'
 $OsTypeIdCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $OsTypeIdCollection.Add($OsTypeIdAttributes)
 $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute(@('Other','Other_64','Windows31','Windows95','Windows98','WindowsMe','WindowsNT3x','WindowsNT4','Windows2000','WindowsXP','WindowsXP_64','Windows2003','Windows2003_64','WindowsVista','WindowsVista_64','Windows2008','Windows2008_64','Windows7','Windows7_64','Windows8','Windows8_64','Windows81','Windows81_64','Windows2012_64','Windows10','Windows10_64','Windows2016_64','Windows2019_64','WindowsNT','WindowsNT_64','Linux22','Linux24','Linux24_64','Linux26','Linux26_64','ArchLinux','ArchLinux_64','Debian','Debian_64','Fedora','Fedora_64','Gentoo','Gentoo_64','Mandriva','Mandriva_64','Oracle','Oracle_64','RedHat','RedHat_64','OpenSUSE','OpenSUSE_64','Turbolinux','Turbolinux_64','Ubuntu','Ubuntu_64','Xandros','Xandros_64','Linux','Linux_64','Solaris','Solaris_64','OpenSolaris','OpenSolaris_64','Solaris11_64','FreeBSD','FreeBSD_64','OpenBSD','OpenBSD_64','NetBSD','NetBSD_64','OS2Warp3','OS2Warp4','OS2Warp45','OS2eCS','OS21x','OS2','MacOS','MacOS_64','MacOS106','MacOS106_64','MacOS107_64','MacOS108_64','MacOS109_64','MacOS1010_64','MacOS1011_64','MacOS1012_64','MacOS1013_64','DOS','Netware','L4','QNX','JRockitVE','VBoxBS_64'))
 if ($global:guestostype.id) {
  $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute($global:guestostype.id)
 }
 $OsTypeIdCollection.Add($ValidateSetOsTypeId)
 $OsTypeId = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("OsTypeId", [string], $OsTypeIdCollection)
 $CustomAttributes = New-Object System.Management.Automation.ParameterAttribute
 $CustomAttributes.Mandatory = $false
 $CustomAttributes.ParameterSetName = 'Custom'
 $CustomAttributes.HelpMessage = 'Enter the paravirtual provider for the virtual machine'
 $ParavirtProvidersCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $ParavirtProvidersCollection.Add($CustomAttributes)
 $ValidateSetParavirtProviders = New-Object System.Management.Automation.ValidateSetAttribute(@('None','Default','Legacy','Minimal','HyperV','KVM'))
 if ($global:systempropertiessupported.ParavirtProviders) {
  $ValidateSetParavirtProviders = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.ParavirtProviders)
 }
 $ParavirtProvidersCollection.Add($ValidateSetParavirtProviders)
 $ParavirtProviders = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ParavirtProvider", [string], $ParavirtProvidersCollection)
 $CustomAttributes.HelpMessage = 'Enter the clipboard mode for the virtual machine'
 $ClipboardModesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $ClipboardModesCollection.Add($CustomAttributes)
 $ValidateSetClipboardModes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disabled','HostToGuest','GuestToHost','Bidirectional'))
 if ($global:systempropertiessupported.ClipboardModes) {
  $ValidateSetClipboardModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.ClipboardModes)
 }
 $ClipboardModesCollection.Add($ValidateSetClipboardModes)
 $ClipboardModes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ClipboardMode", [string], $ClipboardModesCollection)
 $CustomAttributes.HelpMessage = "Enter the drag n' drop mode for the virtual machine"
 $DndModesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $DndModesCollection.Add($CustomAttributes)
 $ValidateSetDndModes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disabled','HostToGuest','GuestToHost','Bidirectional'))
 if ($global:systempropertiessupported.DndModes) {
  $ValidateSetDndModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.DndModes)
 }
 $DndModesCollection.Add($ValidateSetDndModes)
 $DndModes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("DndMode", [string], $DndModesCollection)
 $CustomAttributes.HelpMessage = 'Enter the firmware type for the virtual machine'
 $FirmwareTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $FirmwareTypesCollection.Add($CustomAttributes)
 $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('BIOS','EFI','EFI32','EFI64','EFIDUAL'))
 if ($global:systempropertiessupported.FirmwareTypes) {
  $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.FirmwareTypes)
 }
 $FirmwareTypesCollection.Add($ValidateSetFirmwareTypes)
 $FirmwareTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("FirmwareType", [string], $FirmwareTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the pointing HID type for the virtual machine'
 $PointingHidTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $PointingHidTypesCollection.Add($CustomAttributes)
 $ValidateSetPointingHidTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('PS2Mouse','USBTablet','USBMultiTouch'))
 if ($global:systempropertiessupported.PointingHidTypes) {
  $ValidateSetPointingHidTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.PointingHidTypes)
 }
 $PointingHidTypesCollection.Add($ValidateSetPointingHidTypes)
 $PointingHidTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("PointingHidType", [string], $PointingHidTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the keyboard HID type for the virtual machine'
 $KeyboardHidTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $KeyboardHidTypesCollection.Add($CustomAttributes)
 $ValidateSetKeyboardHidTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('PS2Keyboard','USBKeyboard'))
 if ($global:systempropertiessupported.KeyboardHidTypes) {
  $ValidateSetKeyboardHidTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.KeyboardHidTypes)
 }
 $KeyboardHidTypesCollection.Add($ValidateSetKeyboardHidTypes)
 $KeyboardHidTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("KeyboardHidType", [string], $KeyboardHidTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the Virtual File System type for the virtual machine'
 $VfsTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $VfsTypesCollection.Add($CustomAttributes)
 $ValidateSetVfsTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('File','Cloud','S3'))
 if ($global:systempropertiessupported.VfsTypes) {
  $ValidateSetVfsTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.VfsTypes)
 }
 $VfsTypesCollection.Add($ValidateSetVfsTypes)
 $VfsTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("VfsType", [string], $VfsTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording audio codec for the virtual machine'
 $RecordingAudioCodecsCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingAudioCodecsCollection.Add($CustomAttributes)
 $ValidateSetRecordingAudioCodecs = New-Object System.Management.Automation.ValidateSetAttribute(@('Opus'))
 if ($global:systempropertiessupported.RecordingAudioCodecs) {
  $ValidateSetRecordingAudioCodecs = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingAudioCodecs)
 }
 $RecordingAudioCodecsCollection.Add($ValidateSetRecordingAudioCodecs)
 $RecordingAudioCodecs = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingAudioCodec", [string], $RecordingAudioCodecsCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording video codec for the virtual machine'
 $RecordingVideoCodecsCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingVideoCodecsCollection.Add($CustomAttributes)
 $ValidateSetRecordingVideoCodecs = New-Object System.Management.Automation.ValidateSetAttribute(@('VP8'))
 if ($global:systempropertiessupported.RecordingVideoCodecs) {
  $ValidateSetRecordingVideoCodecs = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingVideoCodecs)
 }
 $RecordingVideoCodecsCollection.Add($ValidateSetRecordingVideoCodecs)
 $RecordingVideoCodecs = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingVideoCodec", [string], $RecordingVideoCodecsCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording VS codec for the virtual machine'
 $RecordingVsMethodsCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingVsMethodsCollection.Add($CustomAttributes)
 $ValidateSetRecordingVsMethods = New-Object System.Management.Automation.ValidateSetAttribute(@('None'))
 if ($global:systempropertiessupported.RecordingVsMethods) {
  $ValidateSetRecordingVsMethods = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingVsMethods)
 }
 $RecordingVsMethodsCollection.Add($ValidateSetRecordingVsMethods)
 $RecordingVsMethods = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingVsMethod", [string], $RecordingVsMethodsCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording VRC mode for the virtual machine'
 $RecordingVrcModesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingVrcModesCollection.Add($CustomAttributes)
 $ValidateSetRecordingVrcModes = New-Object System.Management.Automation.ValidateSetAttribute(@('CBR'))
 if ($global:systempropertiessupported.RecordingVrcModes) {
  $ValidateSetRecordingVrcModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingVrcModes)
 }
 $RecordingVrcModesCollection.Add($ValidateSetRecordingVrcModes)
 $RecordingVrcModes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingVrcMode", [string], $RecordingVrcModesCollection)
 $CustomAttributes.HelpMessage = 'Enter the graphics controller type for the virtual machine'
 $GraphicsControllerTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $GraphicsControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetGraphicsControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('VBoxVGA','VMSVGA','VBoxSVGA','Null'))
 if ($global:systempropertiessupported.GraphicsControllerTypes) {
  $ValidateSetGraphicsControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.GraphicsControllerTypes)
 }
 $GraphicsControllerTypesCollection.Add($ValidateSetGraphicsControllerTypes)
 $GraphicsControllerTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("GraphicsControllerType", [string], $GraphicsControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the auto stop type for the virtual machine'
 $AutostopTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $AutostopTypesCollection.Add($CustomAttributes)
 $ValidateSetAutostopTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disabled','SaveState','PowerOff','AcpiShutdown'))
 if ($global:systempropertiessupported.AutostopTypes) {
  $ValidateSetAutostopTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.AutostopTypes)
 }
 $AutostopTypesCollection.Add($ValidateSetAutostopTypes)
 $AutostopTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("AutostopType", [string], $AutostopTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the VM process priority for the virtual machine'
 $VmProcPrioritiesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $VmProcPrioritiesCollection.Add($CustomAttributes)
 $ValidateSetVmProcPriorities = New-Object System.Management.Automation.ValidateSetAttribute(@('Default','Flat','Low','Normal','High'))
 if ($global:systempropertiessupported.VmProcPriorities) {
  $ValidateSetVmProcPriorities = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.VmProcPriorities)
 }
 $VmProcPrioritiesCollection.Add($ValidateSetVmProcPriorities)
 $VmProcPriorities = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("VmProcPriority", [string], $VmProcPrioritiesCollection)
 $CustomAttributes.HelpMessage = 'Enter the network attachment type for the virtual machine'
 $NetworkAttachmentTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $NetworkAttachmentTypesCollection.Add($CustomAttributes)
 $ValidateSetNetworkAttachmentTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('NAT','Bridged','Internal','HostOnly','Generic','NATNetwork','Null'))
 if ($global:systempropertiessupported.NetworkAttachmentTypes) {
  $ValidateSetNetworkAttachmentTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.NetworkAttachmentTypes)
 }
 $NetworkAttachmentTypesCollection.Add($ValidateSetNetworkAttachmentTypes)
 $NetworkAttachmentTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("NetworkAttachmentType", [string], $NetworkAttachmentTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the network adapter type for the virtual machine'
 $NetworkAdapterTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $NetworkAdapterTypesCollection.Add($CustomAttributes)
 $ValidateSetNetworkAdapterTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('Am79C970A','Am79C973','I82540EM','I82543GC','I82545EM','Virtio','Am79C960'))
 if ($global:systempropertiessupported.NetworkAdapterTypes) {
  $ValidateSetNetworkAdapterTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.NetworkAdapterTypes)
 }
 $NetworkAdapterTypesCollection.Add($ValidateSetNetworkAdapterTypes)
 $NetworkAdapterTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("NetworkAdapterType", [string], $NetworkAdapterTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the port mode for the virtual machine'
 $PortModesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $PortModesCollection.Add($CustomAttributes)
 $ValidateSetPortModes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disconnected','HostPipe','HostDevice','RawFile','TCP'))
 if ($global:systempropertiessupported.PortModes) {
  $ValidateSetPortModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.PortModes)
 }
 $PortModesCollection.Add($ValidateSetPortModes)
 $PortModes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("PortMode", [string], $PortModesCollection)
 $CustomAttributes.HelpMessage = 'Enter the emulated UART implementation type for the virtual machine'
 $UartTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $UartTypesCollection.Add($CustomAttributes)
 $ValidateSetUartTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('U16450','U16550A','U16750'))
 if ($global:systempropertiessupported.UartTypes) {
  $ValidateSetUartTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.UartTypes)
 }
 $UartTypesCollection.Add($ValidateSetUartTypes)
 $UartTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("UartType", [string], $UartTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the USB controller type for the virtual machine'
 $UsbControllerTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $UsbControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetUsbControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('OHCI','EHCI','XHCI'))
 if ($global:systempropertiessupported.UsbControllerTypes) {
  $ValidateSetUsbControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.UsbControllerTypes)
 }
 $UsbControllerTypesCollection.Add($ValidateSetUsbControllerTypes)
 $UsbControllerTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("UsbControllerType", [string], $UsbControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the audio driver type for the virtual machine'
 $AudioDriverTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $AudioDriverTypesCollection.Add($CustomAttributes)
 $ValidateSetAudioDriverTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('DirectSound','Null'))
 if ($global:systempropertiessupported.AudioDriverTypes) {
  $ValidateSetAudioDriverTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.AudioDriverTypes)
 }
 $AudioDriverTypesCollection.Add($ValidateSetAudioDriverTypes)
 $AudioDriverTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("AudioDriverType", [string], $AudioDriverTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the audio controller type for the virtual machine'
 $AudioControllerTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $AudioControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetAudioControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('AC97','SB16','HDA'))
 if ($global:systempropertiessupported.AudioControllerTypes) {
  $ValidateSetAudioControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.AudioControllerTypes)
 }
 $AudioControllerTypesCollection.Add($ValidateSetAudioControllerTypes)
 $AudioControllerTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("AudioControllerType", [string], $AudioControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the storage bus for the virtual machine'
 $StorageBusesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $StorageBusesCollection.Add($CustomAttributes)
 $ValidateSetStorageBuses = New-Object System.Management.Automation.ValidateSetAttribute(@('SATA','IDE','SCSI','Floppy','SAS','USB','PCIe','VirtioSCSI'))
 if ($global:systempropertiessupported.StorageBuses) {
  $ValidateSetStorageBuses = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.StorageBuses)
 }
 $StorageBusesCollection.Add($ValidateSetStorageBuses)
 $StorageBuses = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("StorageBus", [string], $StorageBusesCollection)
 $CustomAttributes.HelpMessage = 'Enter the storage controller type for the virtual machine'
 $StorageControllerTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $StorageControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetStorageControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('IntelAhci','PIIX4','PIIX3','ICH6','LsiLogic','BusLogic','I82078','LsiLogicSas','USB','NVMe','VirtioSCSI'))
 if ($global:systempropertiessupported.StorageControllerTypes) {
  $ValidateSetStorageControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.StorageControllerTypes)
 }
 $StorageControllerTypesCollection.Add($ValidateSetStorageControllerTypes)
 $StorageControllerTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("StorageControllerType", [string], $StorageControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the chipset type for the virtual machine'
 $ChipsetTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $ChipsetTypesCollection.Add($CustomAttributes)
 $ValidateSetChipsetTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('PIIX3','ICH9'))
 if ($global:systempropertiessupported.ChipsetTypes) {
  $ValidateSetChipsetTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.ChipsetTypes)
 }
 $ChipsetTypesCollection.Add($ValidateSetChipsetTypes)
 $ChipsetTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ChipsetType", [string], $ChipsetTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the number of CPUs available to the virtual machine'
 $CpuCountCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $CpuCountCollection.Add($CustomAttributes)
 $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute(1, 32)
 if ($global:systempropertiessupported.MinGuestCPUCount -and $global:systempropertiessupported.MaxGuestCPUCount) {
  $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestCPUCount, $global:systempropertiessupported.MaxGuestCPUCount)
 }
 $CpuCountCollection.Add($ValidateSetCpuCount)
 $CpuCount = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("CpuCount", [uint64], $CpuCountCollection)
 $CustomAttributes.HelpMessage = 'Enter the memory size in MB for the virtual machine'
 $MemorySizeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $MemorySizeCollection.Add($CustomAttributes)
 $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute(4, 2097152)
 if ($global:systempropertiessupported.MinGuestRam -and $global:systempropertiessupported.MaxGuestRam) {
  $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestRam, $global:systempropertiessupported.MaxGuestRam)
 }
 $MemorySizeCollection.Add($ValidateSetMemorySize)
 $MemorySize = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("MemorySize", [uint64], $MemorySizeCollection)
 $CustomAttributes.HelpMessage = 'Enter the memory balloon size in MB for the virtual machine'
 $MemoryBalloonSizeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $MemoryBalloonSizeCollection.Add($CustomAttributes)
 $ValidateSetMemoryBalloonSize = New-Object System.Management.Automation.ValidateRangeAttribute(4, 2097152)
 if ($global:systempropertiessupported.MinGuestRam -and $global:systempropertiessupported.MaxGuestRam) {
  $ValidateSetMemoryBalloonSize = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestRam, $global:systempropertiessupported.MaxGuestRam)
 }
 $MemoryBalloonSizeCollection.Add($ValidateSetMemoryBalloonSize)
 $MemoryBalloonSize = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("MemoryBalloonSize", [uint64], $MemoryBalloonSizeCollection)
 $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
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
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
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
 if (!$global:guestostype) {Write-Host "[Error] Could not find guest defaults. Run Start-VirtualBoxSession with the -Force switch and try again." -ForegroundColor Red -BackgroundColor Black;return}
 $defaultsettings = $global:guestostype | Where-Object {$_.id -eq $OsTypeId}
 if ((Get-VirtualBoxVM -Name $Name -SkipCheck).Name -eq $Name) {Write-Host "[Error] Machine $Name already exists. Enter another name and try again." -ForegroundColor Red -BackgroundColor Black;return}
 try {
  # create a reference object for the new machine
  Write-Verbose "Creating reference object for $Name"
  $imachine = New-Object VirtualBoxVM
  if ($ModuleHost.ToLower() -eq 'websrv') {
   $imachine.Id = $global:vbox.IVirtualBox_createMachine($global:ivbox, $Location, $Name, $Group, $OsTypeId, $Flags)
   $global:vbox.IMachine_applyDefaults($imachine.Id, $null)
   if ($PsCmdlet.ParameterSetName -eq 'Custom') {
    try {
     if ($Icon) {
      if (!(Test-Path "$env:TEMP\VirtualBoxPS")) {New-Item -ItemType Directory -Path "$env:TEMP\VirtualBoxPS\" -Force -Confirm:$false | Write-Verbose}
      # convert to png
      Add-Type -AssemblyName system.drawing
      $imageFormat = "System.Drawing.Imaging.ImageFormat" -as [type]
      $image = [drawing.image]::FromFile($Icon)
      $image.Save("$env:TEMP\VirtualBoxPS\icon.png", $imageFormat::Png)
      $octet = [convert]::ToBase64String((Get-Content "$env:TEMP\VirtualBoxPS\icon.png" -Encoding Byte))
      $global:vbox.IMachine_setIcon($imachine.Id, $octet)
      Remove-Item -Path "$env:TEMP\VirtualBoxPS\icon.png" -Confirm:$false -Force
     }
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
     if ($CpuProfile) {$global:vbox.IMachine_setCPUProfile($imachine.Id, $CpuProfile)}
    }
    catch {
     Write-Verbose 'Exception applying custom parameters to machine'
     Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
     Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
    }
   }
   $global:vbox.IMachine_saveSettings($imachine.Id)
   $global:vbox.IVirtualBox_registerMachine($global:ivbox, $imachine.Id)
  } # end if websrv
  elseif ($ModuleHost.ToLower() -eq 'com') {
   Write-Verbose "Location: `"$Location`""
   Write-Verbose "Name: `"$Name`""
   Write-Verbose "Group: `"$Group`""
   Write-Verbose "OsTypeId: `"$OsTypeId`""
   Write-Verbose "Flags: `"$Flags`""
   if (!$Location) {$Location = ''}
   if (!$Group) {$Group = '/'}
   if (!$Flags) {$Flags = ''}
   $imachine.ComObject = $global:vbox.CreateMachine($Location, $Name, [string[]]@($Group), $OsTypeId, $Flags)
   $imachine.ComObject.ApplyDefaults($null)
   if ($PsCmdlet.ParameterSetName -eq 'Custom') {
    try {
     if ($Icon) {
      if (!(Test-Path "$env:TEMP\VirtualBoxPS")) {New-Item -ItemType Directory -Path "$env:TEMP\VirtualBoxPS\" -Force -Confirm:$false | Write-Verbose}
      # convert to png
      Add-Type -AssemblyName system.drawing
      $imageFormat = "System.Drawing.Imaging.ImageFormat" -as [type]
      $image = [drawing.image]::FromFile($Icon)
      $image.Save("$env:TEMP\VirtualBoxPS\icon.png", $imageFormat::Png)
      [byte[]]$bytes = Get-Content "$env:TEMP\VirtualBoxPS\icon.png" -Encoding Byte
      $imachine.ComObject.Icon($bytes)
      Remove-Item -Path "$env:TEMP\VirtualBoxPS\icon.png" -Confirm:$false -Force
     }
     if ($Description) {$imachine.ComObject.Description = $Description}
     if ($HardwareUuid) {$imachine.ComObject.HardwareUUID = $HardwareUuid}
     if ($CpuCount) {$imachine.ComObject.CPUCount = $CpuCount}
     if ($CpuHotPlugEnabled) {$imachine.ComObject.CPUHotPlugEnabled = [int]$CpuHotPlugEnabled}
     if ($CpuExecutionCap) {$imachine.ComObject.CPUExecutionCap = $CpuExecutionCap}
     if ($CpuIdPortabilityLevel) {$imachine.ComObject.CPUIDPortabilityLevel = $CpuIdPortabilityLevel}
     if ($MemorySize) {$imachine.ComObject.MemorySize = $MemorySize}
     if ($MemoryBalloonSize) {$imachine.ComObject.MemoryBalloonSize = $MemoryBalloonSize}
     if ($PageFusionEnabled) {$imachine.ComObject.PageFusionEnabled = [int]$PageFusionEnabled}
     if ($FirmwareType) {$imachine.ComObject.FirmwareType = [FirmwareType]::new().ToInt($FirmwareType)}
     if ($PointingHidType) {$imachine.ComObject.PointingHIDType = [PointingHIDType]::new().ToInt($PointingHidType - 1)}
     if ($KeyboardHidType) {$imachine.ComObject.KeyboardHIDType = [KeyboardHIDType]::new().ToInt($KeyboardHidType - 1)}
     if ($HpetEnabled) {$imachine.ComObject.HPETEnabled = [int]$HpetEnabled}
     if ($ChipsetType) {$imachine.ComObject.ChipsetType = [ChipsetType]::new().ToInt($ChipsetType)}
     if ($EmulatedUsbCardReaderEnabled) {$imachine.ComObject.EmulatedUSBCardReaderEnabled = [int]$EmulatedUsbCardReaderEnabled}
     if ($ClipboardMode) {$imachine.ComObject.ClipboardMode = [ClipboardMode]::new().ToInt($ClipboardMode)}
     if ($ClipboardFileTransfersEnabled) {$imachine.ComObject.ClipboardFileTransfersEnabled = [int]$ClipboardFileTransfersEnabled}
     if ($DndMode) {$imachine.ComObject.DnDMode = [DnDMode]::new().ToInt($DndMode)}
     if ($TeleporterEnabled) {$imachine.ComObject.TeleporterEnabled = [int]$TeleporterEnabled}
     if ($TeleporterPort) {$imachine.ComObject.TeleporterPort = $TeleporterPort}
     if ($TeleporterAddress) {$imachine.ComObject.TeleporterAddress = $TeleporterAddress}
     if ($TeleporterPassword) {$imachine.ComObject.TeleporterPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($TeleporterPassword))}
     if ($ParavirtProvider) {$imachine.ComObject.ParavirtProvider = [ParavirtProvider]::new().ToInt($ParavirtProvider)}
     if ($RtcUseUtc) {$imachine.ComObject.RTCUseUTC = [int]$RtcUseUtc}
     if ($IoCacheEnabled) {$imachine.ComObject.IOCacheEnabled = [int]$IoCacheEnabled}
     if ($IoCacheSize) {$imachine.ComObject.IOCacheSize = $IoCacheSize}
     if ($TracingEnabled) {$imachine.ComObject.TracingEnabled = [int]$TracingEnabled}
     if ($TracingConfig) {$imachine.ComObject.TracingConfig = $TracingConfig}
     if ($AllowTracingToAccessVM) {$imachine.ComObject.AllowTracingToAccessVM = [int]$AllowTracingToAccessVM}
     if ($AutostartEnabled) {$imachine.ComObject.AutostartEnabled = [int]$AutostartEnabled}
     if ($AutostartDelay) {$imachine.ComObject.AutostartDelay = $AutostartDelay}
     if ($AutostopType) {$imachine.ComObject.AutostopType = [AutostopType]::new().ToInt($AutostopType - 1)}
     if ($CpuProfile) {$imachine.ComObject.CPUProfile = $CpuProfile}
    }
    catch {
     Write-Verbose 'Exception applying custom parameters to machine'
     Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
     Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
    }
   }
   $imachine.ComObject.SaveSettings()
   $global:vbox.RegisterMachine($imachine.ComObject)
  } # end elseif com
 } # Try
 catch {
  Write-Verbose 'Exception creating machine'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Remove-VirtualBoxVM {
<#
.SYNOPSIS
Remove a virtual machine
.DESCRIPTION
Removes a new virtual machine from inventory. This command requires confirmation before taking any destructive actions. Use the -Confirm:$false parameter to silence all confirmation prompts. The name provided by the Name parameter must exist in the VirtualBox inventory, or this command will fail. By default, this command will unregister the machine from the VirtualBox inventory. Optionally you can use one of the three provided switches (DetachAllReturnNone, DetachAllReturnHardDisksOnly, and Full) to remove media attached to the machine or even delete them. See the individual parameter help below for more information.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER DetachAllReturnNone
A switch to delete all snapshots and detach all media. This will keep all media registered.
.PARAMETER DetachAllReturnHardDisksOnly
A switch to delete all snapshots, detach all media and return hard disks for deletion, but not removable media.
.PARAMETER Full
A switch to delete all snapshots, detach all media and return all media for deletion. Removable media will be removed from the VirtualBox inventory.
.PARAMETER ProgressBar
A switch to display a progress bar when deleting files.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Remove-VirtualBoxVM -Name "My VM I Hate"
Removes the virtual machine named "My VM I Hate" from the VirtualBox inventory
.EXAMPLE
PS C:\> Remove-VirtualBoxVM -Name "My VM I Hate" -DetachAllReturnNone
Deletes the virtual machine named "My VM I Hate" from the host machine and deletes its snapshots and detaches its media
.EXAMPLE
PS C:\> Remove-VirtualBoxVM -Name "My VM I Hate" -DetachAllReturnHardDisksOnly
Deletes the virtual machine named "My VM I Hate" from the host machine and deletes its snapshots and all attached disks
.EXAMPLE
PS C:\> Remove-VirtualBoxVM -Name "My VM I Hate" -Full
Deletes the virtual machine named "My VM I Hate" from the host machine and deletes its snapshots and all attached persistent media
.NOTES
NAME        :  Remove-VirtualBoxVM
VERSION     :  1.0
LAST UPDATED:  1/16/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
New-VirtualBoxVM
Import-VirtualBoxVM
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
.OUTPUTS
None
#>
[CmdletBinding(DefaultParameterSetName='UnregisterOnly',SupportsShouldProcess,ConfirmImpact='High')]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)"
,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(ParameterSetName='DetachAllReturnNone',Mandatory=$true,
HelpMessage="Use this switch to delete all snapshots and detach all media but return none; this will keep all media registered")]
  [switch]$DetachAllReturnNone,
[Parameter(ParameterSetName='DetachAllReturnHardDisksOnly',Mandatory=$true,
HelpMessage="Use this switch to delete all snapshots, detach all media and return hard disks for closing, but not removable media")]
  [switch]$DetachAllReturnHardDisksOnly,
[Parameter(ParameterSetName='Full',Mandatory=$true,
HelpMessage="Use this switch to delete all snapshots, detach all media and return all media for closing")]
  [switch]$Full,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 if ($imachines) {
  try {
   foreach ($imachine in $imachines) {
    if ($imachine.State -ne 'PoweredOff') {Write-Host "[Error] Machine $($imachine.Name) is not powered off. Power it off and try again." -ForegroundColor Red -BackgroundColor Black;return}
    if ($ModuleHost.ToLower() -eq 'websrv') {
     if ($PSCmdlet.ParameterSetName -eq 'DetachAllReturnNone' -and $PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Delete virtual machine from host, delete all snapshots, and detach all media from virtual machine ")) {
      Write-Verbose "Removing virtual machine $($imachine.Name) from inventory"
      $imediums = $global:vbox.IMachine_unregister($imachine.Id, [CleanupMode]::new().ToULong('DetachAllReturnNone'))
      # delete VM files
      Write-Verbose "Running cleanup action $($PSCmdlet.ParameterSetName) for $($imachine.Name) machine's files"
      $imachine.IProgress.Id = $global:vbox.IMachine_deleteConfig($imachine.Id, $imediums)
     } # end if ParameterSetName -eq DetachAllReturnNone
     elseif ($PSCmdlet.ParameterSetName -eq 'DetachAllReturnHardDisksOnly' -and $PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Delete virtual machine from host, delete all snapshots, and delete all virtual disks but keep removable media in inventory ")) {
      Write-Verbose "Removing virtual machine $($imachine.Name) from inventory"
      $imediums = $global:vbox.IMachine_unregister($imachine.Id, [CleanupMode]::new().ToULong('DetachAllReturnHardDisksOnly'))
      # delete VM files and virtual disk(s)
      Write-Verbose "Running cleanup action $($PSCmdlet.ParameterSetName) for $($imachine.Name) machine's files"
      $imachine.IProgress.Id = $global:vbox.IMachine_deleteConfig($imachine.Id, $imediums)
     } # end elseif ParameterSetName -eq DetachAllReturnHardDisksOnly
     elseif ($PSCmdlet.ParameterSetName -eq 'Full' -and $PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Delete virtual machine from host, delete all snapshots, delete all virtual disks, and remove removable media from inventory ")) {
      Write-Verbose "Removing virtual machine $($imachine.Name) from inventory"
      $imediums = $global:vbox.IMachine_unregister($imachine.Id, [CleanupMode]::new().ToULong('Full'))
      # delete VM files and virtual disk(s)
      Write-Verbose "Running cleanup action $($PSCmdlet.ParameterSetName) for $($imachine.Name) machine's files"
      $imachine.IProgress.Id = $global:vbox.IMachine_deleteConfig($imachine.Id, $imediums)
      # Remove-VirtualBoxDisc command goes here
     } # end elseif ParameterSetName -eq Full
     elseif ($PSCmdlet.ParameterSetName -eq 'UnregisterOnly' -and $PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Remove virtual machine from inventory ")) {
      Write-Verbose "Removing virtual machine $($imachine.Name) from inventory"
      $imediums = $global:vbox.IMachine_unregister($imachine.Id, [CleanupMode]::new().ToULong('UnregisterOnly'))
     } # end elseif ParameterSetName -eq UnregisterOnly
     if ($imachine.IProgress.Id) {
      Write-Verbose 'Displaying progress bar'
      if ($ProgressBar) {Write-Progress -Activity "Removing virtual machine $($imachine.Name) ($($PSCmdlet.ParameterSetName))" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)}
      do {
       # update iprogress data
       $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
       if ($ProgressBar) {Write-Progress -Activity "Removing virtual machine $($imachine.Name) ($($PSCmdlet.ParameterSetName))" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)}
       if ($ProgressBar) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
      } until ($imachine.IProgress.Completed -eq $true) # continue once completed
      if ($imachine.IProgress.ResultCode -ne 0) {Write-Verbose $imachine.IProgress.ErrorInfo}
     }
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     if ($PSCmdlet.ParameterSetName -eq 'DetachAllReturnNone' -and $PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Delete virtual machine from host, delete all snapshots, and detach all media from virtual machine ")) {
      Write-Verbose "[Warning] Configuration cleanup has been disabled due to a bug in the VirtualBox COM"
      Write-Verbose "          Running cleanup manually"
      # this mess will go away when the com is fixed, or when Get/Remove-VirtualBoxSnapshot commands are created
      $Location = ($vbox.Machines | Where-Object {$_.Name -eq $imachine.Name}).SettingsFilePath
      $imediumattachments = ($vbox.Machines | Where-Object {$_.Name -eq $imachine.Name}).MediumAttachments | Where-Object {$_.Medium.Id -ne $null}
      foreach ($imediumattachment in $imediumattachments) {
      $imediums = $imediumattachment.Medium
       if ($imediums) {
        foreach ($imedium in $imediums) {
         Write-Verbose "Dismounting virtual disk: $($imediumattachment.Medium.Name) from $($imachine.Name) machine"
         Write-Verbose "Controller: $($imediumattachment.Controller)"
         Write-Verbose "ControllerPort: $($imediumattachment.Port)"
         Write-Verbose "ControllerSlot: $($imediumattachment.Device)"
         Dismount-VirtualBoxDisk -Guid $imediumattachment.Medium.Id -MachineGuid $imachine.Guid
        } # foreach $imedium in $imediums
       } # end if $imediums
      } # foreach $imediumattachment in $imediumattachments
      Write-Verbose "Removing virtual machine $($imachine.Name) from inventory"
      $imediums = $imachine.ComObject.Unregister([CleanupMode]::new().ToULong('DetachAllReturnNone'))
      # delete VM files
      Write-Verbose "Running cleanup action $($PSCmdlet.ParameterSetName) for $($imachine.Name) machine's files"
      #$imachine.IProgress.Progress = $imachine.ComObject.DeleteConfig($imediums)
      # additional temp cleanup
      $Location = $Location.Substring(0,$Location.LastIndexOf('\'))
      Remove-Item -Path $Location -Recurse -Force -Confirm:$false
     } # end if ParameterSetName -eq DetachAllReturnNone
     elseif ($PSCmdlet.ParameterSetName -eq 'DetachAllReturnHardDisksOnly' -and $PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Delete virtual machine from host, delete all snapshots, and delete all virtual disks but keep removable media in inventory ")) {
      Write-Verbose "[Warning] Configuration cleanup has been disabled due to a bug in the VirtualBox COM"
      Write-Verbose "          Running cleanup manually"
      # this mess will go away when the com is fixed, or when Get/Remove-VirtualBoxSnapshot commands are created
      $Location = ($vbox.Machines | Where-Object {$_.Name -eq $imachine.Name}).SettingsFilePath
      $imediumattachments = ($vbox.Machines | Where-Object {$_.Name -eq $imachine.Name}).MediumAttachments | Where-Object {$_.Medium.Id -ne $null}
      foreach ($imediumattachment in $imediumattachments) {
       $imediums = $imediumattachment.Medium
       if ($imediums) {
        foreach ($imedium in $imediums) {
         Write-Verbose "Dismounting virtual disk: $($imediumattachment.Medium.Name) from $($imachine.Name) machine"
         Write-Verbose "Controller: $($imediumattachment.Controller)"
         Write-Verbose "ControllerPort: $($imediumattachment.Port)"
         Write-Verbose "ControllerSlot: $($imediumattachment.Device)"
         Dismount-VirtualBoxDisk -Guid $imediumattachment.Medium.Id -MachineGuid $imachine.Guid
        } # foreach $imedium in $imediums
        if ($ProgressBar) {Remove-VirtualBoxDisk -Name $imedium.Name -DeleteFromHost -ProgressBar -Confirm:$false -SkipCheck}
        else {Remove-VirtualBoxDisk -Name $imedium.Name -DeleteFromHost -Confirm:$false -SkipCheck}
       } # end if $imediums
      } # foreach $imediumattachment in $imediumattachments
      Write-Verbose "Removing virtual machine $($imachine.Name) from inventory"
      $imediums = $imachine.ComObject.Unregister([CleanupMode]::new().ToULong('DetachAllReturnHardDisksOnly'))
      # delete VM files and virtual disk(s)
      Write-Verbose "Running cleanup action $($PSCmdlet.ParameterSetName) for $($imachine.Name) machine's files"
      #$imachine.IProgress.Progress = $imachine.ComObject.DeleteConfig($imediums)
      # additional temp cleanup
      $Location = $Location.Substring(0,$Location.LastIndexOf('\'))
      Remove-Item -Path $Location -Recurse -Force -Confirm:$false
     } # end elseif ParameterSetName -eq DetachAllReturnHardDisksOnly
     elseif ($PSCmdlet.ParameterSetName -eq 'Full' -and $PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Delete virtual machine from host, delete all snapshots, delete all virtual disks, and remove removable media from inventory ")) {
      Write-Verbose "[Warning] Configuration cleanup has been disabled due to a bug in the VirtualBox COM"
      Write-Verbose "          Running cleanup manually"
      # this mess will go away when the com is fixed, or when Get/Remove-VirtualBoxSnapshot commands are created
      $Location = ($vbox.Machines | Where-Object {$_.Name -eq $imachine.Name}).SettingsFilePath
      $imediumattachments = ($vbox.Machines | Where-Object {$_.Name -eq $imachine.Name}).MediumAttachments | Where-Object {$_.Medium.Id -ne $null}
      foreach ($imediumattachment in $imediumattachments) {
       $imediums = $imediumattachment.Medium
       if ($imediums) {
        foreach ($imedium in $imediums) {
         Write-Verbose "Dismounting virtual disk: $($imediumattachment.Medium.Name) from $($imachine.Name) machine"
         Write-Verbose "Controller: $($imediumattachment.Controller)"
         Write-Verbose "ControllerPort: $($imediumattachment.Port)"
         Write-Verbose "ControllerSlot: $($imediumattachment.Device)"
         Dismount-VirtualBoxDisk -Guid $imediumattachment.Medium.Id -SkipCheck
        } # foreach $imedium in $imediums
        if ($imedium.Format -match 'RAW') {
         if ($ProgressBar) {Remove-VirtualBoxDisc -Guid $imedium.Id -ProgressBar -SkipCheck}
         else {Remove-VirtualBoxDisc -Guid $imedium.Id -SkipCheck}
        } # end if CD/DVD
        else {
         if ($ProgressBar) {Remove-VirtualBoxDisk -Guid $imedium.Id -DeleteFromHost -ProgressBar -Confirm:$false -SkipCheck}
         else {Remove-VirtualBoxDisk -Guid $imedium.Id -DeleteFromHost -Confirm:$false -SkipCheck}
        } # end else
       } # end if $imediums
      } # foreach $imediumattachment in $imediumattachments
      Write-Verbose "Removing virtual machine $($imachine.Name) from inventory"
      $imediums = $imachine.ComObject.Unregister([CleanupMode]::new().ToULong('Full'))
      # delete VM files and virtual disk(s)
      Write-Verbose "Running cleanup action $($PSCmdlet.ParameterSetName) for $($imachine.Name) machine's files"
      #$imachine.IProgress.Progress = $imachine.ComObject.DeleteConfig($imediums)
      # additional temp cleanup
      $Location = $Location.Substring(0,$Location.LastIndexOf('\'))
      Remove-Item -Path $Location -Recurse -Force -Confirm:$false
      # Remove-VirtualBoxDisc command goes here
     } # end elseif ParameterSetName -eq Full
     elseif ($PSCmdlet.ParameterSetName -eq 'UnregisterOnly' -and $PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Remove virtual machine from inventory ")) {
      Write-Verbose "Removing virtual machine $($imachine.Name) from inventory"
      $imediums = $imachine.ComObject.Unregister([CleanupMode]::new().ToULong('UnregisterOnly'))
     } # end elseif ParameterSetName -eq UnregisterOnly
     if ($imachine.IProgress.Progress) {
      Write-Verbose 'Displaying progress bar'
      if ($ProgressBar) {Write-Progress -Activity "Removing virtual machine $($imachine.Name) ($($PSCmdlet.ParameterSetName))" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.Progress.TimeRemaining)}
      do {
       # update iprogress data
       if ($ProgressBar) {Write-Progress -Activity "Removing virtual machine $($imachine.Name) ($($PSCmdlet.ParameterSetName))" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.Progress.TimeRemaining)}
       if ($ProgressBar) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
      } until ($imachine.IProgress.Progress.Percent -eq 100 -and $imachine.IProgress.Progress.Completed -eq $true) # continue once completed
      if ($imachine.IProgress.Progress.ResultCode -ne 0) {Write-Verbose $imachine.IProgress.Progress.ErrorInfo}
     }
    } # end elseif com
   } # foreach $imachine in $imachines
  } # Try
  catch {
   Write-Verbose 'Exception removing machine'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  finally {
   # obligatory session unlock
   Write-Verbose 'Cleaning up machine sessions'
   if ($imachines) {
    foreach ($imachine in $imachines) {
     if ($imachine.ISession.Id) {
      if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
       Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
       $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      } # end if session state not unlocked
     } # end if $imachine.ISession.Id
     if ($imachine.ISession.Session) {
      if ($imachine.ISession.Session.State -gt 1) {
       $imachine.ISession.Session.UnlockMachine()
      } # end if $imachine.ISession.Session locked
     } # end if $imachine.ISession.Session
     if ($imachine.IConsole) {
      # release the iconsole session
      Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
      $global:vbox.IManagedObjectRef_release($imachine.IConsole)
     } # end if $imachine.IConsole
     #$imachine.ISession.Id = $null
     $imachine.IConsole = $null
     if ($imachine.IPercent) {$imachine.IPercent = $null}
     $imachine.MSession = $null
     $imachine.MConsole = $null
     $imachine.MMachine = $null
    } # end foreach $imachine in $imachines
   } # end if $imachines
  } # Finally
 } # end if $imachines
 else {Write-Host "[Error] No matching virtual machines were found using specified parameters" -ForegroundColor Red -BackgroundColor Black;return}
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Import-VirtualBoxVM {
<#
.SYNOPSIS
Import a virtual machine
.DESCRIPTION
Imports an existing virtual machine. The name provided by the Name parameter must not exist in the VirtualBox inventory, or this command will fail. The path provided by the Location parameter must exist, and contain the machine settings file, or the command will fail.
.PARAMETER Name
The name of the virtual machine. This is a required parameter.
.PARAMETER Location
The location of the virtual machine settings file.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Import-VirtualBoxVM -Name "My VM I Love" -Location "C:\Users\SmithersTheOracle\VirtualBox VMs\My VM I Love"
Import an existing virtual machine named "My VM I Love" from the "C:\Users\SmithersTheOracle\VirtualBox VMs\My VM I Love" folder
.NOTES
NAME        :  Import-VirtualBoxVM
VERSION     :  1.0
LAST UPDATED:  1/16/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
New-VirtualBoxVM
Remove-VirtualBoxVM
.INPUTS
String        :  String for virtual machine name
String        :  String for virtual machine path
.OUTPUTS
None
#>
[CmdletBinding()]
Param(
[Parameter(HelpMessage="Enter a virtual machine name",
Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(HelpMessage="Enter the path for the virtual machine",
ParameterSetName='Custom',Mandatory=$false)]
  [string]$Location,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 if (!(Test-Path (Join-Path -Path $Location -ChildPath "$($Name).vbox"))) {Write-Host "[Error] $(Join-Path -Path $Location -ChildPath "$($Name).vbox") does not exist. Check the path and try again." -ForegroundColor Red -BackgroundColor Black;return}
 if ((Get-VirtualBoxVM -Name $Name -SkipCheck).Name -eq $Name) {Write-Host "[Error] Machine $Name already exists. Enter another name and try again." -ForegroundColor Red -BackgroundColor Black;return}
 try {
  if ($ModuleHost.ToLower() -eq 'websrv') {
   # create a reference object for the new machine
   Write-Verbose "Creating reference object for $Name"
   $imachine = New-Object VirtualBoxVM
   Write-Verbose "Importing virtual machine $Name"
   $imachine.Id = $global:vbox.IVirtualBox_openMachine($global:ivbox, (Join-Path -Path $Location -ChildPath "$($Name).vbox"))
   Write-Verbose "Saving settings for $Name"
   $global:vbox.IMachine_saveSettings($imachine.Id)
   Write-Verbose "Registering $Name in the VirtualBox inventory"
   $global:vbox.IVirtualBox_registerMachine($global:ivbox, $imachine.Id)
  } # end if websrv
  elseif ($ModuleHost.ToLower() -eq 'com') {
   # create a reference object for the new machine
   Write-Verbose "Creating reference object for $Name"
   $imachine = New-Object VirtualBoxVM
   Write-Verbose "Importing virtual machine $Name"
   $imachine.ComObject = $global:vbox.OpenMachine((Join-Path -Path $Location -ChildPath "$($Name).vbox"))
   Write-Verbose "Saving settings for $Name"
   $imachine.ComObject.SaveSettings()
   Write-Verbose "Registering $Name in the VirtualBox inventory"
   $global:vbox.RegisterMachine($imachine.ComObject)
  } # end elseif com
 } # Try
 catch {
  Write-Verbose 'Exception importing machine'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Edit-VirtualBoxVM {
<#
.SYNOPSIS
Create a virtual machine
.DESCRIPTION
Creates a new virtual machine. The name provided by the Name parameter must not exist in the VirtualBox inventory, or this command will fail. You can optionally supply custom values using a large number of parameters available to this command. There are too many to fully document in this help text, so tab completion has been added where it is possible. The values provided by tab completion are updated when Start-VirtualBoxSession is successfully run. To force the values to be updated again, use the -Force switch with Start-VirtualBoxSession.
.PARAMETER Name
The name of the virtual machine. This is a required parameter.
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
.PARAMETER Icon
The custom icon for the virtual machine. To restore the default icon, pass a blank string to this parameter (Ex. -Icon ''). Must be a valid image file or the command will fail. Path must not be null or the command will fail.
.PARAMETER IoCacheEnabled
The Enable or disable IO cache for the virtual machine.
.PARAMETER IoCacheSize
The IO cache size in MB for the virtual machine.
.PARAMETER KeyboardHidType
The keyboard HID type for the virtual machine.
.PARAMETER Location
The location for the virtual machine files.
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
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Edit-VirtualBoxVM -Name "My New Win10 VM" -OsTypeId Windows10_64
Create a new virtual machine named "My New Win10 VM" with the all the recommended 64bit Windows10 defaults
.NOTES
NAME        :  Edit-VirtualBoxVM
VERSION     :  1.0
LAST UPDATED:  1/15/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Remove-VirtualBoxVM
Import-VirtualBoxVM
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for virtual machine OS Type ID
Other optional input parameters available. Use "Get-Help Edit-VirtualBoxVM -Full" for a complete list.
.OUTPUTS
None
#>
[CmdletBinding()]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)"
,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Enter the full path to the icon for the virtual machine",
ParameterSetName='Custom',Mandatory=$false)]
[ValidateNotNull()]
[ValidateScript({if ($_ -eq '') {return $true} elseif (Test-Path $_) {return $true}})]
  [string]$Icon,
[Parameter(HelpMessage="Enter the description for the virtual machine",
Mandatory=$false)]
  [string]$Description,
[Parameter(HelpMessage="Enter the hardware UUID for the virtual machine",
Mandatory=$false)]
  [guid]$HardwareUuid,
[Parameter(HelpMessage="Enable or disable CPU hotplug for the virtual machine",
Mandatory=$false)]
  [bool]$CpuHotPlugEnabled,
[Parameter(HelpMessage="Enter the CPU execution cap for the virtual machine",
Mandatory=$false)]
[ValidateRange(1, 100)]
  [uint64]$CpuExecutionCap = 100,
[Parameter(HelpMessage="Enter the CPUID portability level for the virtual machine",
Mandatory=$false)]
[ValidateRange(0, 3)]
  [uint64]$CpuIdPortabilityLevel = 0,
[Parameter(HelpMessage="Enable or disable page fusion for the virtual machine",
Mandatory=$false)]
  [bool]$PageFusionEnabled = $false,
[Parameter(HelpMessage="Enable or disable HPET for the virtual machine",
Mandatory=$false)]
  [bool]$HpetEnabled = $false,
[Parameter(HelpMessage="Enable or disable emulated USB card reader for the virtual machine",
Mandatory=$false)]
  [bool]$EmulatedUsbCardReaderEnabled = $false,
[Parameter(HelpMessage="Enable or disable clipboard file transfers for the virtual machine",
Mandatory=$false)]
  [bool]$ClipboardFileTransfersEnabled = $false,
[Parameter(HelpMessage="Enable or disable teleporter for the virtual machine",
Mandatory=$false)]
  [bool]$TeleporterEnabled = $false,
[Parameter(HelpMessage="Enter the teleporter TCP port for the virtual machine",
Mandatory=$false)]
[ValidateRange(0, 65535)]
  [uint16]$TeleporterPort = 0,
[Parameter(HelpMessage="Enter the teleporter address for the virtual machine",
Mandatory=$false)]
  [string]$TeleporterAddress = '',
[Parameter(HelpMessage="Enter the teleporter password for the virtual machine",
Mandatory=$false)]
  [securestring]$TeleporterPassword,
[Parameter(HelpMessage="Enable or disable RTC to UTC conversion for the virtual machine",
Mandatory=$false)]
  [bool]$RtcUseUtc = $false,
[Parameter(HelpMessage="Enable or disable IO cache for the virtual machine",
Mandatory=$false)]
  [bool]$IoCacheEnabled = $false,
[Parameter(HelpMessage="Enter the IO cache size in MB for the virtual machine",
Mandatory=$false)]
  [uint32]$IoCacheSize,
[Parameter(HelpMessage="Enable or disable tracing for the virtual machine",
Mandatory=$false)]
  [bool]$TracingEnabled = $false,
[Parameter(HelpMessage="Enter the tracing configuration for the virtual machine",
Mandatory=$false)]
  [string]$TracingConfig,
[Parameter(HelpMessage="Enable or disable tracing access to the virtual machine",
Mandatory=$false)]
  [bool]$AllowTracingToAccessVM = $false,
[Parameter(HelpMessage="Enable or disable auto start for the virtual machine",
Mandatory=$false)]
  [bool]$AutostartEnabled = $false,
[Parameter(HelpMessage="Enter the auto start delay in seconds for the virtual machine",
Mandatory=$false)]
  [uint32]$AutostartDelay = 300,
[Parameter(HelpMessage="Enter the CPU profile for the virtual machine",
Mandatory=$false)]
  [string]$CpuProfile,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
DynamicParam {
 $CustomAttributes = New-Object System.Management.Automation.ParameterAttribute
 $CustomAttributes.Mandatory = $false
 $CustomAttributes.HelpMessage = 'Enter the type ID for the virtual machine guest OS'
 $OsTypeIdCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $OsTypeIdCollection.Add($CustomAttributes)
 $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute(@('Other','Other_64','Windows31','Windows95','Windows98','WindowsMe','WindowsNT3x','WindowsNT4','Windows2000','WindowsXP','WindowsXP_64','Windows2003','Windows2003_64','WindowsVista','WindowsVista_64','Windows2008','Windows2008_64','Windows7','Windows7_64','Windows8','Windows8_64','Windows81','Windows81_64','Windows2012_64','Windows10','Windows10_64','Windows2016_64','Windows2019_64','WindowsNT','WindowsNT_64','Linux22','Linux24','Linux24_64','Linux26','Linux26_64','ArchLinux','ArchLinux_64','Debian','Debian_64','Fedora','Fedora_64','Gentoo','Gentoo_64','Mandriva','Mandriva_64','Oracle','Oracle_64','RedHat','RedHat_64','OpenSUSE','OpenSUSE_64','Turbolinux','Turbolinux_64','Ubuntu','Ubuntu_64','Xandros','Xandros_64','Linux','Linux_64','Solaris','Solaris_64','OpenSolaris','OpenSolaris_64','Solaris11_64','FreeBSD','FreeBSD_64','OpenBSD','OpenBSD_64','NetBSD','NetBSD_64','OS2Warp3','OS2Warp4','OS2Warp45','OS2eCS','OS21x','OS2','MacOS','MacOS_64','MacOS106','MacOS106_64','MacOS107_64','MacOS108_64','MacOS109_64','MacOS1010_64','MacOS1011_64','MacOS1012_64','MacOS1013_64','DOS','Netware','L4','QNX','JRockitVE','VBoxBS_64'))
 if ($global:guestostype.id) {
  $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute($global:guestostype.id)
 }
 $OsTypeIdCollection.Add($ValidateSetOsTypeId)
 $OsTypeId = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("OsTypeId", [string], $OsTypeIdCollection)
 $CustomAttributes.HelpMessage = 'Enter the paravirtual provider for the virtual machine'
 $ParavirtProvidersCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $ParavirtProvidersCollection.Add($CustomAttributes)
 $ValidateSetParavirtProviders = New-Object System.Management.Automation.ValidateSetAttribute(@('None','Default','Legacy','Minimal','HyperV','KVM'))
 if ($global:systempropertiessupported.ParavirtProviders) {
  $ValidateSetParavirtProviders = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.ParavirtProviders)
 }
 $ParavirtProvidersCollection.Add($ValidateSetParavirtProviders)
 $ParavirtProviders = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ParavirtProvider", [string], $ParavirtProvidersCollection)
 $CustomAttributes.HelpMessage = 'Enter the clipboard mode for the virtual machine'
 $ClipboardModesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $ClipboardModesCollection.Add($CustomAttributes)
 $ValidateSetClipboardModes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disabled','HostToGuest','GuestToHost','Bidirectional'))
 if ($global:systempropertiessupported.ClipboardModes) {
  $ValidateSetClipboardModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.ClipboardModes)
 }
 $ClipboardModesCollection.Add($ValidateSetClipboardModes)
 $ClipboardModes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ClipboardMode", [string], $ClipboardModesCollection)
 $CustomAttributes.HelpMessage = "Enter the drag n' drop mode for the virtual machine"
 $DndModesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $DndModesCollection.Add($CustomAttributes)
 $ValidateSetDndModes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disabled','HostToGuest','GuestToHost','Bidirectional'))
 if ($global:systempropertiessupported.DndModes) {
  $ValidateSetDndModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.DndModes)
 }
 $DndModesCollection.Add($ValidateSetDndModes)
 $DndModes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("DndMode", [string], $DndModesCollection)
 $CustomAttributes.HelpMessage = 'Enter the firmware type for the virtual machine'
 $FirmwareTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $FirmwareTypesCollection.Add($CustomAttributes)
 $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('BIOS','EFI','EFI32','EFI64','EFIDUAL'))
 if ($global:systempropertiessupported.FirmwareTypes) {
  $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.FirmwareTypes)
 }
 $FirmwareTypesCollection.Add($ValidateSetFirmwareTypes)
 $FirmwareTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("FirmwareType", [string], $FirmwareTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the pointing HID type for the virtual machine'
 $PointingHidTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $PointingHidTypesCollection.Add($CustomAttributes)
 $ValidateSetPointingHidTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('PS2Mouse','USBTablet','USBMultiTouch'))
 if ($global:systempropertiessupported.PointingHidTypes) {
  $ValidateSetPointingHidTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.PointingHidTypes)
 }
 $PointingHidTypesCollection.Add($ValidateSetPointingHidTypes)
 $PointingHidTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("PointingHidType", [string], $PointingHidTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the keyboard HID type for the virtual machine'
 $KeyboardHidTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $KeyboardHidTypesCollection.Add($CustomAttributes)
 $ValidateSetKeyboardHidTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('PS2Keyboard','USBKeyboard'))
 if ($global:systempropertiessupported.KeyboardHidTypes) {
  $ValidateSetKeyboardHidTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.KeyboardHidTypes)
 }
 $KeyboardHidTypesCollection.Add($ValidateSetKeyboardHidTypes)
 $KeyboardHidTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("KeyboardHidType", [string], $KeyboardHidTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the Virtual File System type for the virtual machine'
 $VfsTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $VfsTypesCollection.Add($CustomAttributes)
 $ValidateSetVfsTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('File','Cloud','S3'))
 if ($global:systempropertiessupported.VfsTypes) {
  $ValidateSetVfsTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.VfsTypes)
 }
 $VfsTypesCollection.Add($ValidateSetVfsTypes)
 $VfsTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("VfsType", [string], $VfsTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording audio codec for the virtual machine'
 $RecordingAudioCodecsCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingAudioCodecsCollection.Add($CustomAttributes)
 $ValidateSetRecordingAudioCodecs = New-Object System.Management.Automation.ValidateSetAttribute(@('Opus'))
 if ($global:systempropertiessupported.RecordingAudioCodecs) {
  $ValidateSetRecordingAudioCodecs = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingAudioCodecs)
 }
 $RecordingAudioCodecsCollection.Add($ValidateSetRecordingAudioCodecs)
 $RecordingAudioCodecs = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingAudioCodec", [string], $RecordingAudioCodecsCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording video codec for the virtual machine'
 $RecordingVideoCodecsCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingVideoCodecsCollection.Add($CustomAttributes)
 $ValidateSetRecordingVideoCodecs = New-Object System.Management.Automation.ValidateSetAttribute(@('VP8'))
 if ($global:systempropertiessupported.RecordingVideoCodecs) {
  $ValidateSetRecordingVideoCodecs = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingVideoCodecs)
 }
 $RecordingVideoCodecsCollection.Add($ValidateSetRecordingVideoCodecs)
 $RecordingVideoCodecs = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingVideoCodec", [string], $RecordingVideoCodecsCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording VS codec for the virtual machine'
 $RecordingVsMethodsCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingVsMethodsCollection.Add($CustomAttributes)
 $ValidateSetRecordingVsMethods = New-Object System.Management.Automation.ValidateSetAttribute(@('None'))
 if ($global:systempropertiessupported.RecordingVsMethods) {
  $ValidateSetRecordingVsMethods = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingVsMethods)
 }
 $RecordingVsMethodsCollection.Add($ValidateSetRecordingVsMethods)
 $RecordingVsMethods = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingVsMethod", [string], $RecordingVsMethodsCollection)
 $CustomAttributes.HelpMessage = 'Enter the recording VRC mode for the virtual machine'
 $RecordingVrcModesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $RecordingVrcModesCollection.Add($CustomAttributes)
 $ValidateSetRecordingVrcModes = New-Object System.Management.Automation.ValidateSetAttribute(@('CBR'))
 if ($global:systempropertiessupported.RecordingVrcModes) {
  $ValidateSetRecordingVrcModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.RecordingVrcModes)
 }
 $RecordingVrcModesCollection.Add($ValidateSetRecordingVrcModes)
 $RecordingVrcModes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("RecordingVrcMode", [string], $RecordingVrcModesCollection)
 $CustomAttributes.HelpMessage = 'Enter the graphics controller type for the virtual machine'
 $GraphicsControllerTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $GraphicsControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetGraphicsControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('VBoxVGA','VMSVGA','VBoxSVGA','Null'))
 if ($global:systempropertiessupported.GraphicsControllerTypes) {
  $ValidateSetGraphicsControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.GraphicsControllerTypes)
 }
 $GraphicsControllerTypesCollection.Add($ValidateSetGraphicsControllerTypes)
 $GraphicsControllerTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("GraphicsControllerType", [string], $GraphicsControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the auto stop type for the virtual machine'
 $AutostopTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $AutostopTypesCollection.Add($CustomAttributes)
 $ValidateSetAutostopTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disabled','SaveState','PowerOff','AcpiShutdown'))
 if ($global:systempropertiessupported.AutostopTypes) {
  $ValidateSetAutostopTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.AutostopTypes)
 }
 $AutostopTypesCollection.Add($ValidateSetAutostopTypes)
 $AutostopTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("AutostopType", [string], $AutostopTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the VM process priority for the virtual machine'
 $VmProcPrioritiesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $VmProcPrioritiesCollection.Add($CustomAttributes)
 $ValidateSetVmProcPriorities = New-Object System.Management.Automation.ValidateSetAttribute(@('Default','Flat','Low','Normal','High'))
 if ($global:systempropertiessupported.VmProcPriorities) {
  $ValidateSetVmProcPriorities = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.VmProcPriorities)
 }
 $VmProcPrioritiesCollection.Add($ValidateSetVmProcPriorities)
 $VmProcPriorities = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("VmProcPriority", [string], $VmProcPrioritiesCollection)
 $CustomAttributes.HelpMessage = 'Enter the network attachment type for the virtual machine'
 $NetworkAttachmentTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $NetworkAttachmentTypesCollection.Add($CustomAttributes)
 $ValidateSetNetworkAttachmentTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('NAT','Bridged','Internal','HostOnly','Generic','NATNetwork','Null'))
 if ($global:systempropertiessupported.NetworkAttachmentTypes) {
  $ValidateSetNetworkAttachmentTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.NetworkAttachmentTypes)
 }
 $NetworkAttachmentTypesCollection.Add($ValidateSetNetworkAttachmentTypes)
 $NetworkAttachmentTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("NetworkAttachmentType", [string], $NetworkAttachmentTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the network adapter type for the virtual machine'
 $NetworkAdapterTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $NetworkAdapterTypesCollection.Add($CustomAttributes)
 $ValidateSetNetworkAdapterTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('Am79C970A','Am79C973','I82540EM','I82543GC','I82545EM','Virtio','Am79C960'))
 if ($global:systempropertiessupported.NetworkAdapterTypes) {
  $ValidateSetNetworkAdapterTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.NetworkAdapterTypes)
 }
 $NetworkAdapterTypesCollection.Add($ValidateSetNetworkAdapterTypes)
 $NetworkAdapterTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("NetworkAdapterType", [string], $NetworkAdapterTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the port mode for the virtual machine'
 $PortModesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $PortModesCollection.Add($CustomAttributes)
 $ValidateSetPortModes = New-Object System.Management.Automation.ValidateSetAttribute(@('Disconnected','HostPipe','HostDevice','RawFile','TCP'))
 if ($global:systempropertiessupported.PortModes) {
  $ValidateSetPortModes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.PortModes)
 }
 $PortModesCollection.Add($ValidateSetPortModes)
 $PortModes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("PortMode", [string], $PortModesCollection)
 $CustomAttributes.HelpMessage = 'Enter the emulated UART implementation type for the virtual machine'
 $UartTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $UartTypesCollection.Add($CustomAttributes)
 $ValidateSetUartTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('U16450','U16550A','U16750'))
 if ($global:systempropertiessupported.UartTypes) {
  $ValidateSetUartTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.UartTypes)
 }
 $UartTypesCollection.Add($ValidateSetUartTypes)
 $UartTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("UartType", [string], $UartTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the USB controller type for the virtual machine'
 $UsbControllerTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $UsbControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetUsbControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('OHCI','EHCI','XHCI'))
 if ($global:systempropertiessupported.UsbControllerTypes) {
  $ValidateSetUsbControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.UsbControllerTypes)
 }
 $UsbControllerTypesCollection.Add($ValidateSetUsbControllerTypes)
 $UsbControllerTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("UsbControllerType", [string], $UsbControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the audio driver type for the virtual machine'
 $AudioDriverTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $AudioDriverTypesCollection.Add($CustomAttributes)
 $ValidateSetAudioDriverTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('DirectSound','Null'))
 if ($global:systempropertiessupported.AudioDriverTypes) {
  $ValidateSetAudioDriverTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.AudioDriverTypes)
 }
 $AudioDriverTypesCollection.Add($ValidateSetAudioDriverTypes)
 $AudioDriverTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("AudioDriverType", [string], $AudioDriverTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the audio controller type for the virtual machine'
 $AudioControllerTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $AudioControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetAudioControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('AC97','SB16','HDA'))
 if ($global:systempropertiessupported.AudioControllerTypes) {
  $ValidateSetAudioControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.AudioControllerTypes)
 }
 $AudioControllerTypesCollection.Add($ValidateSetAudioControllerTypes)
 $AudioControllerTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("AudioControllerType", [string], $AudioControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the storage bus for the virtual machine'
 $StorageBusesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $StorageBusesCollection.Add($CustomAttributes)
 $ValidateSetStorageBuses = New-Object System.Management.Automation.ValidateSetAttribute(@('SATA','IDE','SCSI','Floppy','SAS','USB','PCIe','VirtioSCSI'))
 if ($global:systempropertiessupported.StorageBuses) {
  $ValidateSetStorageBuses = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.StorageBuses)
 }
 $StorageBusesCollection.Add($ValidateSetStorageBuses)
 $StorageBuses = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("StorageBus", [string], $StorageBusesCollection)
 $CustomAttributes.HelpMessage = 'Enter the storage controller type for the virtual machine'
 $StorageControllerTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $StorageControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetStorageControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('IntelAhci','PIIX4','PIIX3','ICH6','LsiLogic','BusLogic','I82078','LsiLogicSas','USB','NVMe','VirtioSCSI'))
 if ($global:systempropertiessupported.StorageControllerTypes) {
  $ValidateSetStorageControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.StorageControllerTypes)
 }
 $StorageControllerTypesCollection.Add($ValidateSetStorageControllerTypes)
 $StorageControllerTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("StorageControllerType", [string], $StorageControllerTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the chipset type for the virtual machine'
 $ChipsetTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $ChipsetTypesCollection.Add($CustomAttributes)
 $ValidateSetChipsetTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('PIIX3','ICH9'))
 if ($global:systempropertiessupported.ChipsetTypes) {
  $ValidateSetChipsetTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.ChipsetTypes)
 }
 $ChipsetTypesCollection.Add($ValidateSetChipsetTypes)
 $ChipsetTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ChipsetType", [string], $ChipsetTypesCollection)
 $CustomAttributes.HelpMessage = 'Enter the number of CPUs available to the virtual machine'
 $CpuCountCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $CpuCountCollection.Add($CustomAttributes)
 $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute(1, 32)
 if ($global:systempropertiessupported.MinGuestCPUCount -and $global:systempropertiessupported.MaxGuestCPUCount) {
  $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestCPUCount, $global:systempropertiessupported.MaxGuestCPUCount)
 }
 $CpuCountCollection.Add($ValidateSetCpuCount)
 $CpuCount = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("CpuCount", [uint64], $CpuCountCollection)
 $CustomAttributes.HelpMessage = 'Enter the memory size in MB for the virtual machine'
 $MemorySizeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $MemorySizeCollection.Add($CustomAttributes)
 $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute(4, 2097152)
 if ($global:systempropertiessupported.MinGuestRam -and $global:systempropertiessupported.MaxGuestRam) {
  $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestRam, $global:systempropertiessupported.MaxGuestRam)
 }
 $MemorySizeCollection.Add($ValidateSetMemorySize)
 $MemorySize = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("MemorySize", [uint64], $MemorySizeCollection)
 $CustomAttributes.HelpMessage = 'Enter the memory balloon size in MB for the virtual machine'
 $MemoryBalloonSizeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $MemoryBalloonSizeCollection.Add($CustomAttributes)
 $ValidateSetMemoryBalloonSize = New-Object System.Management.Automation.ValidateRangeAttribute(4, 2097152)
 if ($global:systempropertiessupported.MinGuestRam -and $global:systempropertiessupported.MaxGuestRam) {
  $ValidateSetMemoryBalloonSize = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestRam, $global:systempropertiessupported.MaxGuestRam)
 }
 $MemoryBalloonSizeCollection.Add($ValidateSetMemoryBalloonSize)
 $MemoryBalloonSize = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("MemoryBalloonSize", [uint64], $MemoryBalloonSizeCollection)
 $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
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
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
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
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 if ($imachines) {
  foreach ($imachine in $imachines) {
   if ($imachine.State -ne 'PoweredOff') {Write-Host "[Error] Machine $($imachine.Name) is not powered off. Power it off and try again." -ForegroundColor Red -BackgroundColor Black;return}
   try {
    if ($ModuleHost.ToLower() -eq 'websrv') {
     Write-Verbose "Getting write lock on machine $($imachine.Name)"
     $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
     # create a new machine object
     $mmachine = New-Object VirtualBoxVM
     # get the mutable machine object
     Write-Verbose "Getting the mutable machine object"
     $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
     $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
     try {
      Write-Verbose "Modifying requeseted settings on machine $($mmachine.Name)"
      if ($Icon -or $Icon -eq '') {
       if ($Icon -eq '') {$global:vbox.IMachine_setIcon($mmachine.Id, $null)}
       else {
        # convert to png
        if (!(Test-Path "$env:TEMP\VirtualBoxPS")) {New-Item -ItemType Directory -Path "$env:TEMP\VirtualBoxPS\" -Force -Confirm:$false | Write-Verbose}
        Add-Type -AssemblyName system.drawing
        $imageFormat = "System.Drawing.Imaging.ImageFormat" -as [type]
        $image = [drawing.image]::FromFile($Icon)
        $image.Save("$env:TEMP\VirtualBoxPS\icon.png", $imageFormat::Png)
        $octet = [convert]::ToBase64String((Get-Content "$env:TEMP\VirtualBoxPS\icon.png" -Encoding Byte))
        $global:vbox.IMachine_setIcon($mmachine.Id, $octet)
        Remove-Item -Path "$env:TEMP\VirtualBoxPS\icon.png" -Confirm:$false -Force
       }
      }
      if ($Description) {$global:vbox.IMachine_setDescription($mmachine.Id, $Description)}
      if ($HardwareUuid) {$global:vbox.IMachine_setHardwareUUID($mmachine.Id, $HardwareUuid)}
      if ($CpuCount) {$global:vbox.IMachine_setCPUCount($mmachine.Id, $CpuCount)}
      if ($CpuHotPlugEnabled) {$global:vbox.IMachine_setCPUHotPlugEnabled($mmachine.Id, $CpuHotPlugEnabled)}
      if ($CpuExecutionCap) {$global:vbox.IMachine_setCPUExecutionCap($mmachine.Id, $CpuExecutionCap)}
      if ($CpuIdPortabilityLevel) {$global:vbox.IMachine_setCPUIDPortabilityLevel($mmachine.Id, $CpuIdPortabilityLevel)}
      if ($MemorySize) {$global:vbox.IMachine_setMemorySize($mmachine.Id, $MemorySize)}
      if ($MemoryBalloonSize) {$global:vbox.IMachine_setMemoryBalloonSize($mmachine.Id, $MemoryBalloonSize)}
      if ($PageFusionEnabled) {$global:vbox.IMachine_setPageFusionEnabled($mmachine.Id, $PageFusionEnabled)}
      if ($FirmwareType) {$global:vbox.IMachine_setFirmwareType($mmachine.Id, $FirmwareType)}
      if ($PointingHidType) {$global:vbox.IMachine_setPointingHIDType($mmachine.Id, $PointingHidType)}
      if ($KeyboardHidType) {$global:vbox.IMachine_setKeyboardHIDType($mmachine.Id, $KeyboardHidType)}
      if ($HpetEnabled) {$global:vbox.IMachine_setHPETEnabled($mmachine.Id, $HpetEnabled)}
      if ($ChipsetType) {$global:vbox.IMachine_setChipsetType($mmachine.Id, $ChipsetType)}
      if ($EmulatedUsbCardReaderEnabled) {$global:vbox.IMachine_setEmulatedUSBCardReaderEnabled($mmachine.Id, $EmulatedUsbCardReaderEnabled)}
      if ($ClipboardMode) {$global:vbox.IMachine_setClipboardMode($mmachine.Id, $ClipboardMode)}
      if ($ClipboardFileTransfersEnabled) {$global:vbox.IMachine_setClipboardFileTransfersEnabled($mmachine.Id, $ClipboardFileTransfersEnabled)}
      if ($DndMode) {$global:vbox.IMachine_setDnDMode($mmachine.Id, $DndMode)}
      if ($TeleporterEnabled) {$global:vbox.IMachine_setTeleporterEnabled($mmachine.Id, $TeleporterEnabled)}
      if ($TeleporterPort) {$global:vbox.IMachine_setTeleporterPort($mmachine.Id, $TeleporterPort)}
      if ($TeleporterAddress) {$global:vbox.IMachine_setTeleporterAddress($mmachine.Id, $TeleporterAddress)}
      if ($TeleporterPassword) {$global:vbox.IMachine_setTeleporterPassword($mmachine.Id, [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($TeleporterPassword)))}
      if ($ParavirtProvider) {$global:vbox.IMachine_setParavirtProvider($mmachine.Id, $ParavirtProvider)}
      if ($RtcUseUtc) {$global:vbox.IMachine_setRTCUseUTC($mmachine.Id, $RtcUseUtc)}
      if ($IoCacheEnabled) {$global:vbox.IMachine_setIOCacheEnabled($mmachine.Id, $IoCacheEnabled)}
      if ($IoCacheSize) {$global:vbox.IMachine_setIOCacheSize($mmachine.Id, $IoCacheSize)}
      if ($TracingEnabled) {$global:vbox.IMachine_setTracingEnabled($mmachine.Id, $TracingEnabled)}
      if ($TracingConfig) {$global:vbox.IMachine_setTracingConfig($mmachine.Id, $TracingConfig)}
      if ($AllowTracingToAccessVM) {$global:vbox.IMachine_setAllowTracingToAccessVM($mmachine.Id, $AllowTracingToAccessVM)}
      if ($AutostartEnabled) {$global:vbox.IMachine_setAutostartEnabled($mmachine.Id, $AutostartEnabled)}
      if ($AutostartDelay) {$global:vbox.IMachine_setAutostartDelay($mmachine.Id, $AutostartDelay)}
      if ($AutostopType) {$global:vbox.IMachine_setAutostopType($mmachine.Id, $AutostopType)}
      if ($CpuProfile) {$global:vbox.IMachine_setCPUProfile($mmachine.Id, $CpuProfile)}
     }
     catch {
      Write-Verbose 'Exception applying custom parameters to machine'
      Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
      Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
     }
     # save new settings
     Write-Verbose "Saving new settings"
     $global:vbox.IMachine_saveSettings($mmachine.Id)
     # unlock machine session
     Write-Verbose "Unlocking machine session"
     $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     Write-Verbose "Getting write lock on machine $($imachine.Name)"
     $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
     # create a new machine object
     $mmachine = New-Object VirtualBoxVM
     # get the mutable machine object
     Write-Verbose "Getting the mutable machine object"
     $mmachine.ComObject = $imachine.ISession.Session.Machine
     $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
     try {
      Write-Verbose "Modifying requeseted settings on machine $($mmachine.Name)"
      if ($Icon -or $Icon -eq '') {
       if ($Icon -eq '') {$mmachine.ComObject.Icon = [byte[]]@($null)}
       else {
        # convert to png
        if (!(Test-Path "$env:TEMP\VirtualBoxPS")) {New-Item -ItemType Directory -Path "$env:TEMP\VirtualBoxPS\" -Force -Confirm:$false | Write-Verbose}
        Add-Type -AssemblyName system.drawing
        $imageFormat = "System.Drawing.Imaging.ImageFormat" -as [type]
        $image = [drawing.image]::FromFile($Icon)
        $image.Save("$env:TEMP\VirtualBoxPS\icon.png", $imageFormat::Png)
        [byte[]]$bytes = Get-Content "$env:TEMP\VirtualBoxPS\icon.png" -Encoding Byte
        $mmachine.ComObject.Icon = $bytes
        Remove-Item -Path "$env:TEMP\VirtualBoxPS\icon.png" -Confirm:$false -Force
       }
      }
      if ($Description) {$mmachine.ComObject.Description = $Description}
      if ($HardwareUuid) {$mmachine.ComObject.HardwareUUID = $HardwareUuid}
      if ($CpuCount) {$mmachine.ComObject.CPUCount = $CpuCount}
      if ($CpuHotPlugEnabled) {$mmachine.ComObject.CPUHotPlugEnabled = [int]$CpuHotPlugEnabled}
      if ($CpuExecutionCap) {$mmachine.ComObject.CPUExecutionCap = $CpuExecutionCap}
      if ($CpuIdPortabilityLevel) {$mmachine.ComObject.CPUIDPortabilityLevel = $CpuIdPortabilityLevel}
      if ($MemorySize) {$mmachine.ComObject.MemorySize = $MemorySize}
      if ($MemoryBalloonSize) {$mmachine.ComObject.MemoryBalloonSize = $MemoryBalloonSize}
      if ($PageFusionEnabled) {$mmachine.ComObject.PageFusionEnabled = [int]$PageFusionEnabled}
      if ($FirmwareType) {$mmachine.ComObject.FirmwareType = [FirmwareType]::new().ToInt($FirmwareType)}
      if ($PointingHidType) {$mmachine.ComObject.PointingHIDType = [PointingHIDType]::new().ToInt($PointingHidType - 1)}
      if ($KeyboardHidType) {$mmachine.ComObject.KeyboardHIDType = [KeyboardHIDType]::new().ToInt($KeyboardHidType - 1)}
      if ($HpetEnabled) {$mmachine.ComObject.HPETEnabled = [int]$HpetEnabled}
      if ($ChipsetType) {$mmachine.ComObject.ChipsetType = [ChipsetType]::new().ToInt($ChipsetType)}
      if ($EmulatedUsbCardReaderEnabled) {$mmachine.ComObject.EmulatedUSBCardReaderEnabled = [int]$EmulatedUsbCardReaderEnabled}
      if ($ClipboardMode) {$mmachine.ComObject.ClipboardMode = [ClipboardMode]::new().ToInt($ClipboardMode)}
      if ($ClipboardFileTransfersEnabled) {$mmachine.ComObject.ClipboardFileTransfersEnabled = [int]$ClipboardFileTransfersEnabled}
      if ($DndMode) {$mmachine.ComObject.DnDMode = [DnDMode]::new().ToInt($DndMode)}
      if ($TeleporterEnabled) {$mmachine.ComObject.TeleporterEnabled = [int]$TeleporterEnabled}
      if ($TeleporterPort) {$mmachine.ComObject.TeleporterPort = $TeleporterPort}
      if ($TeleporterAddress) {$mmachine.ComObject.TeleporterAddress = $TeleporterAddress}
      if ($TeleporterPassword) {$mmachine.ComObject.TeleporterPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($TeleporterPassword))}
      if ($ParavirtProvider) {$mmachine.ComObject.ParavirtProvider = [ParavirtProvider]::new().ToInt($ParavirtProvider)}
      if ($RtcUseUtc) {$mmachine.ComObject.RTCUseUTC = [int]$RtcUseUtc}
      if ($IoCacheEnabled) {$mmachine.ComObject.IOCacheEnabled = [int]$IoCacheEnabled}
      if ($IoCacheSize) {$mmachine.ComObject.IOCacheSize = $IoCacheSize}
      if ($TracingEnabled) {$mmachine.ComObject.TracingEnabled = [int]$TracingEnabled}
      if ($TracingConfig) {$mmachine.ComObject.TracingConfig = $TracingConfig}
      if ($AllowTracingToAccessVM) {$mmachine.ComObject.AllowTracingToAccessVM = [int]$AllowTracingToAccessVM}
      if ($AutostartEnabled) {$mmachine.ComObject.AutostartEnabled = [int]$AutostartEnabled}
      if ($AutostartDelay) {$mmachine.ComObject.AutostartDelay = $AutostartDelay}
      if ($AutostopType) {$mmachine.ComObject.AutostopType = [AutostopType]::new().ToInt($AutostopType - 1)}
      if ($CpuProfile) {$mmachine.ComObject.CPUProfile = $CpuProfile}
     }
     catch {
      Write-Verbose 'Exception applying custom parameters to machine'
      Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
      Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
     }
     # save new settings
     Write-Verbose "Saving new settings"
     $mmachine.ComObject.SaveSettings()
     # unlock machine session
     Write-Verbose "Unlocking machine session"
     $imachine.ISession.Session.UnlockMachine()
    } # end elseif com
   } # Try
   catch {
    Write-Verbose 'Exception creating machine'
    Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
    Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
   } # Catch
   finally {
    # release mutable machine objects if they exist
    if ($mmachine) {
     if ($mmachine.ISession.Id) {
      # release mutable session object
      Write-Verbose "Releasing mutable session object"
      $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
     }
     if ($mmachine.ISession.Session) {
      if ($mmachine.ISession.Session.State -gt 1) {
       $mmachine.ISession.Session.UnlockMachine()
      } # end if $mmachine.ISession.Session locked
     } # end if $mmachine.ISession.Session
     if ($mmachine.Id) {
      # release mutable object
      Write-Verbose "Releasing mutable object"
      $global:vbox.IManagedObjectRef_release($mmachine.Id)
     }
    }
    # obligatory session unlock
    Write-Verbose 'Cleaning up machine sessions'
    if ($imachines) {
     foreach ($imachine in $imachines) {
      if ($imachine.ISession.Id) {
       if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
        Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
        $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
       } # end if session state not unlocked
      } # end if $imachine.ISession.Id
      if ($imachine.ISession.Session) {
       if ($imachine.ISession.Session.State -gt 1) {
        $imachine.ISession.Session.UnlockMachine()
       } # end if $imachine.ISession.Session locked
      } # end if $imachine.ISession.Session
      if ($imachine.IConsole) {
       # release the iconsole session
       Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
       $global:vbox.IManagedObjectRef_release($imachine.IConsole)
      } # end if $imachine.IConsole
      #$imachine.ISession.Id = $null
      $imachine.IConsole = $null
      if ($imachine.IPercent) {$imachine.IPercent = $null}
      $imachine.MSession = $null
      $imachine.MConsole = $null
      $imachine.MMachine = $null
     } # end foreach $imachine in $imachines
    } # end if $imachines
   } # Finally
  } # foreach $imachine in $imachines
 } # end if $imachines
 else {Write-Host "[Error] No matching virtual machines were found using specified parameters" -ForegroundColor Red -BackgroundColor Black;return}
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Edit-VirtualBoxVMFirmware {
<#
.SYNOPSIS
Create a virtual machine
.DESCRIPTION
Creates a new virtual machine. The name provided by the Name parameter must not exist in the VirtualBox inventory, or this command will fail. You can optionally supply custom values using a large number of parameters available to this command. There are too many to fully document in this help text, so tab completion has been added where it is possible. The values provided by tab completion are updated when Start-VirtualBoxSession is successfully run. To force the values to be updated again, use the -Force switch with Start-VirtualBoxSession.
.PARAMETER Name
The name of the virtual machine. This is a required parameter.
.PARAMETER OsTypeId
The type ID for the virtual machine guest OS. This is a required parameter.
.PARAMETER LogoFadeIn
Enable or disable the BIOS logo fade in animation for the virtual machine.
.PARAMETER LogoFadeOut
Enable or disable the BIOS logo fade out animation for the virtual machine.
.PARAMETER LogoDisplayTime
The time (in milliseconds) to display the firmware logo for the virtual machine.
.PARAMETER LogoImagePath
The full local path to the external firmware splash image for the virtual machine. Specify an empty string to display the default image.
.PARAMETER BootMenuMode
The boot menu mode for the virtual machine.
.PARAMETER AcpiEnabled
Enable or disable ACPI support for the virtual machine.
.PARAMETER IoAcpiEnabled
Enable or disable I/O APIC support for the virtual machine.
.PARAMETER ApicMode
The APIC mode for the virtual machine.
.PARAMETER TimeOffset
The time offset (in milliseconds) from the host for the virtual machine. This allows for guests running with a different system date/time than the host. It is equivalent to setting the system date/time in the BIOS except it is not an absolute value but a relative one. Guest Additions time synchronization honors this offset.
.PARAMETER PxeDebugEnabled
Enable or disable PXE debug logging for the virtual machine.
.PARAMETER SmBiosUuidLittleEndian
Enable or disable the virtual machine presenting the SMBIOS UUID in little endian to the guest as mandated by the SMBIOS spec chapter 7.2.1. Before VirtualBox version 6.1 it was always presented in big endian form and to retain the old behavior this flag was introduced so it can be changed. Virtual machines created with VirtualBox 6.1 will default to true for this flag.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Edit-VirtualBoxVMFirmware -Name "My New Win10 VM" -BootMenuMode Disabled
Disable the boot menu for the virtual machine named "My New Win10 VM"
.NOTES
NAME        :  Edit-VirtualBoxVMFirmware
VERSION     :  1.0
LAST UPDATED:  2/16/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
None
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for virtual machine OS Type ID
Other optional input parameters available. Use "Get-Help Edit-VirtualBoxVMFirmware -Full" for a complete list.
.OUTPUTS
None
#>
[CmdletBinding()]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)"
,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Enable or disable the logo fade in for the virtual machine",
Mandatory=$false)]
  [bool]$LogoFadeIn,
[Parameter(HelpMessage="Enable or disable the logo fade out for the virtual machine",
Mandatory=$false)]
  [bool]$LogoFadeOut,
[Parameter(HelpMessage="Enter the logo display time (in ms) for the virtual machine",
Mandatory=$false)]
  [uint32]$LogoDisplayTime,
[Parameter(HelpMessage="Enter the full path to the firmware logo for the virtual machine",
ParameterSetName='Custom',Mandatory=$false)]
[ValidateNotNull()]
[ValidateScript({if ($_ -eq '') {return $true} elseif (Test-Path $_) {return $true}})]
  [string]$LogoImagePath,
[Parameter(HelpMessage="Enter the boot menu mode for the virtual machine",
Mandatory=$false)]
[ValidateSet('Disabled', 'MenuOnly', 'MessageAndMenu')]
  [string]$BootMenuMode,
[Parameter(HelpMessage="Enable or disable ACPI support for the virtual machine",
Mandatory=$false)]
  [bool]$AcpiEnabled,
[Parameter(HelpMessage="Enable or disable I/O APIC support for the virtual machine",
Mandatory=$false)]
  [bool]$IoApicEnabled,
[Parameter(HelpMessage="Enter the APIC mode for the virtual machine",
Mandatory=$false)]
[ValidateSet('Disabled', 'APIC', 'X2APIC')]
  [string]$ApicMode,
[Parameter(HelpMessage="Enter the time offset (in ms) from the host for the virtual machine",
Mandatory=$false)]
  [int64]$TimeOffset,
[Parameter(HelpMessage="Enable or disable PXE debug logging for the virtual machine",
Mandatory=$false)]
  [bool]$PxeDebugEnabled,
[Parameter(HelpMessage="Enable or disable encoding the SMBIOS UUID in little endian for the virtual machine",
Mandatory=$false)]
  [bool]$SmBiosUuidLittleEndian,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 if ($imachines) {
  foreach ($imachine in $imachines) {
   try {
    if ($ModuleHost.ToLower() -eq 'websrv') {
     Write-Verbose "Getting write lock on machine $($imachine.Name)"
     $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
     # create a new machine object
     $mmachine = New-Object VirtualBoxVM
     # get the mutable machine object
     Write-Verbose "Getting the mutable machine object"
     $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
     $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
     $mmachine.IBiosSettings.Id = $global:vbox.IMachine_BIOSSettings($mmachine.Id)
     try {
      Write-Verbose "Modifying requeseted settings on machine $($mmachine.Name)"
      if ($LogoFadeIn -ne $null) {$global:vbox.IBIOSSettings_setLogoFadeIn($mmachine.IBiosSettings.Id, $LogoFadeIn)}
      if ($LogoFadeOut -ne $null) {$global:vbox.IBIOSSettings_setLogoFadeOut($mmachine.IBiosSettings.Id, $LogoFadeOut)}
      if ($LogoDisplayTime) {$global:vbox.IBIOSSettings_setLogoDisplayTime($mmachine.IBiosSettings.Id, $LogoDisplayTime)}
      if ($LogoImagePath) {$global:vbox.IBIOSSettings_setLogoImagePath($mmachine.IBiosSettings.Id, $LogoImagePath)}
      if ($BootMenuMode) {$global:vbox.IBIOSSettings_setBootMenuMode($mmachine.IBiosSettings.Id, $BootMenuMode)}
      if ($AcpiEnabled -ne $null) {$global:vbox.IBIOSSettings_setACPIEnabled($mmachine.IBiosSettings.Id, $AcpiEnabled)}
      if ($IoApicEnabled -ne $null) {$global:vbox.IBIOSSettings_setIOAPICEnabled($mmachine.IBiosSettings.Id, $IoApicEnabled)}
      if ($ApicMode) {$global:vbox.IBIOSSettings_setAPICMode($mmachine.IBiosSettings.Id, $ApicMode)}
      if ($TimeOffset) {$global:vbox.IBIOSSettings_setTimeOffset($mmachine.IBiosSettings.Id, $TimeOffset)}
      if ($PxeDebugEnabled -ne $null) {$global:vbox.IBIOSSettings_setPXEDebugEnabled($mmachine.IBiosSettings.Id, $PxeDebugEnabled)}
      if ($SmBiosUuidLittleEndian -ne $null) {$global:vbox.IBIOSSettings_setSMBIOSUuidLittleEndian($mmachine.IBiosSettings.Id, $SmBiosUuidLittleEndian)}
     }
     catch {
      Write-Verbose 'Exception applying custom parameters to machine'
      Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
      Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
     }
     # save new settings
     Write-Verbose "Saving new settings"
     $global:vbox.IMachine_saveSettings($mmachine.Id)
     # unlock machine session
     Write-Verbose "Unlocking machine session"
     $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     Write-Verbose "Getting write lock on machine $($imachine.Name)"
     $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
     # create a new machine object
     $mmachine = New-Object VirtualBoxVM
     # get the mutable machine object
     Write-Verbose "Getting the mutable machine object"
     $mmachine.ComObject = $imachine.ISession.Session.Machine
     $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
     try {
      Write-Verbose "Modifying requeseted settings on machine $($mmachine.Name)"
      if ($LogoFadeIn -ne $null) {$mmachine.ComObject.BIOSSettings.LogoFadeIn = $LogoFadeIn}
      if ($LogoFadeOut -ne $null) {$mmachine.ComObject.BIOSSettings.LogoFadeOut = $LogoFadeOut}
      if ($LogoDisplayTime) {$mmachine.ComObject.BIOSSettings.LogoDisplayTime = $LogoDisplayTime}
      if ($LogoImagePath) {$mmachine.ComObject.BIOSSettings.LogoImagePath = $LogoImagePath}
      if ($BootMenuMode) {$mmachine.ComObject.BIOSSettings.BootMenuMode = $BootMenuMode}
      if ($AcpiEnabled -ne $null) {$mmachine.ComObject.BIOSSettings.ACPIEnabled = $AcpiEnabled}
      if ($IoApicEnabled -ne $null) {$mmachine.ComObject.BIOSSettings.IOAPICEnabled = $IoApicEnabled}
      if ($ApicMode) {$mmachine.ComObject.BIOSSettings.APICMode = $ApicMode}
      if ($TimeOffset) {$mmachine.ComObject.BIOSSettings.TimeOffset = $TimeOffset}
      if ($PxeDebugEnabled -ne $null) {$mmachine.ComObject.BIOSSettings.PXEDebugEnabled = $PxeDebugEnabled}
      if ($SmBiosUuidLittleEndian -ne $null) {$mmachine.ComObject.BIOSSettings.SMBIOSUuidLittleEndian = $SmBiosUuidLittleEndian}
     }
     catch {
      Write-Verbose 'Exception applying custom parameters to machine'
      Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
      Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
     }
     # save new settings
     Write-Verbose "Saving new settings"
     $mmachine.ComObject.SaveSettings()
     # unlock machine session
     Write-Verbose "Unlocking machine session"
     $imachine.ISession.Session.UnlockMachine()
    } # end elseif com
   } # Try
   catch {
    Write-Verbose 'Exception creating machine'
    Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
    Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
   } # Catch
   finally {
    # release mutable machine objects if they exist
    if ($mmachine) {
     if ($mmachine.ISession.Id) {
      # release mutable session object
      Write-Verbose "Releasing mutable session object"
      $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
     }
     if ($mmachine.ISession.Session) {
      if ($mmachine.ISession.Session.State -gt 1) {
       $mmachine.ISession.Session.UnlockMachine()
      } # end if $mmachine.ISession.Session locked
     } # end if $mmachine.ISession.Session
     if ($mmachine.Id) {
      # release mutable object
      Write-Verbose "Releasing mutable object"
      $global:vbox.IManagedObjectRef_release($mmachine.Id)
     }
    }
    # obligatory session unlock
    Write-Verbose 'Cleaning up machine sessions'
    if ($imachines) {
     foreach ($imachine in $imachines) {
      if ($imachine.ISession.Id) {
       if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
        Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
        $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
       } # end if session state not unlocked
      } # end if $imachine.ISession.Id
      if ($imachine.ISession.Session) {
       if ($imachine.ISession.Session.State -gt 1) {
        $imachine.ISession.Session.UnlockMachine()
       } # end if $imachine.ISession.Session locked
      } # end if $imachine.ISession.Session
      if ($imachine.IConsole) {
       # release the iconsole session
       Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
       $global:vbox.IManagedObjectRef_release($imachine.IConsole)
      } # end if $imachine.IConsole
      #$imachine.ISession.Id = $null
      $imachine.IConsole = $null
      if ($imachine.IPercent) {$imachine.IPercent = $null}
      $imachine.MSession = $null
      $imachine.MConsole = $null
      $imachine.MMachine = $null
     } # end foreach $imachine in $imachines
    } # end if $imachines
   } # Finally
  } # foreach $imachine in $imachines
 } # end if $imachines
 else {Write-Host "[Error] No matching virtual machines were found using specified parameters" -ForegroundColor Red -BackgroundColor Black;return}
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Get-VirtualBoxVMStorageController {
<#
.SYNOPSIS
Get VirtualBox storage controller information
.DESCRIPTION
Retrieves storage controllers by name, bus, type or all for a machine object, name, or GUID.
.PARAMETER Name
At least one storage controller name.
.PARAMETER Bus
At least one storage controller bus.
.PARAMETER ControllerType
At least one storage controller type.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER MachineName
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER MachineGuid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Get-VirtualBoxVMStorageController -Name sata -MachineName 2016

Name           : SATA
Bus            : SATA
ControllerType : IntelAhci
Bootable       : True

Get storage controller for the 2016 Core machine by name and machine name
.EXAMPLE
PS C:\> Get-VirtualBoxVM -Name 2016 | Get-VirtualBoxVMStorageController

Name           : SATA
Bus            : SATA
ControllerType : IntelAhci
Bootable       : True

Get all storage controllers by machine object from pipeline input
.EXAMPLE
PS C:\> Get-VirtualBoxVMStorageController -MachineName 2016

Name           : SATA
Bus            : SATA
ControllerType : IntelAhci
Bootable       : True

Get all storage controllers for the 2016 Core machine by machine name
.NOTES
NAME        :  Get-VirtualBoxVMStorageController
VERSION     :  1.0
LAST UPDATED:  2/8/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
New-VirtualBoxDisk
.INPUTS
String[]      :  Strings for storage controller names
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
.OUTPUTS
VirtualBoxVHD[]
#>
[cmdletbinding()]
Param(
[Parameter(HelpMessage="Enter one or more storage controller name(s)",
Mandatory=$false,Position=0)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(HelpMessage="Enter one or more storage controller bus(es)",
Mandatory=$false)]
[ValidateSet('IDE','SATA','SCSI','Floppy','SAS','USB','PCIe','VirtioSCSI')]
  [string[]]$Bus,
[Parameter(HelpMessage="Enter one or more storage controller type(s)",
Mandatory=$false)]
[ValidateSet('LsiLogic','BusLogic','IntelAhci','PIIX3','PIIX4','ICH6','I82078','LsiLogicSas','USB','NVMe','VirtioSCSI')]
  [string[]]$ControllerType,
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
Mandatory=$false,Position=1)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
Mandatory=$false)]
  [string[]]$MachineName,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
Mandatory=$false)]
  [guid[]]$MachineGuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Bus: `"$Bus`""
 Write-Verbose "Pipeline - ControllerType: `"$ControllerType`""
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - MachineName: `"$MachineName`""
 Write-Verbose "Pipeline - MachineGuid: `"$MachineGuid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $MachineName -or $MachineGuid)) {Write-Host "[Error] You must supply at least one virtual machine object, name, or GUID using supplied parameters.";return}
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($MachineName) {
  foreach ($item in $MachineName) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($MachineGuid) {
  foreach ($item in $MachineGuid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 $storagecontrollers = @()
 $obj = @()
 if ($imachines) {
  foreach ($imachine in $imachines) {
   try {
    # get storage controller inventory
    Write-Verbose "Getting storage controller inventory for the $($imachine.Name) machine"
    if ($ModuleHost.ToLower() -eq 'websrv') {
     foreach ($istoragecontrollerid in $global:vbox.IMachine_getStorageControllers($imachine.Id)) {
      Write-Verbose "Getting storage controller: $($istoragecontrollerid)"
      $storagecontroller = New-Object IStorageControllers
      $storagecontroller.Name = $global:vbox.IStorageController_getName($istoragecontrollerid)
      $storagecontroller.MaxDevicesPerPortCount = $global:vbox.IStorageController_getMaxDevicesPerPortCount($istoragecontrollerid)
      $storagecontroller.MinPortCount = $global:vbox.IStorageController_getMinPortCount($istoragecontrollerid)
      $storagecontroller.MaxPortCount = $global:vbox.IStorageController_getMaxPortCount($istoragecontrollerid)
      $storagecontroller.Instance = $global:vbox.IStorageController_getInstance($istoragecontrollerid)
      $storagecontroller.PortCount = $global:vbox.IStorageController_getPortCount($istoragecontrollerid)
      $storagecontroller.Bus = $global:vbox.IStorageController_getBus($istoragecontrollerid)
      $storagecontroller.ControllerType = $global:vbox.IStorageController_getControllerType($istoragecontrollerid)
      $storagecontroller.UseHostIOCache = $global:vbox.IStorageController_getUseHostIOCache($istoragecontrollerid)
      $storagecontroller.Bootable = $global:vbox.IStorageController_getBootable($istoragecontrollerid)
      $storagecontroller.Id = $istoragecontrollerid
      [IStorageControllers[]]$storagecontrollers += [IStorageControllers]@{Name=$storagecontroller.Name;MaxDevicesPerPortCount=$storagecontroller.MaxDevicesPerPortCount;MinPortCount=$storagecontroller.MinPortCount;MaxPortCount=$storagecontroller.MaxPortCount;Instance=$storagecontroller.Instance;PortCount=$storagecontroller.PortCount;Bus=$storagecontroller.Bus;ControllerType=$storagecontroller.ControllerType;UseHostIOCache=$storagecontroller.UseHostIOCache;Bootable=$storagecontroller.Bootable;Id=$storagecontroller.Id}
     } # end foreach loop inventory
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     foreach ($istoragecontroller in $imachine.ComObject.StorageControllers) {
      Write-Verbose "Getting storage controller: $($istoragecontroller.Name)"
      $storagecontroller = New-Object IStorageControllers
      $storagecontroller.Name = $istoragecontroller.Name
      $storagecontroller.MaxDevicesPerPortCount = $istoragecontroller.MaxDevicesPerPortCount
      $storagecontroller.MinPortCount = $istoragecontroller.MinPortCount
      $storagecontroller.MaxPortCount = $istoragecontroller.MaxPortCount
      $storagecontroller.Instance = $istoragecontroller.Instance
      $storagecontroller.PortCount = $istoragecontroller.PortCount
      $storagecontroller.Bus = [StorageBus]::new().ToStr($istoragecontroller.Bus)
      $storagecontroller.ControllerType = [StorageControllerType]::new().ToStr($istoragecontroller.ControllerType)
      $storagecontroller.UseHostIOCache = $istoragecontroller.UseHostIOCache
      $storagecontroller.Bootable = $istoragecontroller.Bootable
      $storagecontroller.ComObject = $istoragecontroller
      [IStorageControllers[]]$storagecontrollers += [IStorageControllers]@{Name=$storagecontroller.Name;MaxDevicesPerPortCount=$storagecontroller.MaxDevicesPerPortCount;MinPortCount=$storagecontroller.MinPortCount;MaxPortCount=$storagecontroller.MaxPortCount;Instance=$storagecontroller.Instance;PortCount=$storagecontroller.PortCount;Bus=$storagecontroller.Bus;ControllerType=$storagecontroller.ControllerType;UseHostIOCache=$storagecontroller.UseHostIOCache;Bootable=$storagecontroller.Bootable;ComObject=$storagecontroller.ComObject}
     } # end foreach loop inventory
    } # end elseif com
   } # Try
   catch {
    Write-Verbose 'Exception retrieving storage controller information'
    Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
    Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
   } # Catch
   # filter by storage controller name
   if ($Name) {
    foreach ($storagecontroller in $storagecontrollers) {
     $matched = $false
     Write-Verbose "Matching $($storagecontroller.Name) to $Name"
     if ($storagecontroller.Name -match $Name) {Write-Verbose "Matched $($storagecontroller.Name) to $Name";$matched = $true}
     if ($matched -eq $true) {[IStorageControllers[]]$obj += [IStorageControllers]@{Name=$storagecontroller.Name;MaxDevicesPerPortCount=$storagecontroller.MaxDevicesPerPortCount;MinPortCount=$storagecontroller.MinPortCount;MaxPortCount=$storagecontroller.MaxPortCount;Instance=$storagecontroller.Instance;PortCount=$storagecontroller.PortCount;Bus=$storagecontroller.Bus;ControllerType=$storagecontroller.ControllerType;UseHostIOCache=$storagecontroller.UseHostIOCache;Bootable=$storagecontroller.Bootable;Id=$storagecontroller.Id;ComObject=$storagecontroller.ComObject}}
    } # foreach $storagecontroller in $storagecontrollers
   } # end if $Name
   # filter by storage controller bus
   elseif ($Bus) {
    foreach ($storagecontroller in $storagecontrollers) {
     $matched = $false
     Write-Verbose "Matching $($storagecontroller.Bus) to $Bus"
     if ($storagecontroller.Bus -match $Bus) {Write-Verbose "Matched $($storagecontroller.Bus) to $Bus";$matched = $true}
     if ($matched -eq $true) {[IStorageControllers[]]$obj += [IStorageControllers]@{Name=$storagecontroller.Name;MaxDevicesPerPortCount=$storagecontroller.MaxDevicesPerPortCount;MinPortCount=$storagecontroller.MinPortCount;MaxPortCount=$storagecontroller.MaxPortCount;Instance=$storagecontroller.Instance;PortCount=$storagecontroller.PortCount;Bus=$storagecontroller.Bus;ControllerType=$storagecontroller.ControllerType;UseHostIOCache=$storagecontroller.UseHostIOCache;Bootable=$storagecontroller.Bootable;Id=$storagecontroller.Id;ComObject=$storagecontroller.ComObject}}
    } # foreach $storagecontroller in $storagecontrollers
   } # end elseif $Bus
   # filter by storage controller type
   elseif ($ControllerType) {
    foreach ($storagecontroller in $storagecontrollers) {
     $matched = $false
     Write-Verbose "Matching $($storagecontroller.ControllerType) to $ControllerType"
     if ($storagecontroller.ControllerType -match $ControllerType) {Write-Verbose "Matched $($storagecontroller.ControllerType) to $ControllerType";$matched = $true}
     if ($matched -eq $true) {[IStorageControllers[]]$obj += [IStorageControllers]@{Name=$storagecontroller.Name;MaxDevicesPerPortCount=$storagecontroller.MaxDevicesPerPortCount;MinPortCount=$storagecontroller.MinPortCount;MaxPortCount=$storagecontroller.MaxPortCount;Instance=$storagecontroller.Instance;PortCount=$storagecontroller.PortCount;Bus=$storagecontroller.Bus;ControllerType=$storagecontroller.ControllerType;UseHostIOCache=$storagecontroller.UseHostIOCache;Bootable=$storagecontroller.Bootable;Id=$storagecontroller.Id;ComObject=$storagecontroller.ComObject}}
    } # foreach $storagecontroller in $storagecontrollers
   } # end elseif $ControllerType
   # no filter
   elseif (!($Name -or $Bus -or $ControllerType)) {foreach ($storagecontroller in $storagecontrollers) {[IStorageControllers[]]$obj += [IStorageControllers]@{Name=$storagecontroller.Name;MaxDevicesPerPortCount=$storagecontroller.MaxDevicesPerPortCount;MinPortCount=$storagecontroller.MinPortCount;MaxPortCount=$storagecontroller.MaxPortCount;Instance=$storagecontroller.Instance;PortCount=$storagecontroller.PortCount;Bus=$storagecontroller.Bus;ControllerType=$storagecontroller.ControllerType;UseHostIOCache=$storagecontroller.UseHostIOCache;Bootable=$storagecontroller.Bootable;Id=$storagecontroller.Id;ComObject=$storagecontroller.ComObject}}}
   else {Write-Verbose "[Warning] No storage controllers found"}
   Write-Verbose "Found $(($obj | Measure-Object).count) storage controller(s)"
   if ($obj) {
    # write virtual machines object to the pipeline as an array
    Write-Output ([System.Array]$obj)
   } # end if $obj
   else {
    Write-Verbose "[Warning] No matching storage controllers found."
   } # end else
  } # foreach $imachine in $imachines
 } # end if $imachines
 else {Write-Host "[Error] No matching virtual machines found using specified parameters." -ForegroundColor Red -BackgroundColor Black}
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function New-VirtualBoxVMStorageController {
<#
.SYNOPSIS
Create VirtualBox disk
.DESCRIPTION
Creates VirtualBox disks. The command will fail if a virtual disk with the same name exists in the VirtualBox inventory.
.PARAMETER Name
The virtual disk name.
.PARAMETER Format
The virtual disk format.
.PARAMETER Location
The location to store the virtual disk. If the path does not exist it will be created.
.PARAMETER AccessMode
Either Readonly or ReadWrite.
.PARAMETER LogicalSize
The size of the virtual disk in bytes.
.PARAMETER VariantType
The variant type of the virtual disk.
.PARAMETER VariantFlag
The variant flag of the virtual disk.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> New-VirtualBoxVMStorageController -AccessMode ReadWrite -Format VMDK -Location C:\Disks -LogicalSize 4194304 -Name TestDisk -VariantFlag Fixed -VariantType Standard -ProgressBar

Create a standard, fixed 4MB virtual disk named "TestDisk.vmdk" in the C:\Disks\ location and display a progress bar
.NOTES
NAME        :  New-VirtualBoxVMStorageController
VERSION     :  1.0
LAST UPDATED:  1/16/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisk
.INPUTS
String        :  String for virtual disk name
String        :  String for virtual disk format
String        :  String for virtual disk location
String        :  String for virtual disk access mode
UInt64        :  UInt64 for virtual disk size
String        :  String for virtual disk variant type
String        :  String for virtual disk variant flag
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(HelpMessage="Enter the storage controller name",
Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(Mandatory=$false,ValueFromPipeline=$true,
HelpMessage="Enter a virtual machine object to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter the name of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [string]$MachineName,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter the GUID of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [guid]$MachineGuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
DynamicParam {
 $CustomAttributes = New-Object System.Management.Automation.ParameterAttribute
 $CustomAttributes.Mandatory = $true
 $CustomAttributes.HelpMessage = 'Enter the storage bus for the virtual machine'
 $StorageBusesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $StorageBusesCollection.Add($CustomAttributes)
 $ValidateSetStorageBuses = New-Object System.Management.Automation.ValidateSetAttribute(@('SATA','IDE','SCSI','Floppy','SAS','USB','PCIe','VirtioSCSI'))
 if ($global:systempropertiessupported.StorageBuses) {
  $ValidateSetStorageBuses = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.StorageBuses)
 }
 $StorageBusesCollection.Add($ValidateSetStorageBuses)
 $StorageBuses = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("StorageBus", [string], $StorageBusesCollection)
 $CustomAttributes.HelpMessage = 'Enter the storage controller type for the virtual machine'
 $StorageControllerTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $StorageControllerTypesCollection.Add($CustomAttributes)
 $ValidateSetStorageControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('IntelAhci','PIIX4','PIIX3','ICH6','LsiLogic','BusLogic','I82078','LsiLogicSas','USB','NVMe','VirtioSCSI'))
 if ($global:systempropertiessupported.StorageControllerTypes) {
  $ValidateSetStorageControllerTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.StorageControllerTypes)
 }
 $StorageControllerTypesCollection.Add($ValidateSetStorageControllerTypes)
 $StorageControllerTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("StorageControllerType", [string], $StorageControllerTypesCollection)
 $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
 $paramDictionary.Add("StorageBus", $StorageBuses)
 $paramDictionary.Add("StorageControllerType", $StorageControllerTypes)
 return $paramDictionary
}
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 $StorageBus = $PSBoundParameters['StorageBus']
 $StorageControllerType = $PSBoundParameters['StorageControllerType']
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - MachineName: `"$MachineName`""
 Write-Verbose "Pipeline - MachineGuid: `"$MachineGuid`""
 if (!($Machine -or $MachineName -or $MachineGuid)) {Write-Host "[Error] You must supply at least one machine object, name, or GUID using provided parameters." -ForegroundColor Red -BackgroundColor Black;return}
 if ($Machine) {
  $imachines = $Machine
 } # get vm inventory (by $Machine)
 elseif ($MachineName) {
  $imachines = Get-VirtualBoxVM -Name $MachineName -SkipCheck
 } # get vm inventory (by $MachineName)
 elseif ($MachineGuid) {
  $imachines = Get-VirtualBoxVM -Guid $MachineGuid -SkipCheck
 } # get vm inventory (by $MachineGuid)
 if ($imachines) {
  foreach ($imachine in $imachines) {
   if ($imachine.State -match 'PoweredOff') {
    try {
     if ($ModuleHost.ToLower() -eq 'websrv') {
      Write-Verbose "Getting write lock on machine $($imachine.Name)"
      $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
      $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
      # create the storage controller on the current machine
      Write-Verbose "Creating the $StorageBus storage controller $Name for the $($imachine.Name) machine"
      $istoragecontroller = $global:vbox.IMachine_addStorageController($mmachine.Id, $Name, [StorageBus]::new().ToInt($StorageBus))
      # set the storage controller type
      Write-Verbose "Setting storage controller type to $StorageControllerType"
      $global:vbox.IStorageController_setControllerType($istoragecontroller, [StorageControllerType]::new().ToInt($StorageControllerType))
      # save new settings
      Write-Verbose "Saving new settings"
      $global:vbox.IMachine_saveSettings($mmachine.Id)
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if websrv
     elseif ($ModuleHost.ToLower() -eq 'com') {
      Write-Verbose "Getting write lock on machine $($imachine.Name)"
      $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.ComObject = $imachine.ISession.Session.Machine
      $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
      # create the storage controller on the current machine
      Write-Verbose "Creating the $StorageBus storage controller $Name for the $($imachine.Name) machine"
      $istoragecontroller = $mmachine.ComObject.AddStorageController($Name, [StorageBus]::new().ToInt($StorageBus))
      # set the storage controller type
      Write-Verbose "Setting storage controller type to $StorageControllerType"
      $istoragecontroller.ControllerType = [StorageControllerType]::new().ToInt($StorageControllerType)
      # save new settings
      Write-Verbose "Saving new settings"
      $mmachine.ComObject.SaveSettings()
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $imachine.ISession.Session.UnlockMachine()
     } # end elseif com
    } # Try
    catch {
     Write-Verbose 'Exception creating storage controller'
     Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
     Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
    } # Catch
    finally {
     # release mutable machine objects if they exist
     if ($mmachine) {
      if ($mmachine.ISession.Id) {
       # release mutable session object
       Write-Verbose "Releasing mutable session object"
       $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
      }
      if ($mmachine.ISession.Session) {
       if ($mmachine.ISession.Session.State -gt 1) {
        $mmachine.ISession.Session.UnlockMachine()
       } # end if $mmachine.ISession.Session locked
      } # end if $mmachine.ISession.Session
      if ($mmachine.Id) {
       # release mutable object
       Write-Verbose "Releasing mutable object"
       $global:vbox.IManagedObjectRef_release($mmachine.Id)
      }
     }
     # obligatory session unlock
     Write-Verbose 'Cleaning up machine sessions'
     if ($imachines) {
      foreach ($imachine in $imachines) {
       if ($imachine.ISession.Id) {
        if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
         Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
         $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
        } # end if session state not unlocked
       } # end if $imachine.ISession.Id
       if ($imachine.ISession.Session) {
        if ($imachine.ISession.Session.State -gt 1) {
         $imachine.ISession.Session.UnlockMachine()
        } # end if $imachine.ISession.Session locked
       } # end if $imachine.ISession.Session
       if ($imachine.IConsole) {
        # release the iconsole session
        Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
        $global:vbox.IManagedObjectRef_release($imachine.IConsole)
       } # end if $imachine.IConsole
       #$imachine.ISession.Id = $null
       $imachine.IConsole = $null
       if ($imachine.IPercent) {$imachine.IPercent = $null}
       $imachine.MSession = $null
       $imachine.MConsole = $null
       $imachine.MMachine = $null
      } # end foreach $imachine in $imachines
     } # end if $imachines
    } # Finally
   } # end if machine powered off
   else {Write-Host "[Error] The requested $($imachine.Name) machine is not powered off. Power the machine off and try again." -ForegroundColor Red -BackgroundColor Black}
  } # foreach $imachine in $imachines
 } # end if $imachines
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function New-VirtualBoxVMOpticalDrive {
<#
.SYNOPSIS
Mount VirtualBox disc
.DESCRIPTION
Mounts VirtualBox discs. The command will fail if a virtual disc is already mounted to the specified virtual machine.
.PARAMETER Disc
At least one virtual disc object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER MachineName
The name of the virtual machine to mount the disc to. This is a required parameter.
.PARAMETER Controller
The name of the storage controller to mount the disc to. This is a required parameter.
.PARAMETER ControllerPort
The port of the storage controller to mount the disc to. This is a required parameter.
.PARAMETER ControllerSlot
The slot of the storage controller to mount the disc to. This is a required parameter.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> New-VirtualBoxVMOpticalDrive -Name TestDisc -MachineName Win10 -Controller SATA -ControllerPort 0 -ControllerSlot 0

Mounts the virtual disc named "TestDisc.vmdk" to the Win10 virtual machine SATA controller on port 0 slot 0
.NOTES
NAME        :  New-VirtualBoxVMOpticalDrive
VERSION     :  1.0
LAST UPDATED:  1/20/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisc
.INPUTS
VirtualBoxVHD[]:  VirtualBoxVHDs for virtual disc objects
String[]       :  Strings for virtual disc names
GUID[]         :  GUIDS for virtual disc GUIDS
String         :  String for virtual machine name
String         :  String for controller name
Int            :  Integer for controller port
Int            :  Integer for controller slot
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter a virtual disc object",
Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVHD]$Disc,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter a virtual disc name")]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter a virtual disc GUID")]
[ValidateNotNullorEmpty()]
  [guid]$Guid,
[Parameter(Mandatory=$false,HelpMessage="Enter a virtual machine object to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(Mandatory=$false,HelpMessage="Enter the name of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [string]$MachineName,
[Parameter(Mandatory=$false,HelpMessage="Enter the GUID of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [guid]$MachineGuid,
[Parameter(Mandatory=$true,HelpMessage="Enter the name of the controller to mount the disc to",
ParameterSetName="HardDisc")]
[ValidateNotNullorEmpty()]
  [string]$Controller,
[Parameter(Mandatory=$true,HelpMessage="Enter the port number to mount the disc to",
ParameterSetName="HardDisc")]
[ValidateNotNullorEmpty()]
  [int]$ControllerPort,
[Parameter(Mandatory=$true,HelpMessage="Enter the slot number to mount the disc to",
ParameterSetName="HardDisc")]
[ValidateNotNullorEmpty()]
  [int]$ControllerSlot,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 Write-Verbose "Pipeline - Disc: `"$Disc`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 Write-Verbose "Machine: `"$Machine`""
 Write-Verbose "Machine Name: `"$MachineName`""
 Write-Verbose "Machine Guid: `"$MachineGuid`""
 Write-Verbose "Controller Name: `"$Controller`""
 Write-Verbose "Controller Port: `"$ControllerPort`""
 Write-Verbose "Controller Slot: `"$ControllerSlot`""
 if (!($Machine -or $MachineName -or $MachineGuid)) {Write-Host "[Error] You must supply at least one machine object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imediums = @()
 if ($Disc) {
  Write-Verbose "Getting disc inventory from Disc(s) object"
  $imediums = $Disc
  $imediums = $imediums | Where-Object {$_ -ne $null}
 } # get disc inventory (by $Disc)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting disc inventory from Name(s)"
   $imediums += Get-VirtualBoxDisc -Name $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 } # get disc inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting disc inventory from GUID(s)"
   $imediums += Get-VirtualBoxDisc -Guid $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 } # get vm inventory (by $Guid)
 if ($Machine) {
  $imachines = $Machine
 } # get vm inventory (by $Machine)
 elseif ($MachineName) {
  $imachines = Get-VirtualBoxVM -Name $MachineName -SkipCheck
 } # get vm inventory (by $MachineName)
 elseif ($MachineGuid) {
  $imachines = Get-VirtualBoxVM -Guid $MachineGuid -SkipCheck
 } # get vm inventory (by $MachineGuid)
 if ($imachines) {
  foreach ($imachine in $imachines) {
   if ($imachine.State -ne 'PoweredOff') {Write-Host "[Error] The machine $($imachine.Name) is not powered off. Hotswap is not supported at this time. Power the machine off and try again." -ForegroundColor Red -BackgroundColor Black;return}
   if ($imediums) {
    Write-Verbose "[Info] Found discs"
    foreach ($imedium in $imediums) {
     Write-Verbose "Found disc: $($imedium.Name)"
     if ($imedium.MachineName) {
      # make sure it's not already attached to the requested vm
      foreach ($vmname in $imedium.MachineName) {
       Write-Verbose "Disc attached to VM: $vmname"
       if (Get-VirtualBoxDisc -MachineName $vmname -SkipCheck) {Write-Host "[Error] The disc $($imedium.Name) is already mounted to the $($imachine.Name) machine." -ForegroundColor Red -BackgroundColor Black;return}
      } # foreach $vmname in $imedium.MachineName
     } # end if $imedium.MachineName
    } # foreach $imedium in $imediums
   } # end if $imediums
   try {
    if ($ModuleHost.ToLower() -eq 'websrv') {
     <#
     #$istoragecontrollers = New-Object IStorageController
     #$istoragecontrollers = $istoragecontrollers.Fetch($imachine.Id)
     foreach ($istoragecontroller in $imachine.IStorageControllers) {
      if ($istoragecontroller.Name -eq $Controller) {
       if ($ControllerPort -lt 0 -or $ControllerPort -gt $istoragecontroller.PortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough available ports. Specify a new port number and try again." -ForegroundColor Red -BackgroundColor Black;return}
       if ($ControllerSlot -lt 0 -or $ControllerSlot -gt $istoragecontroller.MaxDevicesPerPortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough slots available on the requseted port. Specify a new slot number and try again." -ForegroundColor Red -BackgroundColor Black;return}
       $controllerfound = $true
      } # end if $istoragecontroller.Name -eq $Controller
      if (!$controllerfound) {Write-Host "[Error] The controller $Controller was not found. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black;return}
     } # foreach $istoragecontroller in $imachine.IStorageControllers
     #>
     $storagecontrollers = Get-VirtualBoxVMStorageController -Name $Controller -Machine $imachine -SkipCheck
     if (!$storagecontrollers) {Write-Host "[Error] The controller $Controller was not found on the $($imachine.Name) machine. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black}
     else {
      foreach ($storagecontroller in $storagecontrollers) {
       Write-Verbose "Getting write lock on machine $($imachine.Name)"
       $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
       # create a new machine object
       $mmachine = New-Object VirtualBoxVM
       # get the mutable machine object
       Write-Verbose "Getting the mutable machine object"
       $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
       $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
       if ($imediums) {
        foreach ($imedium in $imediums) {
         # attach the disc drive with disc
         Write-Verbose "Attaching optical drive to machine $($imachine.Name) and mounting $($imedium.Name)"
         $global:vbox.IMachine_attachDevice($mmachine.Id, $storagecontroller.Name, $ControllerPort, $ControllerSlot, [DeviceType]::new().ToULong('DVD'), $imedium.Id)
        } # foreach $imedium in $imediums
       } # end if $imediums
       else {
        # attach the disc drive
        Write-Verbose "Attaching optical drive $($imedium.Name) to machine $($imachine.Name)" 
        $global:vbox.IMachine_attachDeviceWithoutMedium($mmachine.Id, $storagecontroller.Name, $ControllerPort, $ControllerSlot, [DeviceType]::new().ToULong('DVD'))
       } # end else
       # save new settings
       Write-Verbose "Saving new settings"
       $global:vbox.IMachine_saveSettings($mmachine.Id)
       # unlock machine session
       Write-Verbose "Unlocking machine session"
       $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      } # foreach $storagecontroller in $storagecontrollers
     } # end else
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     <#
     $istoragecontrollers = $imachine.ComObject.StorageControllers
     foreach ($istoragecontroller in $istoragecontrollers) {
      if ($istoragecontroller.Name -eq $Controller) {
       if ($ControllerPort -lt 0 -or $ControllerPort -gt $istoragecontroller.PortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough available ports. Specify a new port number and try again." -ForegroundColor Red -BackgroundColor Black;return}
       if ($ControllerSlot -lt 0 -or $ControllerSlot -gt $istoragecontroller.MaxDevicesPerPortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough slots available on the requseted port. Specify a new slot number and try again." -ForegroundColor Red -BackgroundColor Black;return}
       $controllerfound = $true
      } # end if $istoragecontroller.Name -eq $Controller
      if (!$controllerfound) {Write-Host "[Error] The controller $($istoragecontroller.Name) was not found. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black;return}
     } # foreach $istoragecontroller in $istoragecontrollers
     #>
     $storagecontrollers = Get-VirtualBoxVMStorageController -Name $Controller -Machine $imachine -SkipCheck
     if (!$storagecontrollers) {Write-Host "[Error] The controller $Controller was not found on the $($imachine.Name) machine. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black}
     else {
      foreach ($storagecontroller in $storagecontrollers) {
       Write-Verbose "Getting write lock on machine $($imachine.Name)"
       $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
       # create a new machine object
       $mmachine = New-Object VirtualBoxVM
       # get the mutable machine object
       Write-Verbose "Getting the mutable machine object"
       $mmachine.ComObject = $imachine.ISession.Session.Machine
       $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
       if ($imediums) {
        foreach ($imedium in $imediums) {
         # wait for the disc to become available
         Write-Verbose "Waiting for the disc to become available"
         do {} until ($imedium.ComObject.State -eq 1)
         # attach the disc drive with disc
         Write-Verbose "Attaching optical drive to machine $($imachine.Name) and mounting $($imedium.Name)"
         $mmachine.ComObject.AttachDevice($storagecontroller.Name, $ControllerPort, $ControllerSlot, [DeviceType]::new().ToULong('DVD'), $imedium.ComObject)
        } # foreach $imedium in $imediums
       } # end if $imediums
       else {
        # attach the disc drive
        Write-Verbose "Attaching optical drive $($imedium.Name) to machine $($imachine.Name)"
        $mmachine.ComObject.AttachDeviceWithoutMedium($storagecontroller.Name, $ControllerPort, $ControllerSlot, [DeviceType]::new().ToULong('DVD'))
       } # end else
       # save new settings
       Write-Verbose "Saving new settings"
       $mmachine.ComObject.SaveSettings()
       # unlock machine session
       Write-Verbose "Unlocking machine session"
       $imachine.ISession.Session.UnlockMachine()
      } # foreach $storagecontroller in $storagecontrollers
     } # end else
    } # end elseif com
   } # Try
   catch {
    Write-Verbose 'Exception mounting virtual disc'
    Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
    Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
   } # Catch
   finally {
    # release mutable machine objects if they exist
    if ($mmachine) {
     if ($mmachine.ISession.Id) {
      # release mutable session object
      Write-Verbose "Releasing mutable session object"
      $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
     }
     if ($mmachine.ISession.Session) {
      if ($mmachine.ISession.Session.State -gt 1) {
       $mmachine.ISession.Session.UnlockMachine()
      } # end if $mmachine.ISession.Session locked
     } # end if $mmachine.ISession.Session
     if ($mmachine.Id) {
      # release mutable object
      Write-Verbose "Releasing mutable object"
      $global:vbox.IManagedObjectRef_release($mmachine.Id)
     }
    }
    # obligatory session unlock
    Write-Verbose 'Cleaning up machine sessions'
    if ($imachines) {
     foreach ($imachine in $imachines) {
      if ($imachine.ISession.Id) {
       if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
        Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
        $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
       } # end if session state not unlocked
      } # end if $imachine.ISession.Id
      if ($imachine.ISession.Session) {
       if ($imachine.ISession.Session.State -gt 1) {
        $imachine.ISession.Session.UnlockMachine()
       } # end if $imachine.ISession.Session locked
      } # end if $imachine.ISession.Session
      if ($imachine.IConsole) {
       # release the iconsole session
       Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
       $global:vbox.IManagedObjectRef_release($imachine.IConsole)
      } # end if $imachine.IConsole
      #$imachine.ISession.Id = $null
      $imachine.IConsole = $null
      if ($imachine.IPercent) {$imachine.IPercent = $null}
      $imachine.MSession = $null
      $imachine.MConsole = $null
      $imachine.MMachine = $null
     } # end foreach $imachine in $imachines
    } # end if $imachines
   } # Finally
  } # foreach $imachine in $imachines
 } # end if $imachines
 else {Write-Host "[Error] No machines found using specified filters." -ForegroundColor Red -BackgroundColor Black;return}
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Remove-VirtualBoxVMOpticalDrive {
<#
.SYNOPSIS
Dismount VirtualBox disc
.DESCRIPTION
Dismounts VirtualBox discs. The command will fail if the virtual disc is not attached to the specified virtual machine.
.PARAMETER Disc
At least one virtual disc object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER MachineName
The name of the virtual machine to dismount the disc from. This is a required parameter.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Remove-VirtualBoxVMOpticalDrive -Name TestDisc -MachineName Win10 -Controller SATA -ControllerPort 0 -ControllerSlot 0

Dismounts the virtual disc named "TestDisc.vmdk" from the Win10 virtual machine SATA controller on port 0 slot 0
.NOTES
NAME        :  Remove-VirtualBoxVMOpticalDrive
VERSION     :  1.0
LAST UPDATED:  1/20/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisc
.INPUTS
VirtualBoxVHD[]:  VirtualBoxVHDs for virtual disc objects
String[]       :  Strings for virtual disc names
GUID[]         :  GUIDS for virtual disc GUIDS
String         :  String for virtual machine name
String         :  String for controller name
Int            :  Integer for controller port
Int            :  Integer for controller slot
.OUTPUTS
None
#>
[cmdletbinding(SupportsShouldProcess,ConfirmImpact='High')]
Param(
[Parameter(Mandatory=$false,HelpMessage="Enter a storage controller name")]
[ValidateNotNullorEmpty()]
  [string]$Controller,
[Parameter(Mandatory=$false,ValueFromPipeline=$true,
HelpMessage="Enter a virtual machine object to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter the name of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [string]$MachineName,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter the GUID of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [guid]$MachineGuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - MachineName: `"$MachineName`""
 Write-Verbose "Pipeline - MachineGuid: `"$MachineGuid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if ($Machine) {
  $imachines = $Machine
 } # get vm inventory (by $Machine)
 elseif ($MachineName) {
  $imachines = Get-VirtualBoxVM -Name $MachineName -SkipCheck
 } # get vm inventory (by $MachineName)
 elseif ($MachineGuid) {
  $imachines = Get-VirtualBoxVM -Guid $MachineGuid -SkipCheck
 } # get vm inventory (by $MachineGuid)
 elseif ($Machine -or $MachineName -or $MachineGuid -and !$imachines) {
  Write-Verbose "[Warning] No machines found using provided filters. Matching all machines."
  $imachines = Get-VirtualBoxVM -SkipCheck
 } # get vm inventory (fallback)
 if (!$imachines) {$imachines = Get-VirtualBoxVM -SkipCheck}
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.State -ne 'PoweredOff') {Write-Host "[Error] The machine $($imachine.Name) is not powered off. Hotswap is not supported at this time. Power the machine off and try again." -ForegroundColor Red -BackgroundColor Black}
    if ($PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Dismount storage medium $($imedium.Name) ")) {
     if ($ModuleHost.ToLower() -eq 'websrv') {
      Write-Verbose "Getting medium attachment information"
      $imediumattachment = $global:vbox.IMachine_getMediumAttachments($imachine.Id) | Where-Object {$_.machine -match $imachine.Id} | Where-Object {$_.Controller -match $Controller} | Where-Object {$_.Type -match 'DVD'}
      Write-Verbose "Getting write lock on machine $($imachine.Name)"
      $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
      $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
      Write-Verbose "Attempting to unmount disc $($imedium.Name) from machine: $($imachine.Name)"
      $global:vbox.IMachine_detachDevice($mmachine.Id, $imediumattachment.controller, $imediumattachment.port, $imediumattachment.device)
      # save new settings
      Write-Verbose "Saving new settings"
      $global:vbox.IMachine_saveSettings($mmachine.Id)
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if websrv
     elseif ($ModuleHost.ToLower() -eq 'com') {
      Write-Verbose "Getting medium attachment information"
      $imediumattachment = ($global:vbox.Machines | Where-Object {$_.Id -match $imachine.Guid}).MediumAttachments | Where-Object {$_.Controller -match $Controller} | Where-Object {$_.Type -match [DeviceType]::new().ToULong('DVD')}
      Write-Verbose "Getting write lock on machine $($imachine.Name)"
      $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.ComObject = $imachine.ISession.Session.Machine
      $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
      Write-Verbose "Attempting to unmount disc $($imedium.Name) from machine: $($imachine.Name)"
      Write-Verbose "Controller: `"$($imediumattachment.Controller)`""
      Write-Verbose "Port: `"$($imediumattachment.Port)`""
      Write-Verbose "Device: `"$($imediumattachment.Device)`""
      $mmachine.ComObject.DetachDevice($imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device)
      # save new settings
      Write-Verbose "Saving new settings"
      $mmachine.ComObject.SaveSettings()
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $imachine.ISession.Session.UnlockMachine()
     } # end elseif com
    } # end if $PSCmdlet.ShouldProcess(
   } # foreach $imachine in $imachines
  } # end if $imachines
 } # Try
 catch {
  Write-Verbose 'Exception dismounting virtual disc'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # release mutable machine objects if they exist
  if ($mmachine) {
   if ($mmachine.ISession.Id) {
    # release mutable session object
    Write-Verbose "Releasing mutable session object"
    $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
   }
   if ($mmachine.ISession.Session) {
    if ($mmachine.ISession.Session.State -gt 1) {
     $mmachine.ISession.Session.UnlockMachine()
    } # end if $mmachine.ISession.Session locked
   } # end if $mmachine.ISession.Session
   if ($mmachine.Id) {
    # release mutable object
    Write-Verbose "Releasing mutable object"
    $global:vbox.IManagedObjectRef_release($mmachine.Id)
   }
  }
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Set-VirtualBoxVMGuestProperty {
<#
.SYNOPSIS
Set a virtual machine guest property
.DESCRIPTION
Sets a virtual machine guest property.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Property
The property name.
.PARAMETER Value
The property value.
.PARAMETER Flags
A comma-separated list of property flags. (i.e. -Flags 'name=value','name=value','name=value')
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Set-VirtualBoxVMGuestProperty -Property '/VirtualBox/GuestAdd/VBoxService/--timesync-interval' -Value 60000
Set all running virtual machines to synchronize the guest time with the host every 60 seconds (Default 10 seconds)
.EXAMPLE
PS C:\> Set-VirtualBoxVMGuestProperty -Name "2016" -Property '/VirtualBox/GuestAdd/VBoxService/--timesync-min-adjust' -Value 1000
Set the "2016 Core" virtual machine to adjust guest time in drift increments of 1 second (Default 100 milliseconds)
.EXAMPLE
PS C:\> Set-VirtualBoxVMGuestProperty -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6 -Property '/VirtualBox/GuestAdd/VBoxService/--timesync-set-threshold' -Value 30000
Set the virtual machine with GUID 7353caa6-8cb6-4066-aec9-6c6a69a001b6 to adjust guest time if out of sync by more than 30 seconds (Default 20 minutes)
.NOTES
NAME        :  Set-VirtualBoxVMGuestProperty
VERSION     :  1.0
LAST UPDATED:  1/26/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Remove-VirtualBoxVMGuestProperty
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for property name
String        :  String for property value
String[]      :  Strings for property flags
.OUTPUTS
VirtualBoxVM  :  Updated virtual machine object(s)
#>
[CmdletBinding()]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Enter the property name",
Mandatory=$false)]
  [string]$Property,
[Parameter(HelpMessage="Enter the property value",
Mandatory=$false)]
  [string]$Value,
[Parameter(HelpMessage="Enter a comma-separated list of property flags",
Mandatory=$false)]
  [string[]]$Flags,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # join flags
 [string]$Flags = $Flags -join ','
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($ModuleHost.ToLower() -eq 'websrv') {
     # set the requested property
     Write-Verbose "Setting property `"$Property`" value to `"$Value`" with flag(s) `"$Flags`" for the $($imachine.Name) machine"
     $global:vbox.IMachine_setGuestProperty($imachine.Id, $Property, $Value, $Flags)
     # output the updated machine object to the pipeline
     Write-Verbose "Outputting the updated machine object to the pipeline"
     Write-Output (Get-VirtualBoxVM -Guid $imachine.Guid -SkipCheck)
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     # set the requested property
     Write-Verbose "Setting property `"$Property`" value to `"$Value`" with flag(s) `"$Flags`" for the $($imachine.Name) machine"
     $imachine.ComObject.SetGuestProperty($Property, $Value, $Flags)
     # output the updated machine object to the pipeline
     Write-Verbose "Outputting the updated machine object to the pipeline"
     Write-Output (Get-VirtualBoxVM -Guid $imachine.Guid -SkipCheck)
    } # end elseif com
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Verbose "[Warning] No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception setting guest property'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Remove-VirtualBoxVMGuestProperty {
<#
.SYNOPSIS
Remove a virtual machine guest property
.DESCRIPTION
Removes a virtual machine guest property.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Property
The property name.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Remove-VirtualBoxVMGuestProperty -Property '/VirtualBox/GuestAdd/VBoxService/--timesync-interval'
Remove the '/VirtualBox/GuestAdd/VBoxService/--timesync-interval' property from all running virtual machines
.EXAMPLE
PS C:\> Remove-VirtualBoxVMGuestProperty -Name "2016" -Property '/VirtualBox/GuestAdd/VBoxService/--timesync-min-adjust'
Remove the '/VirtualBox/GuestAdd/VBoxService/--timesync-min-adjust' property from the "2016 Core" virtual machine
.EXAMPLE
PS C:\> Remove-VirtualBoxVMGuestProperty -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6 -Property '/VirtualBox/GuestAdd/VBoxService/--timesync-set-threshold'
Remove the '/VirtualBox/GuestAdd/VBoxService/--timesync-set-threshold' property from the virtual machine with GUID 7353caa6-8cb6-4066-aec9-6c6a69a001b6
.NOTES
NAME        :  Remove-VirtualBoxVMGuestProperty
VERSION     :  1.0
LAST UPDATED:  1/26/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Remove-VirtualBoxVMGuestProperty
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for property name
.OUTPUTS
VirtualBoxVM  :  Updated virtual machine object(s)
#>
[CmdletBinding()]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Enter the property name",
Mandatory=$false)]
  [string]$Property,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($ModuleHost.ToLower() -eq 'websrv') {
     # remove the requested property
     Write-Verbose "Removing property `"$Property`" from the $($imachine.Name) machine"
     $global:vbox.IMachine_deleteGuestProperty($imachine.Id, $Property)
     # output the updated machine object to the pipeline
     Write-Verbose "Outputting the updated machine object to the pipeline"
     Write-Output (Get-VirtualBoxVM -Guid $imachine.Guid -SkipCheck)
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     # remove the requested property
     Write-Verbose "Removing property `"$Property`" from the $($imachine.Name) machine"
     $imachine.ComObject.DeleteGuestProperty($Property)
     # output the updated machine object to the pipeline
     Write-Verbose "Outputting the updated machine object to the pipeline"
     Write-Output (Get-VirtualBoxVM -Guid $imachine.Guid -SkipCheck)
    } # end elseif com
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Verbose "[Warning] No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception removing guest property'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Enable-VirtualBoxVMVRDEServer {
<#
.SYNOPSIS
Enable VRDE server for a virtual machine
.DESCRIPTION
Enables VRDE server for a virtual machine if it is disabled.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Enable-VirtualBoxVMVRDEServer
Enable VRDE server for all running virtual machines
.EXAMPLE
PS C:\> Enable-VirtualBoxVMVRDEServer -Name "2016"
Enable VRDE server for the "2016 Core" virtual machine
.EXAMPLE
PS C:\> Enable-VirtualBoxVMVRDEServer -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6
Enable VRDE server for the virtual machine with GUID 7353caa6-8cb6-4066-aec9-6c6a69a001b6
.NOTES
NAME        :  Enable-VirtualBoxVMVRDEServer
VERSION     :  1.0
LAST UPDATED:  1/24/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Disable-VirtualBoxVMVRDEServer
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
.OUTPUTS
VirtualBoxVM  :  Updated virtual machine object(s)
#>
[CmdletBinding()]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.IVrdeServer.Enabled -eq $false) {
     if ($ModuleHost.ToLower() -eq 'websrv') {
      Write-Verbose "Getting shared lock on machine $($imachine.Name)"
      $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Shared'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
      $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
      $mmachine.IVrdeServer = $mmachine.IVrdeServer.Fetch($mmachine.Id)
      # enable VRDE server
      Write-Verbose "Enabling VRDE server for $($imachine.Name)"
      $global:vbox.IVRDEServer_setEnabled($mmachine.IVrdeServer.Id, $true)
      # save new settings
      Write-Verbose "Saving new settings"
      $global:vbox.IMachine_saveSettings($mmachine.Id)
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      # update the machine object
      Write-Verbose "Updating the machine object"
      $imachine.IVrdeServer.Update()
      # output the updated machine object to the pipeline
      Write-Verbose "Outputting the updated machine object to the pipeline"
      Write-Output $imachine
     } # end if websrv
     elseif ($ModuleHost.ToLower() -eq 'com') {
      Write-Verbose "Getting shared lock on machine $($imachine.Name)"
      $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Shared'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.ComObject = $imachine.ISession.Session.Machine
      $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
      # enable VRDE server
      Write-Verbose "Enabling VRDE server for $($imachine.Name)"
      $mmachine.ComObject.VRDEServer.Enabled = 1
      # save new settings
      Write-Verbose "Saving new settings"
      $mmachine.ComObject.SaveSettings()
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $imachine.ISession.Session.UnlockMachine()
      # output the updated machine object to the pipeline
      Write-Verbose "Outputting the updated machine object to the pipeline"
      Write-Output (Get-VirtualBoxVM -Guid $imachine.Guid -SkipCheck)
     } # end elseif com
    } # end if $imachine.IVrdeServer.Enabled -eq $false
    else {Write-Host "[Error] The VRDE server for the virtual machine `"$($imachine.Name)`" is already enabled or the state is unknown." -ForegroundColor Red -BackgroundColor Black;return}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Verbose "[Warning] No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception enabling VRDE server'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Disable-VirtualBoxVMVRDEServer {
<#
.SYNOPSIS
Disable VRDE server for a virtual machine
.DESCRIPTION
Disables VRDE server for a running virtual machine if it is enabled.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Disable-VirtualBoxVMVRDEServer
Disable VRDE server for all running virtual machines
.EXAMPLE
PS C:\> Disable-VirtualBoxVMVRDEServer -Name "2016"
Disable VRDE server for the "2016 Core" virtual machine
.EXAMPLE
PS C:\> Disable-VirtualBoxVMVRDEServer -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6
Disable VRDE server for the virtual machine with GUID 7353caa6-8cb6-4066-aec9-6c6a69a001b6
.NOTES
NAME        :  Disable-VirtualBoxVMVRDEServer
VERSION     :  1.0
LAST UPDATED:  1/24/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Enable-VirtualBoxVMVRDEServer
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
.OUTPUTS
VirtualBoxVM  :  Updated virtual machine object(s)
#>
[CmdletBinding()]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.IVrdeServer.Enabled -eq $true) {
     if ($ModuleHost.ToLower() -eq 'websrv') {
      Write-Verbose "Getting shared lock on machine $($imachine.Name)"
      $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Shared'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
      $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
      $mmachine.IVrdeServer = $mmachine.IVrdeServer.Fetch($mmachine.Id)
      # enable VRDE server
      Write-Verbose "Disabling VRDE server for $($imachine.Name)"
      $global:vbox.IVRDEServer_setEnabled($mmachine.IVrdeServer.Id, $false)
      # save new settings
      Write-Verbose "Saving new settings"
      $global:vbox.IMachine_saveSettings($mmachine.Id)
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      # update the machine object
      Write-Verbose "Updating the machine object"
      $imachine.IVrdeServer.Update()
      # output the updated machine object to the pipeline
      Write-Verbose "Outputting the updated machine object to the pipeline"
      Write-Output $imachine
     } # end if websrv
     elseif ($ModuleHost.ToLower() -eq 'com') {
      Write-Verbose "Getting shared lock on machine $($imachine.Name)"
      $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Shared'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.ComObject = $imachine.ISession.Session.Machine
      $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
      # enable VRDE server
      Write-Verbose "Disabling VRDE server for $($imachine.Name)"
      $mmachine.ISession.Session.Machine.VRDEServer.Enabled.Equals(0)
      # save new settings
      Write-Verbose "Saving new settings"
      $mmachine.ComObject.SaveSettings()
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $imachine.ISession.Session.UnlockMachine()
      # output the updated machine object to the pipeline
      Write-Verbose "Outputting the updated machine object to the pipeline"
      Write-Output (Get-VirtualBoxVM -Guid $imachine.Guid -SkipCheck)
     } # end elseif com
    } # end if $imachine.IVrdeServer.Enabled -eq $true
    else {Write-Host "[Error] The VRDE server for the virtual machine `"$($imachine.Name)`" is already disabled or the state is unknown." -ForegroundColor Red -BackgroundColor Black;return}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Verbose "[Warning] No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception disabling VRDE server'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Edit-VirtualBoxVMVRDEServer {
<#
.SYNOPSIS
Edit VRDE server for a virtual machine
.DESCRIPTION
Edits VRDE server settings for a virtual machine.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER AuthType
The authorization type for the VRDE server. Must be 'Null', 'External', or 'Guest'.
.PARAMETER AuthTimeout
The authorization timeout in milliseconds for the VRDE server.
.PARAMETER AllowMultiConnection
Enable or disable reusing a single connection to the VRDE server.
.PARAMETER ReuseSingleConnection
Enable or disable reusing a single connection to the VRDE server.
.PARAMETER VrdeExtPack
The VRDE extension pack for the VRDE server.
.PARAMETER AuthLibrary
The authorization library for the VRDE server.
.PARAMETER VrdePort
The TCP port for the VRDE server.
The authorization library for the VRDE server.
.PARAMETER IpAddress
The IP address for the VRDE server.
.PARAMETER VideoChannelEnabled
Enable or disable the video channel for the VRDE server.
.PARAMETER VideoChannelQuality
The video channel quality for the VRDE server.
.PARAMETER VideoChannelDownscaleProtection
The video channel downscale protection for the VRDE server.
.PARAMETER DisableClientDisplay
Disable or enable the client display for the VRDE server.
.PARAMETER DisableClientInput
Disable or enable the client input for the VRDE server.
.PARAMETER DisableClientAudio
Disable or enable the client audio for the VRDE server.
.PARAMETER DisableClientUsb
Disable or enable the client USB for the VRDE server.
.PARAMETER DisableClientClipboard
Disable or enable the client clipboard for the VRDE server.
.PARAMETER DisableClientUpstreamAudio
Disable or enable the client upstream audio for the VRDE server.
.PARAMETER DisableClientRdpdr
Disable or enable the client RDPDR for the VRDE server.
.PARAMETER H3dRedirectEnabled
Enable or disable the H3D redirect for the VRDE server.
.PARAMETER SecurityMethod
The security method for the VRDE server.
.PARAMETER SecurityServerCertificate
The security server certificate for the VRDE server.
.PARAMETER SecurityServerPrivateKey
The security server private key for the VRDE server.
.PARAMETER SecurityCaCertificate
The security CA certificate for the VRDE server.
.PARAMETER AudioRateCorrectionMode
The audio rate correction mode for the VRDE server.
.PARAMETER AudioLogPath
The audio log path for the VRDE server.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Edit-VirtualBoxVMVRDEServer -AuthTimeout 5000
Set the VRDE server authorization timeout to 5 seconds for all running virtual machines
.EXAMPLE
PS C:\> Edit-VirtualBoxVMVRDEServer -Name "2016" -VrdePort 3389
Set the VRDE server TCP port to 3389 for the "2016 Core" virtual machine
.EXAMPLE
PS C:\> Edit-VirtualBoxVMVRDEServer -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6 -AllowMultiConnection $true
Permit multiple connections to the VRDE server for the virtual machine with GUID 7353caa6-8cb6-4066-aec9-6c6a69a001b6
.NOTES
NAME        :  Edit-VirtualBoxVMVRDEServer
VERSION     :  1.0
LAST UPDATED:  1/25/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Enable-VirtualBoxVMVRDEServer
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for VRDE server AuthType
Uint32        :  Uint32 for VRDE server AuthTimeout
bool          :  bool for VRDE server AllowMultiConnection
bool          :  bool for VRDE server ReuseSingleConnection
String        :  String for VRDE server VrdeExtPack
String        :  String for VRDE server AuthLibrary
Uint32        :  Uint32 for VRDE server VrdePort
.OUTPUTS
VirtualBoxVM  :  Updated virtual machine object(s)
#>
[CmdletBinding()]
Param(
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Enter the authorization type for the VRDE server",
Mandatory=$false)]
[ValidateSet('Null','External','Guest')]
  [string]$AuthType,
[Parameter(HelpMessage="Enter the authorization timeout in milliseconds for the VRDE server",
Mandatory=$false)]
  [uint32]$AuthTimeout,
[Parameter(HelpMessage="Enable or disable permitting multiple simultaneous connections to the VRDE server",
Mandatory=$false)]
  [bool]$AllowMultiConnection,
[Parameter(HelpMessage="Enable or disable reusing a single connection to the VRDE server",
Mandatory=$false)]
  [bool]$ReuseSingleConnection,
[Parameter(HelpMessage="Enter the VRDE extension pack for the VRDE server",
Mandatory=$false)]
  [string]$VrdeExtPack,
[Parameter(HelpMessage="Enter the authorization library for the VRDE server",
Mandatory=$false)]
  [string]$AuthLibrary,
[Parameter(HelpMessage="Enter the TCP port for the VRDE server",
Mandatory=$false)]
  [uint32]$TcpPort,
[Parameter(HelpMessage="Enter the IP address for the VRDE server",
Mandatory=$false)]
  [string]$IpAddress,
[Parameter(HelpMessage="Enable or disable the video channel for the VRDE server",
Mandatory=$false)]
  [bool]$VideoChannelEnabled,
[Parameter(HelpMessage="Enter the video channel quality for the VRDE server",
Mandatory=$false)]
  [string]$VideoChannelQuality,
[Parameter(HelpMessage="Enter the video channel downscale protection for the VRDE server",
Mandatory=$false)]
  [string]$VideoChannelDownscaleProtection,
[Parameter(HelpMessage="Disable or enable the client display for the VRDE server",
Mandatory=$false)]
  [bool]$DisableClientDisplay,
[Parameter(HelpMessage="Disable or enable the client input for the VRDE server",
Mandatory=$false)]
  [bool]$DisableClientInput,
[Parameter(HelpMessage="Disable or enable the client audio for the VRDE server",
Mandatory=$false)]
  [bool]$DisableClientAudio,
[Parameter(HelpMessage="Disable or enable the client USB for the VRDE server",
Mandatory=$false)]
  [bool]$DisableClientUsb,
[Parameter(HelpMessage="Disable or enable the client clipboard for the VRDE server",
Mandatory=$false)]
  [bool]$DisableClientClipboard,
[Parameter(HelpMessage="Disable or enable the client upstream audio for the VRDE server",
Mandatory=$false)]
  [bool]$DisableClientUpstreamAudio,
[Parameter(HelpMessage="Disable or enable the client RDPDR for the VRDE server",
Mandatory=$false)]
  [bool]$DisableClientRdpdr,
[Parameter(HelpMessage="Enable or disable the H3D redirect for the VRDE server",
Mandatory=$false)]
  [bool]$H3dRedirectEnabled,
[Parameter(HelpMessage="Enter the security method for the VRDE server",
Mandatory=$false)]
  [string]$SecurityMethod,
[Parameter(HelpMessage="Enter the security server certificate for the VRDE server",
Mandatory=$false)]
  [string]$SecurityServerCertificate,
[Parameter(HelpMessage="Enter the security server private key for the VRDE server",
Mandatory=$false)]
  [string]$SecurityServerPrivateKey,
[Parameter(HelpMessage="Enter the security CA certificate for the VRDE server",
Mandatory=$false)]
  [string]$SecurityCaCertificate,
[Parameter(HelpMessage="Enter the audio rate correction mode for the VRDE server",
Mandatory=$false)]
  [string]$AudioRateCorrectionMode,
[Parameter(HelpMessage="Enter the audio log path for the VRDE server",
Mandatory=$false)]
  [string]$AudioLogPath,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.IVrdeServer.Enabled -eq $true) {
     if ($ModuleHost.ToLower() -eq 'websrv') {
      Write-Verbose "Getting shared lock on machine $($imachine.Name)"
      $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Shared'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
      $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
      $mmachine.IVrdeServer = $mmachine.IVrdeServer.Fetch($mmachine.Id)
      # apply custom settings as requested
      Write-Verbose "Processing VRDE server setting: AuthType"
      if ($AuthType) {$global:vbox.IVRDEServer_setAuthType($mmachine.IVrdeServer.Id, $AuthType)}
      Write-Verbose "Processing VRDE server setting: AuthTimeout"
      if ($MyInvocation.BoundParameters.Keys -contains 'AuthTimeout') {$global:vbox.IVRDEServer_setAuthTimeout($mmachine.IVrdeServer.Id, $AuthTimeout)}
      Write-Verbose "Processing VRDE server setting: AllowMultiConnection"
      if ($AllowMultiConnection) {$global:vbox.IVRDEServer_setAllowMultiConnection($mmachine.IVrdeServer.Id, $AllowMultiConnection)}
      Write-Verbose "Processing VRDE server setting: ReuseSingleConnection"
      if ($ReuseSingleConnection) {$global:vbox.IVRDEServer_setReuseSingleConnection($mmachine.IVrdeServer.Id, $ReuseSingleConnection)}
      Write-Verbose "Processing VRDE server setting: VRDEExtPack"
      if ($VRDEExtPack) {$global:vbox.IVRDEServer_setVRDEExtPack($mmachine.IVrdeServer.Id, $VRDEExtPack)}
      Write-Verbose "Processing VRDE server setting: AuthLibrary"
      if ($AuthLibrary) {$global:vbox.IVRDEServer_setAuthLibrary($mmachine.IVrdeServer.Id, $AuthLibrary)}
      Write-Verbose "Processing VRDE server setting: TcpPort"
      if ($TcpPort -ne 0) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'TCP/Ports', $TcpPort.ToString())}
      Write-Verbose "Processing VRDE server setting: IpAddress"
      if ($IpAddress) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'TCP/Address', $IpAddress)}
      Write-Verbose "Processing VRDE server setting: VideoChannelEnabled"
      if ($VideoChannelEnabled -eq $true) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'VideoChannel/Enabled', $true)}
      elseif ($VideoChannelEnabled -eq $false) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'VideoChannel/Enabled', '')}
      Write-Verbose "Processing VRDE server setting: VideoChannelQuality"
      if ($VideoChannelQuality) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'VideoChannel/Quality', $VideoChannelQuality)}
      Write-Verbose "Processing VRDE server setting: VideoChannelDownscaleProtection"
      if ($VideoChannelDownscaleProtection) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'VideoChannel/DownscaleProtection', $VideoChannelDownscaleProtection)}
      Write-Verbose "Processing VRDE server setting: DisableClientDisplay"
      if ($DisableClientDisplay -eq $true) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableDisplay', $false)}
      elseif ($DisableClientDisplay -eq $false) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableDisplay', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientInput"
      if ($DisableClientInput -eq $true) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableInput', $false)}
      elseif ($DisableClientInput -eq $false) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableInput', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientAudio"
      if ($DisableClientAudio -eq $true) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableAudio', $false)}
      elseif ($DisableClientAudio -eq $false) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableAudio', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientUsb"
      if ($DisableClientUsb -eq $true) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableUSB', $false)}
      elseif ($DisableClientUsb -eq $false) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableUSB', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientClipboard"
      if ($DisableClientClipboard -eq $true) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableClipboard', $false)}
      elseif ($DisableClientClipboard -eq $false) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableClipboard', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientUpstreamAudio"
      if ($DisableClientUpstreamAudio -eq $true) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableUpstreamAudio', $false)}
      elseif ($DisableClientUpstreamAudio -eq $false) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableUpstreamAudio', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientRdpdr"
      if ($DisableClientRdpdr -eq $true) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableRDPDR', $false)}
      elseif ($DisableClientRdpdr -eq $false) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Client/DisableRDPDR', '')}
      Write-Verbose "Processing VRDE server setting: H3dRedirectEnabled"
      if ($H3dRedirectEnabled -eq $true) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'H3DRedirect/Enabled', $true)}
      elseif ($H3dRedirectEnabled -eq $false) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'H3DRedirect/Enabled', '')}
      Write-Verbose "Processing VRDE server setting: SecurityMethod"
      if ($SecurityMethod) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Security/Method', $SecurityMethod)}
      Write-Verbose "Processing VRDE server setting: SecurityServerCertificate"
      if ($SecurityServerCertificate) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Security/ServerCertificate', $SecurityServerCertificate)}
      Write-Verbose "Processing VRDE server setting: SecurityServerPrivateKey"
      if ($SecurityServerPrivateKey) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Security/ServerPrivateKey', $SecurityServerPrivateKey)}
      Write-Verbose "Processing VRDE server setting: SecurityCaCertificate"
      if ($SecurityCaCertificate) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Security/CACertificate', $SecurityCaCertificate)}
      Write-Verbose "Processing VRDE server setting: AudioRateCorrectionMode"
      if ($AudioRateCorrectionMode) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Audio/RateCorrectionMode', $AudioRateCorrectionMode)}
      Write-Verbose "Processing VRDE server setting: AudioLogPath"
      if ($AudioLogPath) {$global:vbox.IVRDEServer_setVRDEProperty($mmachine.IVrdeServer.Id, 'Audio/LogPath', $AudioLogPath)}
      # save new settings
      Write-Verbose "Saving new settings"
      $global:vbox.IMachine_saveSettings($mmachine.Id)
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      # update the machine object
      Write-Verbose "Updating the machine object"
      $imachine.IVrdeServer.Update()
      # output the updated machine object to the pipeline
      Write-Verbose "Outputting the updated machine object to the pipeline"
      Write-Output $imachine
     } # end if websrv
     elseif ($ModuleHost.ToLower() -eq 'com') {
      Write-Verbose "Getting shared lock on machine $($imachine.Name)"
      $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Shared'))
      # create a new machine object
      $mmachine = New-Object VirtualBoxVM
      # get the mutable machine object
      Write-Verbose "Getting the mutable machine object"
      $mmachine.ComObject = $imachine.ISession.Session.Machine
      $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
      # apply custom settings as requested
      Write-Verbose "Processing VRDE server setting: AuthType"
      if ($AuthType) {$mmachine.ComObject.VRDEServer.AuthType = ([AuthType]::new().ToInt($AuthType)).ToString()}
      Write-Verbose "Processing VRDE server setting: AuthTimeout"
      if ($MyInvocation.BoundParameters.Keys -contains 'AuthTimeout') {$mmachine.ComObject.VRDEServer.AuthTimeout = $AuthTimeout}
      Write-Verbose "Processing VRDE server setting: AllowMultiConnection"
      if ($AllowMultiConnection) {$mmachine.ComObject.VRDEServer.AllowMultiConnection = [int]$AllowMultiConnection}
      Write-Verbose "Processing VRDE server setting: ReuseSingleConnection"
      if ($ReuseSingleConnection) {$mmachine.ComObject.VRDEServer.ReuseSingleConnection = [int]$ReuseSingleConnection}
      Write-Verbose "Processing VRDE server setting: VRDEExtPack"
      if ($VRDEExtPack) {$mmachine.ComObject.VRDEServer.VRDEExtPack = $VRDEExtPack}
      Write-Verbose "Processing VRDE server setting: AuthLibrary"
      if ($AuthLibrary) {$mmachine.ComObject.VRDEServer.AuthLibrary = $AuthLibrary}
      Write-Verbose "Processing VRDE server setting: TcpPort"
      if ($TcpPort -ne 0) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('TCP/Ports', $TcpPort.ToString())}
      Write-Verbose "Processing VRDE server setting: IpAddress"
      if ($IpAddress) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('TCP/Address', $IpAddress)}
      Write-Verbose "Processing VRDE server setting: VideoChannelEnabled"
      if ($VideoChannelEnabled -eq $true) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('VideoChannel/Enabled', $true)}
      elseif ($VideoChannelEnabled -eq $false) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('VideoChannel/Enabled', '')}
      Write-Verbose "Processing VRDE server setting: VideoChannelQuality"
      if ($VideoChannelQuality) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('VideoChannel/Quality', $VideoChannelQuality)}
      Write-Verbose "Processing VRDE server setting: VideoChannelDownscaleProtection"
      if ($VideoChannelDownscaleProtection) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('VideoChannel/DownscaleProtection', $VideoChannelDownscaleProtection)}
      Write-Verbose "Processing VRDE server setting: DisableClientDisplay"
      if ($DisableClientDisplay -eq $true) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableDisplay', $false)}
      elseif ($DisableClientDisplay -eq $false) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableDisplay', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientInput"
      if ($DisableClientInput -eq $true) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableInput', $false)}
      elseif ($DisableClientInput -eq $false) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableInput', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientAudio"
      if ($DisableClientAudio -eq $true) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableAudio', $false)}
      elseif ($DisableClientAudio -eq $false) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableAudio', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientUsb"
      if ($DisableClientUsb -eq $true) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableUSB', $false)}
      elseif ($DisableClientUsb -eq $false) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableUSB', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientClipboard"
      if ($DisableClientClipboard -eq $true) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableClipboard', $false)}
      elseif ($DisableClientClipboard -eq $false) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableClipboard', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientUpstreamAudio"
      if ($DisableClientUpstreamAudio -eq $true) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableUpstreamAudio', $false)}
      elseif ($DisableClientUpstreamAudio -eq $false) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableUpstreamAudio', '')}
      Write-Verbose "Processing VRDE server setting: DisableClientRdpdr"
      if ($DisableClientRdpdr -eq $true) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableRDPDR', $false)}
      elseif ($DisableClientRdpdr -eq $false) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Client/DisableRDPDR', '')}
      Write-Verbose "Processing VRDE server setting: H3dRedirectEnabled"
      if ($H3dRedirectEnabled -eq $true) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('H3DRedirect/Enabled', $true)}
      elseif ($H3dRedirectEnabled -eq $false) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('H3DRedirect/Enabled', '')}
      Write-Verbose "Processing VRDE server setting: SecurityMethod"
      if ($SecurityMethod) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Security/Method', $SecurityMethod)}
      Write-Verbose "Processing VRDE server setting: SecurityServerCertificate"
      if ($SecurityServerCertificate) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Security/ServerCertificate', $SecurityServerCertificate)}
      Write-Verbose "Processing VRDE server setting: SecurityServerPrivateKey"
      if ($SecurityServerPrivateKey) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Security/ServerPrivateKey', $SecurityServerPrivateKey)}
      Write-Verbose "Processing VRDE server setting: SecurityCaCertificate"
      if ($SecurityCaCertificate) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Security/CACertificate', $SecurityCaCertificate)}
      Write-Verbose "Processing VRDE server setting: AudioRateCorrectionMode"
      if ($AudioRateCorrectionMode) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Audio/RateCorrectionMode', $AudioRateCorrectionMode)}
      Write-Verbose "Processing VRDE server setting: AudioLogPath"
      if ($AudioLogPath) {$mmachine.ComObject.VRDEServer.SetVRDEProperty('Audio/LogPath', $AudioLogPath)}
      # save new settings
      Write-Verbose "Saving new settings"
      $mmachine.ComObject.SaveSettings()
      # unlock machine session
      Write-Verbose "Unlocking machine session"
      $imachine.ISession.Session.UnlockMachine()
      # output the updated machine object to the pipeline
      Write-Verbose "Outputting the updated machine object to the pipeline"
      Write-Output (Get-VirtualBoxVM -Guid $imachine.Guid -SkipCheck)
     } # end elseif com
    } # end if $imachine.IVrdeServer.Enabled -eq $true
    else {Write-Host "[Error] The VRDE server for the virtual machine `"$($imachine.Name)`" is already disabled or the state is unknown." -ForegroundColor Red -BackgroundColor Black;return}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Verbose "[Warning] No matching virtual machines were found using specified parameters"}
 } # Try
 catch {
  Write-Verbose 'Exception editing VRDE server'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Connect-VirtualBoxVMVRDEServer {
<#
.SYNOPSIS
Connect to a virtual machine VRDE server
.DESCRIPTION
Connects to a running virtual machine VRDE server using Remote Desktop Connection. Only windows hosts are currently supported.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Where-Object {$_.IVrdeServer.Enabled -eq $true} | Connect-VirtualBoxVMVRDEServer
Launch Remote Desktop Conenction application and connect it to the VRDE server for all running virtual machines that have it enabled
.EXAMPLE
PS C:\> Connect-VirtualBoxVMVRDEServer -Name "2016"
Launch Remote Desktop Conenction application and connect it to the VRDE server for the "2016 Core" virtual machine
.EXAMPLE
PS C:\> Connect-VirtualBoxVMVRDEServer -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6
Launch Remote Desktop Conenction application and connect it to the VRDE server for the virtual machine with GUID 7353caa6-8cb6-4066-aec9-6c6a69a001b6
.NOTES
NAME        :  Connect-VirtualBoxVMVRDEServer
VERSION     :  1.0
LAST UPDATED:  1/24/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
None
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
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
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Name",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Guid",Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 if ($imachines) {
  foreach ($imachine in $imachines) {
   if (Test-Path 'C:\Windows\System32\mstsc.exe') {
    Write-Verbose "Launching Remote Desktop Connection window for `"$($imachine.Name)`""
    Write-Verbose "Command: mstsc.exe /v:$($global:hostaddress):$($imachine.IVrdeServer.VrdePort)"
    Start-Process -FilePath "mstsc.exe" -ArgumentList ("/v:$($global:hostaddress):$($imachine.IVrdeServer.VrdePort)")
   } # end if VBoxSDL.exe exists
   else {Write-Host "[Error] Remote Desktop Connection client not found.";return}
  } # foreach $imachine in $imachines
 } # end if $imachines
 else {Write-Verbose "[Warning] No matching virtual machines were found using specified parameters"}
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Import-VirtualBoxOVF {
<#
.SYNOPSIS
Import an OVF file
.DESCRIPTION
Imports an OVF file and creates a new virtual machine based on its settings. If an OVF file is specified, all files specified by the .ovf file must be in the same folder as the .ovf file or this command will fail. The name provided by the Name parameter must not exist in the VirtualBox inventory, or this command will fail. You can optionally supply custom values using a large number of parameters available to this command. There are too many to fully document in this help text, so tab completion has been added where it is possible. The values provided by tab completion are updated when Start-VirtualBoxSession is successfully run. To force the values to be updated again, use the -Force switch with Start-VirtualBoxSession.
.PARAMETER FileName
The full path of the OVF file. This is a required parameter.
.PARAMETER ImportOptions
You may optionally provide import options. They must be supplied separated by commas. (Ex. -ImportOptions 'KeepAllMACs','ImportToVDI')
.PARAMETER BaseFolder
A custom base folder for the virtual machine.
.PARAMETER CdRom
A custom CD/DVD ROM for the virtual machine. This parameter is currently broken.
.PARAMETER CloudBootDiskSize
A custom cloud boot disk size for the virtual machine.
.PARAMETER CloudBootVolumeId
A custom cloud boot volume ID for the virtual machine.
.PARAMETER CloudBucket
A custom cloud bucket for the virtual machine.
.PARAMETER CloudDomain
A custom cloud domain for the virtual machine.
.PARAMETER CloudImageDisplayName
A custom cloud image display name for the virtual machine.
.PARAMETER CloudImageState
A custom cloud image state for the virtual machine.
.PARAMETER CloudInstanceDisplayName
A custom cloud instance display name for the virtual machine.
.PARAMETER CloudInstanceShape
A custom cloud instance shape for the virtual machine.
.PARAMETER CloudKeepObject
A custom cloud keep object for the virtual machine.
.PARAMETER CloudLaunchInstance
A custom cloud launch instance for the virtual machine.
.PARAMETER CloudOciLaunchMode
A custom cloud OCI launch mode for the virtual machine.
.PARAMETER CloudOciSubnet
A custom cloud OCI subnet for the virtual machine. This must be a valid IP address.
.PARAMETER CloudOciSubnetCompartment
A custom cloud OCI subnet compartment for the virtual machine.
.PARAMETER CloudOciVcn
A custom cloud OCI VCN for the virtual machine.
.PARAMETER CloudOciVcnCompartment
A custom cloud OCI VCN compartment for the virtual machine.
.PARAMETER CloudPrivateIp
A custom cloud private IP address for the virtual machine. This must be a valid IP address.
.PARAMETER CloudProfileName
A custom cloud profile name for the virtual machine.
.PARAMETER CloudPublicIp
A custom cloud public IP address for the virtual machine. This must be a valid IP address.
.PARAMETER CloudPublicSshKey
A custom cloud public SSH key for the virtual machine.
.PARAMETER CpuCount
A custom CPU count for the virtual machine. Must be a valid count reported by VirtualBox.
.PARAMETER Description
A custom description for the virtual machine.
.PARAMETER FirmwareType
A custom firmware type for the virtual machine. Must be a valid firmware type reported by VirtualBox.
.PARAMETER Floppy
A custom floppy disk controller for the virtual machine. This parameter is currently broken.
.PARAMETER HardDiskControllerIdePrimary
Force the addition of a primary IDE controller of a specified type for the virtual machine. Type specified must be either 'PIIX3' or 'PIIX3'.
.PARAMETER HardDiskControllerIdeSecondary
Force the addition of a secondary IDE controller of a specified type for the virtual machine. Type specified must be either 'PIIX3' or 'PIIX3'.
.PARAMETER HardDiskControllerSas
A switch to force the addition of a SAS controller for the virtual machine.
.PARAMETER HardDiskControllerSata
A switch to force the addition of a SATA controller for the virtual machine.
.PARAMETER HardDiskControllerScsi
Force the addition of an SCSI controller of a specified type for the virtual machine. Type specified must be either 'LsiLogic' or 'Bus-Logic'.
.PARAMETER HardDiskImage
A custom hard disk image for the virtual machine. This parameter is currently broken.
.PARAMETER License
A custom license for the virtual machine.
.PARAMETER MemorySize
A custom memory size in MB for the virtual machine. Must be a valid size reported by VirtualBox.
.PARAMETER Miscellaneous
Reserved for future use.
.PARAMETER Name
A custom name for the virtual machine.
.PARAMETER NetworkAdapter
A custom network adapter for the virtual machine. This parameter is currently broken.
.PARAMETER OsTypeId
A custom OS type ID for the virtual machine. Must be a valid OS type ID reported by VirtualBox.
.PARAMETER PrimaryGroup
A custom primary group for the virtual machine.
.PARAMETER Product
A custom product identifier for the virtual machine.
.PARAMETER ProductUrl
A custom product URL for the virtual machine.
.PARAMETER SettingsFile
A custom settings file for the virtual machine.
.PARAMETER SoundCard
A switch to force the addition of a sound card for the virtual machine.
.PARAMETER UsbController
A custom USB Controller for the virtual machine.
.PARAMETER Vendor
A custom vendor identifier for the virtual machine.
.PARAMETER VendorUrl
A custom vendor URL for the virtual machine.
.PARAMETER Version
A custom version for the virtual machine.
.PARAMETER ProgressBar
A switch to display a progress bar.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Import-VirtualBoxOVF -FileName "C:\OVA Files\Win10.ova" -ProgressBar
Imports the Win10.ova file with all VirtualBox recommended defaults and displays a progress bar
.EXAMPLE
PS C:\> Import-VirtualBoxOVF -FileName "C:\OVF Files\Win10\Win10.ovf" -Name "My Win10 OVF VM" -OsTypeId Windows10_64
Imports the Win10.ovf file as a new virtual machine named "My Win10 OVF VM" with the OS ID overridden to 64bit Windows10
.NOTES
NAME        :  Import-VirtualBoxOVF
VERSION     :  0.9
LAST UPDATED:  1/18/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
None
.INPUTS
String        :  String for full OVA/OVF file path
String[]      :  Strings for import options
Other optional input parameters available. Use "Get-Help Import-VirtualBoxOVF -Full" for a complete list.
.OUTPUTS
None
#>
[CmdletBinding(DefaultParameterSetName='Template')]
Param(
[Parameter(HelpMessage="Enter the full path to the OVF file",
Mandatory=$true,Position=0)]
[ValidateScript({Test-Path $_})]
[ValidateNotNullorEmpty()]
  [string]$FileName,
[Parameter(HelpMessage="Enter optional import option(s) separated by commas",
Mandatory=$false)]
[ValidateSet('KeepAllMACs','KeepNATMACs','ImportToVDI')]
[ValidateNotNullorEmpty()]
  [string[]]$ImportOptions = ' ',
[Parameter(HelpMessage="Enter custom virtual machine name",
ParameterSetName='Custom',Mandatory=$false)]
  [string]$Name,
[Parameter(HelpMessage="Enter custom primary virtual machine group",
ParameterSetName='Custom',Mandatory=$false)]
  [string]$PrimaryGroup,
[Parameter(HelpMessage="Enter custom full path to the settings file",
ParameterSetName='Custom',Mandatory=$false)]
  [string]$SettingsFile,
[Parameter(HelpMessage="Enter custom base folder path for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$BaseFolder,
[Parameter(HelpMessage="Enter custom description for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Description,
[Parameter(HelpMessage="This might not work properly - needs more testing",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$UsbController,
[Parameter(HelpMessage="This does not work properly yet",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$NetworkAdapter,
[Parameter(HelpMessage="Enter custom CD/DVD ROM for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CdRom,
[Parameter(HelpMessage="Enter custom SCSI hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
[ValidateSet('LsiLogic','Bus-Logic')]
  [string]$HardDiskControllerScsi,
[Parameter(HelpMessage="Enter custom primary IDE hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
[ValidateSet('PIIX3','PIIX4')]
  [string]$HardDiskControllerIdePrimary,
[Parameter(HelpMessage="Enter custom secondary IDE hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
[ValidateSet('PIIX3','PIIX4')]
  [string]$HardDiskControllerIdeSecondary,
[Parameter(HelpMessage="A switch to force the addition of a SATA hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [switch]$HardDiskControllerSata,
[Parameter(HelpMessage="A switch to force the addition of a SAS hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [switch]$HardDiskControllerSas,
[Parameter(HelpMessage="This does not work properly yet",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$HardDiskImage,
[Parameter(HelpMessage="Enter custom product identifier for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Product,
[Parameter(HelpMessage="Enter custom vendor identifier for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Vendor,
[Parameter(HelpMessage="Enter custom version for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Version,
[Parameter(HelpMessage="Enter custom product URL for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$ProductUrl,
[Parameter(HelpMessage="Enter custom vendor URL for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$VendorUrl,
[Parameter(HelpMessage="Enter custom license for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$License,
[Parameter(HelpMessage="Reserved for future use",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Miscellaneous,
[Parameter(HelpMessage="This does not work properly yet",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Floppy,
[Parameter(HelpMessage="A switch to force the addition of a sound card to the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [switch]$SoundCard,
[Parameter(HelpMessage="Enter custom cloud instance shape for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudInstanceShape,
[Parameter(HelpMessage="Enter custom cloud domain for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudDomain,
[Parameter(HelpMessage="Enter custom cloud boot disk size for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudBootDiskSize,
[Parameter(HelpMessage="Enter custom cloud bucket for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudBucket,
[Parameter(HelpMessage="Enter custom cloud OCI VCN for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudOciVcn,
[Parameter(HelpMessage="Enter custom cloud public IP address for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [ipaddress]$CloudPublicIp,
[Parameter(HelpMessage="Enter custom cloud private IP address for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [ipaddress]$CloudPrivateIp,
[Parameter(HelpMessage="Enter custom cloud OCI subnet for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [ipaddress]$CloudOciSubnet,
[Parameter(HelpMessage="Enter custom cloud profile name for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudProfileName,
[Parameter(HelpMessage="Enter custom cloud keep object setting for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudKeepObject,
[Parameter(HelpMessage="Enter custom cloud launch instance for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudLaunchInstance,
[Parameter(HelpMessage="Enter custom cloud image state for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudImageState,
[Parameter(HelpMessage="Enter custom instance display name for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudInstanceDisplayName,
[Parameter(HelpMessage="Enter custom image display name for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudImageDisplayName,
[Parameter(HelpMessage="Enter custom cloud OCI launch mode for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudOciLaunchMode,
[Parameter(HelpMessage="Enter custom cloud boot volume ID for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudBootVolumeId,
[Parameter(HelpMessage="Enter custom cloud OCI VCN compartment setting for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudOciVcnCompartment,
[Parameter(HelpMessage="Enter custom cloud OCI subnet compartment setting for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudOciSubnetCompartment,
[Parameter(HelpMessage="Enter custom cloud public SSH key for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudPublicSshKey,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
DynamicParam {
 $CustomAttributes = New-Object System.Management.Automation.ParameterAttribute
 $CustomAttributes.Mandatory = $false
 $CustomAttributes.ParameterSetName = 'Custom'
 $CustomAttributes.HelpMessage = 'Enter custom type ID for the virtual machine guest OS'
 $OsTypeIdCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $OsTypeIdCollection.Add($CustomAttributes)
 $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute(@('Other','Other_64','Windows31','Windows95','Windows98','WindowsMe','WindowsNT3x','WindowsNT4','Windows2000','WindowsXP','WindowsXP_64','Windows2003','Windows2003_64','WindowsVista','WindowsVista_64','Windows2008','Windows2008_64','Windows7','Windows7_64','Windows8','Windows8_64','Windows81','Windows81_64','Windows2012_64','Windows10','Windows10_64','Windows2016_64','Windows2019_64','WindowsNT','WindowsNT_64','Linux22','Linux24','Linux24_64','Linux26','Linux26_64','ArchLinux','ArchLinux_64','Debian','Debian_64','Fedora','Fedora_64','Gentoo','Gentoo_64','Mandriva','Mandriva_64','Oracle','Oracle_64','RedHat','RedHat_64','OpenSUSE','OpenSUSE_64','Turbolinux','Turbolinux_64','Ubuntu','Ubuntu_64','Xandros','Xandros_64','Linux','Linux_64','Solaris','Solaris_64','OpenSolaris','OpenSolaris_64','Solaris11_64','FreeBSD','FreeBSD_64','OpenBSD','OpenBSD_64','NetBSD','NetBSD_64','OS2Warp3','OS2Warp4','OS2Warp45','OS2eCS','OS21x','OS2','MacOS','MacOS_64','MacOS106','MacOS106_64','MacOS107_64','MacOS108_64','MacOS109_64','MacOS1010_64','MacOS1011_64','MacOS1012_64','MacOS1013_64','DOS','Netware','L4','QNX','JRockitVE','VBoxBS_64'))
 if ($global:guestostype.id) {
  $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute($global:guestostype.id)
 }
 $OsTypeIdCollection.Add($ValidateSetOsTypeId)
 $OsTypeId = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("OsTypeId", [string], $OsTypeIdCollection)
 $CustomAttributes.HelpMessage = 'Enter custom number of CPUs available to the virtual machine'
 $CpuCountCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $CpuCountCollection.Add($CustomAttributes)
 $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute(1, 32)
 if ($global:systempropertiessupported.MinGuestCPUCount -and $global:systempropertiessupported.MaxGuestCPUCount) {
  $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestCPUCount, $global:systempropertiessupported.MaxGuestCPUCount)
 }
 $CpuCountCollection.Add($ValidateSetCpuCount)
 $CpuCount = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("CpuCount", [uint64], $CpuCountCollection)
 $CustomAttributes.HelpMessage = 'Enter custom memory size in MB for the virtual machine'
 $MemorySizeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $MemorySizeCollection.Add($CustomAttributes)
 $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute(4, 2097152)
 if ($global:systempropertiessupported.MinGuestRam -and $global:systempropertiessupported.MaxGuestRam) {
  $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestRam, $global:systempropertiessupported.MaxGuestRam)
 }
 $MemorySizeCollection.Add($ValidateSetMemorySize)
 $MemorySize = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("MemorySize", [uint64], $MemorySizeCollection)
 $CustomAttributes.HelpMessage = 'Enter custom firmware type for the virtual machine'
 $FirmwareTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $FirmwareTypesCollection.Add($CustomAttributes)
 $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('BIOS','EFI','EFI32','EFI64','EFIDUAL'))
 if ($global:systempropertiessupported.FirmwareTypes) {
  $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.FirmwareTypes)
 }
 $FirmwareTypesCollection.Add($ValidateSetFirmwareTypes)
 $FirmwareTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("FirmwareType", [string], $FirmwareTypesCollection)
 $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
 $paramDictionary.Add("OsTypeId", $OsTypeId)
 $paramDictionary.Add("CpuCount", $CpuCount)
 $paramDictionary.Add("MemorySize", $MemorySize)
 $paramDictionary.Add("FirmwareType", $FirmwareTypes)
 return $paramDictionary
}
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 $OsTypeId = $PSBoundParameters['OsTypeId']
 $CpuCount = $PSBoundParameters['CpuCount']
 $MemorySize = $PSBoundParameters['MemorySize']
 $FirmwareType = $PSBoundParameters['FirmwareType']
} # Begin
Process {
 if ($Name) {if ((Get-VirtualBoxVM -Name $Name -SkipCheck).Name -eq $Name) {Write-Host "[Error] Machine $Name already exists. Enter another name and try again." -ForegroundColor Red -BackgroundColor Black;return}}
 try {
  if ($ModuleHost.ToLower() -eq 'websrv') {
   # create a vm shell
   Write-Verbose "Creating a shell machine object"
   $imachine = New-Object VirtualBoxVM
   # create an appliance shell
   Write-Verbose "Creating a shell appliance object"
   $iappliance = $global:vbox.IVirtualBox_createAppliance($global:ivbox)
   # read the ovf/ova file
   Write-Verbose "Reading the OVf/OVA settings file"
   $imachine.IProgress.Id = $global:vbox.IAppliance_read($iappliance, $FileName)
   # collect iprogress data
   Write-Verbose "Fetching IProgress data"
   $imachine.IProgress = $imachine.IProgress.Fetch($imachine.IProgress.Id)
   if ($ProgressBar) {Write-Progress -Activity "Reading OVF file" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1}
   do {
    # update iprogress data
    $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
    if ($ProgressBar) {
     if ($imachine.IProgress.Percent -lt 20) {
      Write-Progress -Activity "Reading OVF file" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1
      if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
     } # end if $imachine.IProgress.Percent < 20
     else {
      Write-Progress -Activity "Reading OVF file" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)
      if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
     } # end else
    } # end if $ProgressBar
   } until ($imachine.IProgress.Percent -eq 100 -and $imachine.IProgress.Completed -eq $true) # continue once completed
   if ($imachine.IProgress.ResultCode -ne 0) {Write-Verbose $imachine.IProgress.ErrorInfo}
   # interpret the iappliance
   Write-Verbose "Interpreting the OVF/OVA settings"
   $global:vbox.IAppliance_interpret($iappliance)
   # get warnings and display to verbose output
   Write-Verbose "Getting any warnings in reading the OVf/OVA settings file"
   [string[]]$warnings = $global:vbox.IAppliance_getWarnings($iappliance)
   foreach ($warning in $warnings) {
    Write-Verbose $warning
   }
   # get the $ivirtualsystemdescriptions object reference(s) found by interperet()
   Write-Verbose "Getting the IVirtualSystemDescriptions object reference(s) found by interpereter"
   [string[]]$ivirtualsystemdescriptions = $global:vbox.IAppliance_getVirtualSystemDescriptions($iappliance)
   $appliancedescriptions = New-Object IVirtualSystemDescription
   # get an array of iappliance config values to modify before import
   foreach ($ivirtualsystemdescription in $ivirtualsystemdescriptions) {
    # populate the appliance descriptions
    Write-Verbose "Getting appliance descriptions"
    [array]$appliancedescriptions += $appliancedescriptions.Fetch($ivirtualsystemdescription)
    # remove null rows
    $appliancedescriptions = $appliancedescriptions | Where-Object {$_.Types -ne $null}
    if ($PsCmdlet.ParameterSetName -eq 'Custom') {
     Write-Verbose "Setting requested custom appliance setting(s)"
     if ($Name) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Name'}).VBoxValues = $Name
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Name'}).Options = $true
     }
     if ($OsTypeId) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'OS'}).VBoxValues = $OsTypeId
      ($appliancedescriptions | Where-Object {$_.Types -eq 'OS'}).Options = $true
     }
     if ($PrimaryGroup) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'PrimaryGroup'}).VBoxValues = $PrimaryGroup
      ($appliancedescriptions | Where-Object {$_.Types -eq 'PrimaryGroup'}).Options = $true
     }
     if ($SettingsFile) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'SettingsFile'}).VBoxValues = $SettingsFile
      ($appliancedescriptions | Where-Object {$_.Types -eq 'SettingsFile'}).Options = $true
     }
     if ($BaseFolder) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'BaseFolder'}).VBoxValues = $BaseFolder
      ($appliancedescriptions | Where-Object {$_.Types -eq 'BaseFolder'}).Options = $true
     }
     if ($Description) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Description'}).VBoxValues = $Description
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Description'}).Options = $true
     }
     if ($CpuCount) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CPU'}).VBoxValues = $CpuCount
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CPU'}).Options = $true
     }
     if ($MemorySize) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Memory'}).VBoxValues = $MemorySize
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Memory'}).Options = $true
     }
     if ($UsbController) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'USBController'}).VBoxValues = $USBController
      ($appliancedescriptions | Where-Object {$_.Types -eq 'USBController'}).Options = $true
     }
     if ($NetworkAdapter) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'NetworkAdapter'}).VBoxValues = $NetworkAdapter
      ($appliancedescriptions | Where-Object {$_.Types -eq 'NetworkAdapter'}).Options = $true
     }
     if ($CdRom) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CDROM'}).ExtraConfigValues = $CdRom
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CDROM'}).Options = $true
     }
     if ($HardDiskControllerScsi) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).VBoxValues = $HardDiskControllerScsi
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).Refs = '0'
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).Options = $true
     }
     if ($HardDiskControllerIdePrimary) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'} | Where-Object {$_.Refs -ne '6'}).VBoxValues = $HardDiskControllerIdePrimary
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'} | Where-Object {$_.Refs -ne '6'}).Options = $true
     }
     if ($HardDiskControllerIdeSecondary) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'} | Where-Object {$_.Refs -eq '6'}).VBoxValues = $HardDiskControllerIdeSecondary
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'} | Where-Object {$_.Refs -eq '6'}).Options = $true
     }
     if ($HardDiskControllerSata) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSATA'}).Refs = '0'
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSATA'}).Options = $true
     }
     if ($HardDiskControllerSas) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).VBoxValues = 'LsiLogicSas'
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).Refs = '0'
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).Options = $true
     }
     if ($HardDiskImage) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskImage'}).VBoxValues = $HardDiskImage
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskImage'}).Options = $true
     }
     if ($Product) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Product'}).VBoxValues = $Product
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Product'}).Options = $true
     }
     if ($Vendor) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Vendor'}).VBoxValues = $Vendor
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Vendor'}).Options = $true
     }
     if ($Version) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Version'}).VBoxValues = $Version
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Version'}).Options = $true
     }
     if ($ProductUrl) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'ProductUrl'}).VBoxValues = $ProductUrl
      ($appliancedescriptions | Where-Object {$_.Types -eq 'ProductUrl'}).Options = $true
     }
     if ($VendorUrl) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'VendorUrl'}).VBoxValues = $VendorUrl
      ($appliancedescriptions | Where-Object {$_.Types -eq 'VendorUrl'}).Options = $true
     }
     if ($License) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'License'}).VBoxValues = $License
      ($appliancedescriptions | Where-Object {$_.Types -eq 'License'}).Options = $true
     }
     if ($Miscellaneous) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Miscellaneous'}).VBoxValues = $Miscellaneous
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Miscellaneous'}).Options = $true
     }
     if ($Floppy) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Floppy'}).ExtraConfigValues = $Floppy
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Floppy'}).Options = $true
     }
     if ($SoundCard) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'SoundCard'}).Options = $true
     }
     if ($CloudInstanceShape) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceShape'}).VBoxValues = $CloudInstanceShape
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceShape'}).Options = $true
     }
     if ($CloudDomain) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudDomain'}).VBoxValues = $CloudDomain
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudDomain'}).Options = $true
     }
     if ($CloudBootDiskSize) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootDiskSize'}).VBoxValues = $CloudBootDiskSize
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootDiskSize'}).Options = $true
     }
     if ($CloudBucket) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBucket'}).VBoxValues = $CloudBucket
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBucket'}).Options = $true
     }
     if ($CloudOciVcn) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCN'}).VBoxValues = $CloudOciVcn
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCN'}).Options = $true
     }
     if ($CloudPublicIp) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicIP'}).VBoxValues = $CloudPublicIp
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicIP'}).Options = $true
     }
     if ($CloudProfileName) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudProfileName'}).VBoxValues = $CloudProfileName
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudProfileName'}).Options = $true
     }
     if ($CloudOciSubnet) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnet'}).VBoxValues = $CloudOciSubnet
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnet'}).Options = $true
     }
     if ($CloudKeepObject) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudKeepObject'}).VBoxValues = $CloudKeepObject
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudKeepObject'}).Options = $true
     }
     if ($CloudLaunchInstance) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudLaunchInstance'}).VBoxValues = $CloudLaunchInstance
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudLaunchInstance'}).Options = $true
     }
     if ($CloudImageState) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageState'}).VBoxValues = $CloudImageState
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageState'}).Options = $true
     }
     if ($CloudInstanceDisplayName) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceDisplayName'}).VBoxValues = $CloudInstanceDisplayName
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceDisplayName'}).Options = $true
     }
     if ($CloudImageDisplayName) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageDisplayName'}).VBoxValues = $CloudImageDisplayName
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageDisplayName'}).Options = $true
     }
     if ($CloudOciLaunchMode) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCILaunchMode'}).VBoxValues = $CloudOciLaunchMode
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCILaunchMode'}).Options = $true
     }
     if ($CloudPrivateIp) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPrivateIP'}).VBoxValues = $CloudPrivateIp
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPrivateIP'}).Options = $true
     }
     if ($CloudBootVolumeId) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootVolumeId'}).VBoxValues = $CloudBootVolumeId
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootVolumeId'}).Options = $true
     }
     if ($CloudOciVcnCompartment) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCNCompartment'}).VBoxValues = $CloudOciVcnCompartment
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCNCompartment'}).Options = $true
     }
     if ($CloudOciSubnetCompartment) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnetCompartment'}).VBoxValues = $CloudOciSubnetCompartment
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnetCompartment'}).Options = $true
     }
     if ($CloudPublicSshKey) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicSSHKey'}).VBoxValues = $CloudPublicSshKey
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicSSHKey'}).Options = $true
     }
     if ($FirmwareType) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'BootingFirmware'}).VBoxValues = $FirmwareType
      ($appliancedescriptions | Where-Object {$_.Types -eq 'BootingFirmware'}).Options = $true
     }
    }
    Write-Verbose "Applying final settings to appliance"
    $global:vbox.IVirtualSystemDescription_setFinalValues($ivirtualsystemdescription, $appliancedescriptions.Options, $appliancedescriptions.VBoxValues, $appliancedescriptions.ExtraConfigValues)
   } # foreach $ivirtualsystemdescription in $ivirtualsystemdescriptions
   # import the machine to inventory
   Write-Verbose "Importing machine to VirtualBox inventory"
   $imachine.IProgress.Id = $global:vbox.IAppliance_importMachines($iappliance, [ImportOptions]::new().ToInt($ImportOptions))
   # collect iprogress data
   Write-Verbose "Fetching IProgress data"
   $imachine.IProgress = $imachine.IProgress.Fetch($imachine.IProgress.Id)
   if ($ProgressBar) {Write-Progress -Activity "Importing VM $($imachine.Name)" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1}
   do {
    # update iprogress data
    $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
    if ($ProgressBar) {
     if ($imachine.IProgress.Percent -lt 20) {
      Write-Progress -Activity "Importing VM $($imachine.Name)" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1
      if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
     } # end if $imachine.IProgress.Percent < 20
     else {
      Write-Progress -Activity "Importing VM $($imachine.Name)" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)
      if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
     } # end else
    } # end if $ProgressBar
   } until ($imachine.IProgress.Percent -eq 100 -and $imachine.IProgress.Completed -eq $true) # continue once completed
   if ($imachine.IProgress.ResultCode -ne 0) {Write-Verbose $imachine.IProgress.ErrorInfo}
  } # end if websrv
  elseif ($ModuleHost.ToLower() -eq 'com') {
   # create a vm shell
   Write-Verbose "Creating a shell machine object"
   $imachine = New-Object VirtualBoxVM
   # create an appliance shell
   Write-Verbose "Creating a shell appliance object"
   $iappliance = $global:vbox.CreateAppliance()
   # read the ovf/ova file
   Write-Verbose "Reading the OVf/OVA settings file"
   $imachine.IProgress.Progress = $iappliance.Read($FileName)
   if ($ProgressBar) {Write-Progress -Activity "Reading OVF file" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1}
   do {
    # update iprogress data
    if ($ProgressBar) {
     if ($imachine.IProgress.Progress.Percent -lt 20) {
      Write-Progress -Activity "Reading OVF file" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1
      if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
     } # end if $imachine.IProgress.Progress.Percent < 20
     else {
      Write-Progress -Activity "Reading OVF file" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.Progress.TimeRemaining)
      if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
     } # end else
    } # end if $ProgressBar
   } until ($imachine.IProgress.Progress.Percent -eq 100 -and $imachine.IProgress.Progress.Completed -eq $true) # continue once completed
   if ($imachine.IProgress.Progress.ResultCode -ne 0) {Write-Verbose $imachine.IProgress.Progress.ErrorInfo}
   # interpret the iappliance
   Write-Verbose "Interpreting the OVF/OVA settings"
   $iappliance.Interpret()
   # get warnings and display to verbose output
   Write-Verbose "Getting any warnings in reading the OVf/OVA settings file"
   [string[]]$warnings = $iappliance.GetWarnings()
   foreach ($warning in $warnings) {
    Write-Verbose $warning
   }
   # create virtual system description
   #Write-Verbose "Creating virtual system description"
   #$iappliance.CreateVirtualSystemDescriptions(1)
   # get the $ivirtualsystemdescriptions object reference(s) found by interperet()
   Write-Verbose "Getting the IVirtualSystemDescriptions object reference(s) found by interpereter"
   $ivirtualsystemdescriptions = $iappliance.VirtualSystemDescriptions
   $appliancedescriptions = New-Object IVirtualSystemDescription
   # get an array of iappliance config values to modify before import
   foreach ($ivirtualsystemdescription in $ivirtualsystemdescriptions) {
    # populate the appliance descriptions
    Write-Verbose "Getting appliance descriptions"
    [array]$appliancedescriptions += $appliancedescriptions.FetchCom($ivirtualsystemdescription)
    # remove null rows
    $appliancedescriptions = $appliancedescriptions | Where-Object {$_.Types -ne $null}
    if ($PsCmdlet.ParameterSetName -eq 'Custom') {
     Write-Verbose "Setting requested custom appliance setting(s)"
     if ($Name) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Name'}).VBoxValues = $Name
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Name'}).Options = $true
     }
     if ($OsTypeId) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'OS'}).VBoxValues = $OsTypeId
      ($appliancedescriptions | Where-Object {$_.Types -eq 'OS'}).Options = $true
     }
     if ($PrimaryGroup) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'PrimaryGroup'}).VBoxValues = $PrimaryGroup
      ($appliancedescriptions | Where-Object {$_.Types -eq 'PrimaryGroup'}).Options = $true
     }
     if ($SettingsFile) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'SettingsFile'}).VBoxValues = $SettingsFile
      ($appliancedescriptions | Where-Object {$_.Types -eq 'SettingsFile'}).Options = $true
     }
     if ($BaseFolder) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'BaseFolder'}).VBoxValues = $BaseFolder
      ($appliancedescriptions | Where-Object {$_.Types -eq 'BaseFolder'}).Options = $true
     }
     if ($Description) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Description'}).VBoxValues = $Description
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Description'}).Options = $true
     }
     if ($CpuCount) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CPU'}).VBoxValues = $CpuCount
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CPU'}).Options = $true
     }
     if ($MemorySize) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Memory'}).VBoxValues = $MemorySize
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Memory'}).Options = $true
     }
     if ($UsbController) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'USBController'}).VBoxValues = $USBController
      ($appliancedescriptions | Where-Object {$_.Types -eq 'USBController'}).Options = $true
     }
     if ($NetworkAdapter) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'NetworkAdapter'}).VBoxValues = $NetworkAdapter
      ($appliancedescriptions | Where-Object {$_.Types -eq 'NetworkAdapter'}).Options = $true
     }
     if ($CdRom) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CDROM'}).ExtraConfigValues = $CdRom
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CDROM'}).Options = $true
     }
     if ($HardDiskControllerScsi) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).VBoxValues = $HardDiskControllerScsi
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).Refs = '0'
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).Options = $true
     }
     if ($HardDiskControllerIdePrimary) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'} | Where-Object {$_.Refs -ne '6'}).VBoxValues = $HardDiskControllerIdePrimary
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'} | Where-Object {$_.Refs -ne '6'}).Options = $true
     }
     if ($HardDiskControllerIdeSecondary) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'} | Where-Object {$_.Refs -eq '6'}).VBoxValues = $HardDiskControllerIdeSecondary
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'} | Where-Object {$_.Refs -eq '6'}).Options = $true
     }
     if ($HardDiskControllerSata) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSATA'}).Refs = '0'
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSATA'}).Options = $true
     }
     if ($HardDiskControllerSas) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).VBoxValues = 'LsiLogicSas'
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).Refs = '0'
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).Options = $true
     }
     if ($HardDiskImage) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskImage'}).VBoxValues = $HardDiskImage
      ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskImage'}).Options = $true
     }
     if ($Product) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Product'}).VBoxValues = $Product
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Product'}).Options = $true
     }
     if ($Vendor) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Vendor'}).VBoxValues = $Vendor
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Vendor'}).Options = $true
     }
     if ($Version) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Version'}).VBoxValues = $Version
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Version'}).Options = $true
     }
     if ($ProductUrl) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'ProductUrl'}).VBoxValues = $ProductUrl
      ($appliancedescriptions | Where-Object {$_.Types -eq 'ProductUrl'}).Options = $true
     }
     if ($VendorUrl) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'VendorUrl'}).VBoxValues = $VendorUrl
      ($appliancedescriptions | Where-Object {$_.Types -eq 'VendorUrl'}).Options = $true
     }
     if ($License) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'License'}).VBoxValues = $License
      ($appliancedescriptions | Where-Object {$_.Types -eq 'License'}).Options = $true
     }
     if ($Miscellaneous) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Miscellaneous'}).VBoxValues = $Miscellaneous
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Miscellaneous'}).Options = $true
     }
     if ($Floppy) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Floppy'}).ExtraConfigValues = $Floppy
      ($appliancedescriptions | Where-Object {$_.Types -eq 'Floppy'}).Options = $true
     }
     if ($SoundCard) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'SoundCard'}).Options = $true
     }
     if ($CloudInstanceShape) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceShape'}).VBoxValues = $CloudInstanceShape
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceShape'}).Options = $true
     }
     if ($CloudDomain) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudDomain'}).VBoxValues = $CloudDomain
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudDomain'}).Options = $true
     }
     if ($CloudBootDiskSize) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootDiskSize'}).VBoxValues = $CloudBootDiskSize
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootDiskSize'}).Options = $true
     }
     if ($CloudBucket) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBucket'}).VBoxValues = $CloudBucket
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBucket'}).Options = $true
     }
     if ($CloudOciVcn) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCN'}).VBoxValues = $CloudOciVcn
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCN'}).Options = $true
     }
     if ($CloudPublicIp) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicIP'}).VBoxValues = $CloudPublicIp
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicIP'}).Options = $true
     }
     if ($CloudProfileName) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudProfileName'}).VBoxValues = $CloudProfileName
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudProfileName'}).Options = $true
     }
     if ($CloudOciSubnet) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnet'}).VBoxValues = $CloudOciSubnet
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnet'}).Options = $true
     }
     if ($CloudKeepObject) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudKeepObject'}).VBoxValues = $CloudKeepObject
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudKeepObject'}).Options = $true
     }
     if ($CloudLaunchInstance) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudLaunchInstance'}).VBoxValues = $CloudLaunchInstance
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudLaunchInstance'}).Options = $true
     }
     if ($CloudImageState) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageState'}).VBoxValues = $CloudImageState
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageState'}).Options = $true
     }
     if ($CloudInstanceDisplayName) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceDisplayName'}).VBoxValues = $CloudInstanceDisplayName
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceDisplayName'}).Options = $true
     }
     if ($CloudImageDisplayName) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageDisplayName'}).VBoxValues = $CloudImageDisplayName
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageDisplayName'}).Options = $true
     }
     if ($CloudOciLaunchMode) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCILaunchMode'}).VBoxValues = $CloudOciLaunchMode
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCILaunchMode'}).Options = $true
     }
     if ($CloudPrivateIp) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPrivateIP'}).VBoxValues = $CloudPrivateIp
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPrivateIP'}).Options = $true
     }
     if ($CloudBootVolumeId) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootVolumeId'}).VBoxValues = $CloudBootVolumeId
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootVolumeId'}).Options = $true
     }
     if ($CloudOciVcnCompartment) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCNCompartment'}).VBoxValues = $CloudOciVcnCompartment
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCNCompartment'}).Options = $true
     }
     if ($CloudOciSubnetCompartment) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnetCompartment'}).VBoxValues = $CloudOciSubnetCompartment
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnetCompartment'}).Options = $true
     }
     if ($CloudPublicSshKey) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicSSHKey'}).VBoxValues = $CloudPublicSshKey
      ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicSSHKey'}).Options = $true
     }
     if ($FirmwareType) {
      ($appliancedescriptions | Where-Object {$_.Types -eq 'BootingFirmware'}).VBoxValues = $FirmwareType
      ($appliancedescriptions | Where-Object {$_.Types -eq 'BootingFirmware'}).Options = $true
     }
    }
    Write-Verbose "Converting options to integers"
    foreach ($option in $appliancedescriptions.Options) {[int[]]$finalOptions += $option}
    Write-Verbose "Converting values to strings"
    foreach ($value in $appliancedescriptions.VBoxValues) {[string[]]$finalValues += $value}
    Write-Verbose "Converting extra config values to strings"
    foreach ($extraconfigvalue in $appliancedescriptions.ExtraConfigValues) {[string[]]$finalExtraConfigValues += $extraconfigvalue}
    Write-Verbose "Applying final settings to appliance"
    $ivirtualsystemdescription.SetFinalValues($finalOptions, $finalValues, $finalExtraConfigValues)
   } # foreach $ivirtualsystemdescription in $ivirtualsystemdescriptions
   # import the machine to inventory
   Write-Verbose "Importing machine to VirtualBox inventory"
   $imachine.IProgress.Progress = $iappliance.ImportMachines([ImportOptions]::new().ToInt($ImportOptions))
   if ($ProgressBar) {Write-Progress -Activity "Importing VM $($imachine.Name)" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1}
   do {
    # update iprogress data
    if ($ProgressBar) {
     if ($imachine.IProgress.Progress.Percent -lt 20) {
      Write-Progress -Activity "Importing VM $($imachine.Name)" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1
      if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
     } # end if $imachine.IProgress.Progress.Percent < 20
     else {
      Write-Progress -Activity "Importing VM $($imachine.Name)" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.Progress.TimeRemaining)
      if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
     } # end else
    } # end if $ProgressBar
   } until ($imachine.IProgress.Progress.Percent -eq 100 -and $imachine.IProgress.Progress.Completed -eq $true) # continue once completed
   if ($imachine.IProgress.Progress.ResultCode -ne 0) {Write-Verbose $imachine.IProgress.Progress.ErrorInfo}
  } # end elseif com
 } # Try
 catch {
  Write-Verbose 'Exception importing OVF'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.ISession.Session) {
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if $imachine.ISession.Session locked
    } # end if $imachine.ISession.Session
    if ($imachine.IConsole) {
     # release the iconsole session
     Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
    $imachine.MSession = $null
    $imachine.MConsole = $null
    $imachine.MMachine = $null
   } # end foreach $imachine in $imachines
  } # end if $imachines
  if ($ivirtualsystemdescriptions -and $ModuleHost.ToLower() -eq 'websrv') {
   foreach ($ivirtualsystemdescription in $ivirtualsystemdescriptions) {
    $global:vbox.IManagedObjectRef_release($ivirtualsystemdescription)
   }
  }
  if ($iappliance -and $ModuleHost.ToLower() -eq 'websrv') {
   $global:vbox.IManagedObjectRef_release($iappliance)
  }
  $ivirtualsystemdescriptions = $null
  $iappliance = $null
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Export-VirtualBoxOVF {
<#
.SYNOPSIS
Export a virtual machine
.DESCRIPTION
Exports a virtual machine to an OVF file based on its settings. If an OVF file is specified, all files specified by the .ovf file will be created in the same folder as the .ovf file. A machine object, name, or GUID must be provided by one of the parameters or this command will fail. You can optionally supply custom values using a large number of parameters available to this command. There are too many to fully document in this help text, so tab completion has been added where it is possible. The values provided by tab completion are updated when Start-VirtualBoxSession is successfully run. To force the values to be updated again, use the -Force switch with Start-VirtualBoxSession.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER FilePath
The full path of the OVF file. This is a required parameter.
.PARAMETER ExportOptions
You may optionally provide import options. They must be supplied separated by commas. (Ex. -ImportOptions 'KeepAllMACs','ImportToVDI')
.PARAMETER OvfFormat
Specify the OVF format of the OVF file to be created. Default value is 'ovf-2.0'.  If this parameter is set to 'opc-1.0' and the -Ova switch is used, the command will fail.
.PARAMETER BaseFolder
A custom base folder for the virtual machine.
.PARAMETER CdRom
A custom CD/DVD ROM for the virtual machine. This parameter is currently broken.
.PARAMETER CloudBootDiskSize
A custom cloud boot disk size for the virtual machine.
.PARAMETER CloudBootVolumeId
A custom cloud boot volume ID for the virtual machine.
.PARAMETER CloudBucket
A custom cloud bucket for the virtual machine.
.PARAMETER CloudDomain
A custom cloud domain for the virtual machine.
.PARAMETER CloudImageDisplayName
A custom cloud image display name for the virtual machine.
.PARAMETER CloudImageState
A custom cloud image state for the virtual machine.
.PARAMETER CloudInstanceDisplayName
A custom cloud instance display name for the virtual machine.
.PARAMETER CloudInstanceShape
A custom cloud instance shape for the virtual machine.
.PARAMETER CloudKeepObject
A custom cloud keep object for the virtual machine.
.PARAMETER CloudLaunchInstance
A custom cloud launch instance for the virtual machine.
.PARAMETER CloudOciLaunchMode
A custom cloud OCI launch mode for the virtual machine.
.PARAMETER CloudOciSubnet
A custom cloud OCI subnet for the virtual machine. This must be a valid IP address.
.PARAMETER CloudOciSubnetCompartment
A custom cloud OCI subnet compartment for the virtual machine.
.PARAMETER CloudOciVcn
A custom cloud OCI VCN for the virtual machine.
.PARAMETER CloudOciVcnCompartment
A custom cloud OCI VCN compartment for the virtual machine.
.PARAMETER CloudPrivateIp
A custom cloud private IP address for the virtual machine. This must be a valid IP address.
.PARAMETER CloudProfileName
A custom cloud profile name for the virtual machine.
.PARAMETER CloudPublicIp
A custom cloud public IP address for the virtual machine. This must be a valid IP address.
.PARAMETER CloudPublicSshKey
A custom cloud public SSH key for the virtual machine.
.PARAMETER CpuCount
A custom CPU count for the virtual machine. Must be a valid count reported by VirtualBox.
.PARAMETER Description
A custom description for the virtual machine.
.PARAMETER FirmwareType
A custom firmware type for the virtual machine. Must be a valid firmware type reported by VirtualBox.
.PARAMETER Floppy
A custom floppy disk controller for the virtual machine. This parameter is currently broken.
.PARAMETER HardDiskControllerIdePrimary
Force the addition of a primary IDE controller of a specified type for the virtual machine. Type specified must be either 'PIIX3' or 'PIIX3'.
.PARAMETER HardDiskControllerIdeSecondary
Force the addition of a secondary IDE controller of a specified type for the virtual machine. Type specified must be either 'PIIX3' or 'PIIX3'.
.PARAMETER HardDiskControllerSas
A switch to force the addition of a SAS controller for the virtual machine.
.PARAMETER HardDiskControllerSata
A switch to force the addition of a SATA controller for the virtual machine.
.PARAMETER HardDiskControllerScsi
Force the addition of an SCSI controller of a specified type for the virtual machine. Type specified must be either 'LsiLogic' or 'Bus-Logic'.
.PARAMETER HardDiskImage
A custom hard disk image for the virtual machine. This parameter is currently broken.
.PARAMETER License
A custom license for the virtual machine.
.PARAMETER MemorySize
A custom memory size in MB for the virtual machine. Must be a valid size reported by VirtualBox.
.PARAMETER Miscellaneous
Reserved for future use.
.PARAMETER Name
A custom name for the virtual machine.
.PARAMETER NetworkAdapter
A custom network adapter for the virtual machine. This parameter is currently broken.
.PARAMETER OsTypeId
A custom OS type ID for the virtual machine. Must be a valid OS type ID reported by VirtualBox.
.PARAMETER PrimaryGroup
A custom primary group for the virtual machine.
.PARAMETER Product
A custom product identifier for the virtual machine.
.PARAMETER ProductUrl
A custom product URL for the virtual machine.
.PARAMETER SettingsFile
A custom settings file for the virtual machine.
.PARAMETER SoundCard
A switch to force the addition of a sound card for the virtual machine.
.PARAMETER UsbController
A custom USB Controller for the virtual machine.
.PARAMETER Vendor
A custom vendor identifier for the virtual machine.
.PARAMETER VendorUrl
A custom vendor URL for the virtual machine.
.PARAMETER Version
A custom version for the virtual machine.
.PARAMETER Ova
A switch to specify an OVA file is to be created. If OVF format is specified as 'opc-1.0' and this switch is used, the command will fail.
.PARAMETER ProgressBar
A switch to display a progress bar.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Export-VirtualBoxOVF -Name "My Win10 OVA VM" -FileName "C:\OVA Files" -ProgressBar
Exports the "My Win10 OVA VM" virtual machine to the "C:\OVA Files\My Win10 OVA VM.ova" file with all VirtualBox recommended defaults and displays a progress bar
.EXAMPLE
PS C:\> Export-VirtualBoxOVF -Name "My Win10 OVF VM" -FileName "C:\OVF Files\My Win10 OVF VM" -ExportOptions 'CreateManifest','ExportDVDImages'
Exports the "My Win10 OVF VM" virtual machine and all other required files, including attached CD/DVD images to "C:\OVF Files\My Win10 OVF VM\" and creates a manifest file
.NOTES
NAME        :  Export-VirtualBoxOVF
VERSION     :  0.9
LAST UPDATED:  1/18/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
None
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for OVA/OVF file path
String[]      :  Strings for export options
String        :  String for OVF format
Other optional input parameters available. Use "Get-Help Export-VirtualBoxOVF -Full" for a complete list.
.OUTPUTS
None
#>
[CmdletBinding(DefaultParameterSetName='Template')]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)"
,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Enter the full path to where the OVF file will be saved",
Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string]$FilePath,
[Parameter(HelpMessage="Enter optional import option(s) separated by commas",
Mandatory=$false)]
[ValidateSet('ovf-0.9','ovf-1.0','ovf-2.0','opc-1.0')]
[ValidateNotNullorEmpty()]
  [string[]]$OvfFormat = 'ovf-2.0',
[Parameter(HelpMessage="Enter optional import option(s) separated by commas",
Mandatory=$false)]
[ValidateSet('CreateManifest','ExportDVDImages','StripAllMACs','StripAllNonNATMACs')]
[ValidateNotNullorEmpty()]
  [string[]]$ExportOptions,
[Parameter(HelpMessage="Enter custom primary virtual machine group",
ParameterSetName='Custom',Mandatory=$false)]
  [string]$PrimaryGroup,
[Parameter(HelpMessage="Enter custom full path to the settings file",
ParameterSetName='Custom',Mandatory=$false)]
  [string]$SettingsFile,
[Parameter(HelpMessage="Enter custom base folder path for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$BaseFolder,
[Parameter(HelpMessage="Enter custom description for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Description,
[Parameter(HelpMessage="This might not work properly - needs more testing",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$UsbController,
[Parameter(HelpMessage="This does not work properly yet",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$NetworkAdapter,
[Parameter(HelpMessage="Enter custom CD/DVD ROM for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CdRom,
[Parameter(HelpMessage="Enter custom SCSI hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
[ValidateSet('LsiLogic','Bus-Logic')]
  [string]$HardDiskControllerScsi,
[Parameter(HelpMessage="Enter custom primary IDE hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
[ValidateSet('PIIX3','PIIX4')]
  [string]$HardDiskControllerIdePrimary,
[Parameter(HelpMessage="Enter custom secondary IDE hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
[ValidateSet('PIIX3','PIIX4')]
  [string]$HardDiskControllerIdeSecondary,
[Parameter(HelpMessage="A switch to force the addition of a SATA hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [switch]$HardDiskControllerSata,
[Parameter(HelpMessage="A switch to force the addition of a SAS hard disk controller for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [switch]$HardDiskControllerSas,
[Parameter(HelpMessage="This does not work properly yet",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$HardDiskImage,
[Parameter(HelpMessage="Enter custom product identifier for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Product,
[Parameter(HelpMessage="Enter custom vendor identifier for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Vendor,
[Parameter(HelpMessage="Enter custom version for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Version,
[Parameter(HelpMessage="Enter custom product URL for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$ProductUrl,
[Parameter(HelpMessage="Enter custom vendor URL for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$VendorUrl,
[Parameter(HelpMessage="Enter custom license for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$License,
[Parameter(HelpMessage="Reserved for future use",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Miscellaneous,
[Parameter(HelpMessage="This does not work properly yet",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$Floppy,
[Parameter(HelpMessage="A switch to force the addition of a sound card to the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [switch]$SoundCard,
[Parameter(HelpMessage="Enter custom cloud instance shape for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudInstanceShape,
[Parameter(HelpMessage="Enter custom cloud domain for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudDomain,
[Parameter(HelpMessage="Enter custom cloud boot disk size for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudBootDiskSize,
[Parameter(HelpMessage="Enter custom cloud bucket for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudBucket,
[Parameter(HelpMessage="Enter custom cloud OCI VCN for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudOciVcn,
[Parameter(HelpMessage="Enter custom cloud public IP address for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [ipaddress]$CloudPublicIp,
[Parameter(HelpMessage="Enter custom cloud private IP address for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [ipaddress]$CloudPrivateIp,
[Parameter(HelpMessage="Enter custom cloud OCI subnet for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [ipaddress]$CloudOciSubnet,
[Parameter(HelpMessage="Enter custom cloud profile name for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudProfileName,
[Parameter(HelpMessage="Enter custom cloud keep object setting for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudKeepObject,
[Parameter(HelpMessage="Enter custom cloud launch instance for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudLaunchInstance,
[Parameter(HelpMessage="Enter custom cloud image state for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudImageState,
[Parameter(HelpMessage="Enter custom instance display name for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudInstanceDisplayName,
[Parameter(HelpMessage="Enter custom image display name for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudImageDisplayName,
[Parameter(HelpMessage="Enter custom cloud OCI launch mode for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudOciLaunchMode,
[Parameter(HelpMessage="Enter custom cloud boot volume ID for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudBootVolumeId,
[Parameter(HelpMessage="Enter custom cloud OCI VCN compartment setting for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudOciVcnCompartment,
[Parameter(HelpMessage="Enter custom cloud OCI subnet compartment setting for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudOciSubnetCompartment,
[Parameter(HelpMessage="Enter custom cloud public SSH key for the virtual machine",
ParameterSetName="Custom",Mandatory=$false)]
  [string]$CloudPublicSshKey,
[Parameter(HelpMessage="Use this switch to write an OVA file instead of an OVF")]
  [switch]$Ova,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
DynamicParam {
 $CustomAttributes = New-Object System.Management.Automation.ParameterAttribute
 $CustomAttributes.Mandatory = $false
 $CustomAttributes.ParameterSetName = 'Custom'
 $CustomAttributes.HelpMessage = 'Enter custom type ID for the virtual machine guest OS'
 $OsTypeIdCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $OsTypeIdCollection.Add($CustomAttributes)
 $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute(@('Other','Other_64','Windows31','Windows95','Windows98','WindowsMe','WindowsNT3x','WindowsNT4','Windows2000','WindowsXP','WindowsXP_64','Windows2003','Windows2003_64','WindowsVista','WindowsVista_64','Windows2008','Windows2008_64','Windows7','Windows7_64','Windows8','Windows8_64','Windows81','Windows81_64','Windows2012_64','Windows10','Windows10_64','Windows2016_64','Windows2019_64','WindowsNT','WindowsNT_64','Linux22','Linux24','Linux24_64','Linux26','Linux26_64','ArchLinux','ArchLinux_64','Debian','Debian_64','Fedora','Fedora_64','Gentoo','Gentoo_64','Mandriva','Mandriva_64','Oracle','Oracle_64','RedHat','RedHat_64','OpenSUSE','OpenSUSE_64','Turbolinux','Turbolinux_64','Ubuntu','Ubuntu_64','Xandros','Xandros_64','Linux','Linux_64','Solaris','Solaris_64','OpenSolaris','OpenSolaris_64','Solaris11_64','FreeBSD','FreeBSD_64','OpenBSD','OpenBSD_64','NetBSD','NetBSD_64','OS2Warp3','OS2Warp4','OS2Warp45','OS2eCS','OS21x','OS2','MacOS','MacOS_64','MacOS106','MacOS106_64','MacOS107_64','MacOS108_64','MacOS109_64','MacOS1010_64','MacOS1011_64','MacOS1012_64','MacOS1013_64','DOS','Netware','L4','QNX','JRockitVE','VBoxBS_64'))
 if ($global:guestostype.id) {
  $ValidateSetOsTypeId = New-Object System.Management.Automation.ValidateSetAttribute($global:guestostype.id)
 }
 $OsTypeIdCollection.Add($ValidateSetOsTypeId)
 $OsTypeId = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("OsTypeId", [string], $OsTypeIdCollection)
 $CustomAttributes.HelpMessage = 'Enter custom number of CPUs available to the virtual machine'
 $CpuCountCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $CpuCountCollection.Add($CustomAttributes)
 $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute(1, 32)
 if ($global:systempropertiessupported.MinGuestCPUCount -and $global:systempropertiessupported.MaxGuestCPUCount) {
  $ValidateSetCpuCount = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestCPUCount, $global:systempropertiessupported.MaxGuestCPUCount)
 }
 $CpuCountCollection.Add($ValidateSetCpuCount)
 $CpuCount = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("CpuCount", [uint64], $CpuCountCollection)
 $CustomAttributes.HelpMessage = 'Enter custom memory size in MB for the virtual machine'
 $MemorySizeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $MemorySizeCollection.Add($CustomAttributes)
 $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute(4, 2097152)
 if ($global:systempropertiessupported.MinGuestRam -and $global:systempropertiessupported.MaxGuestRam) {
  $ValidateSetMemorySize = New-Object System.Management.Automation.ValidateRangeAttribute($global:systempropertiessupported.MinGuestRam, $global:systempropertiessupported.MaxGuestRam)
 }
 $MemorySizeCollection.Add($ValidateSetMemorySize)
 $MemorySize = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("MemorySize", [uint64], $MemorySizeCollection)
 $CustomAttributes.HelpMessage = 'Enter custom firmware type for the virtual machine'
 $FirmwareTypesCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
 $FirmwareTypesCollection.Add($CustomAttributes)
 $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute(@('BIOS','EFI','EFI32','EFI64','EFIDUAL'))
 if ($global:systempropertiessupported.FirmwareTypes) {
  $ValidateSetFirmwareTypes = New-Object System.Management.Automation.ValidateSetAttribute($global:systempropertiessupported.FirmwareTypes)
 }
 $FirmwareTypesCollection.Add($ValidateSetFirmwareTypes)
 $FirmwareTypes = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("FirmwareType", [string], $FirmwareTypesCollection)
 $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
 $paramDictionary.Add("OsTypeId", $OsTypeId)
 $paramDictionary.Add("CpuCount", $CpuCount)
 $paramDictionary.Add("MemorySize", $MemorySize)
 $paramDictionary.Add("FirmwareType", $FirmwareTypes)
 return $paramDictionary
}
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 $OsTypeId = $PSBoundParameters['OsTypeId']
 $CpuCount = $PSBoundParameters['CpuCount']
 $MemorySize = $PSBoundParameters['MemorySize']
 $FirmwareType = $PSBoundParameters['FirmwareType']
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 if ($Ova -and $OvfFormat -eq 'opc-1.0') {Write-Host "[Error] An OPC-1.0 file cannot be created as an OVA file. Remove one of the parameters and try again." -ForegroundColor Red -BackgroundColor Black;return}
 elseif ($Ova) {$Ext = 'ova'}
 elseif ($OvfFormat -eq 'opc-1.0') {$Ext = 'tar.gz'}
 else {$Ext = 'ovf'}
 if (!(Test-Path $FilePath)) {New-Item -ItemType Directory -Path $FilePath -Force -Confirm:$false | Write-Verbose}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($ModuleHost.ToLower() -eq 'websrv') {
     # create an appliance shell
     Write-Verbose "Creating a shell appliance object"
     $iappliance = $global:vbox.IVirtualBox_createAppliance($global:ivbox)
     # populate the appliance with the requested virtual machine's settings
     Write-Verbose "Getting the IVirtualSystemDescriptions object reference(s) found by interpereter"
     [string[]]$ivirtualsystemdescriptions = $global:vbox.IMachine_exportTo($imachine.Id, $iappliance, $FilePath)
     $appliancedescriptions = New-Object IVirtualSystemDescription
     # get an array of iappliance config values to modify before export
     foreach ($ivirtualsystemdescription in $ivirtualsystemdescriptions) {
      # populate the appliance descriptions
      Write-Verbose "Getting appliance descriptions"
      [array]$appliancedescriptions += $appliancedescriptions.Fetch($ivirtualsystemdescription)
      # remove null rows
      $appliancedescriptions = $appliancedescriptions | Where-Object {$_.Types -ne $null}
      Write-Verbose "Applying requested custom appliance setting(s)"
      if ($OsTypeId) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'OS'}).VBoxValues = $OsTypeId
       ($appliancedescriptions | Where-Object {$_.Types -eq 'OS'}).Options = $true
      }
      if ($Name) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Name'}).VBoxValues = $Name
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Name'}).Options = $true
      }
      if ($PrimaryGroup) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'PrimaryGroup'}).VBoxValues = $PrimaryGroup
       ($appliancedescriptions | Where-Object {$_.Types -eq 'PrimaryGroup'}).Options = $true
      }
      if ($SettingsFile) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'SettingsFile'}).VBoxValues = $SettingsFile
       ($appliancedescriptions | Where-Object {$_.Types -eq 'SettingsFile'}).Options = $true
      }
      if ($BaseFolder) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'BaseFolder'}).VBoxValues = $BaseFolder
       ($appliancedescriptions | Where-Object {$_.Types -eq 'BaseFolder'}).Options = $true
      }
      if ($Description) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Description'}).VBoxValues = $Description
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Description'}).Options = $true
      }
      if ($Cpu) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CPU'}).VBoxValues = $Cpu
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CPU'}).Options = $true
      }
      if ($Memory) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Memory'}).VBoxValues = $Memory
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Memory'}).Options = $true
      }
      if ($UsbController) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'USBController'}).VBoxValues = $USBController
       ($appliancedescriptions | Where-Object {$_.Types -eq 'USBController'}).Options = $true
      }
      if ($NetworkAdapter) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'NetworkAdapter'}).VBoxValues = $NetworkAdapter
       ($appliancedescriptions | Where-Object {$_.Types -eq 'NetworkAdapter'}).Options = $true
      }
      if ($CdRom) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CDROM'}).VBoxValues = $CdRom
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CDROM'}).Options = $true
      }
      if ($HardDiskControllerScsi) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).VBoxValues = $HardDiskControllerScsi
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).Options = $true
      }
      if ($HardDiskControllerIde) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'}).VBoxValues = $HardDiskControllerIde
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'}).Options = $true
      }
      if ($HardDiskImage) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskImage'}).VBoxValues = $HardDiskImage
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskImage'}).Options = $true
      }
      if ($Product) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Product'}).VBoxValues = $Product
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Product'}).Options = $true
      }
      if ($Vendor) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Vendor'}).VBoxValues = $Vendor
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Vendor'}).Options = $true
      }
      if ($Version) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Version'}).VBoxValues = $Version
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Version'}).Options = $true
      }
      if ($ProductUrl) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'ProductUrl'}).VBoxValues = $ProductUrl
       ($appliancedescriptions | Where-Object {$_.Types -eq 'ProductUrl'}).Options = $true
      }
      if ($VendorUrl) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'VendorUrl'}).VBoxValues = $VendorUrl
       ($appliancedescriptions | Where-Object {$_.Types -eq 'VendorUrl'}).Options = $true
      }
      if ($License) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'License'}).VBoxValues = $License
       ($appliancedescriptions | Where-Object {$_.Types -eq 'License'}).Options = $true
      }
      if ($Miscellaneous) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Miscellaneous'}).VBoxValues = $Miscellaneous
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Miscellaneous'}).Options = $true
      }
      if ($HardDiskControllerSata) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSATA'}).VBoxValues = $HardDiskControllerSata
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSATA'}).Options = $true
      }
      if ($HardDiskControllerSas) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).VBoxValues = $HardDiskControllerSas
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).Options = $true
      }
      if ($Floppy) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Floppy'}).VBoxValues = $Floppy
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Floppy'}).Options = $true
      }
      if ($SoundCard) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'SoundCard'}).VBoxValues = $SoundCard
       ($appliancedescriptions | Where-Object {$_.Types -eq 'SoundCard'}).Options = $true
      }
      if ($CloudInstanceShape) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceShape'}).VBoxValues = $CloudInstanceShape
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceShape'}).Options = $true
      }
      if ($CloudDomain) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudDomain'}).VBoxValues = $CloudDomain
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudDomain'}).Options = $true
      }
      if ($CloudBootDiskSize) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootDiskSize'}).VBoxValues = $CloudBootDiskSize
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootDiskSize'}).Options = $true
      }
      if ($CloudBucket) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBucket'}).VBoxValues = $CloudBucket
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBucket'}).Options = $true
      }
      if ($CloudOciVcn) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCN'}).VBoxValues = $CloudOciVcn
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCN'}).Options = $true
      }
      if ($CloudPublicIp) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicIP'}).VBoxValues = $CloudPublicIp
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicIP'}).Options = $true
      }
      if ($CloudProfileName) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudProfileName'}).VBoxValues = $CloudProfileName
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudProfileName'}).Options = $true
      }
      if ($CloudOciSubnet) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnet'}).VBoxValues = $CloudOciSubnet
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnet'}).Options = $true
      }
      if ($CloudKeepObject) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudKeepObject'}).VBoxValues = $CloudKeepObject
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudKeepObject'}).Options = $true
      }
      if ($CloudLaunchInstance) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudLaunchInstance'}).VBoxValues = $CloudLaunchInstance
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudLaunchInstance'}).Options = $true
      }
      if ($CloudImageState) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageState'}).VBoxValues = $CloudImageState
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageState'}).Options = $true
      }
      if ($CloudInstanceDisplayName) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceDisplayName'}).VBoxValues = $CloudInstanceDisplayName
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceDisplayName'}).Options = $true
      }
      if ($CloudImageDisplayName) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageDisplayName'}).VBoxValues = $CloudImageDisplayName
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageDisplayName'}).Options = $true
      }
      if ($CloudOciLaunchMode) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCILaunchMode'}).VBoxValues = $CloudOciLaunchMode
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCILaunchMode'}).Options = $true
      }
      if ($CloudPrivateIp) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPrivateIP'}).VBoxValues = $CloudPrivateIp
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPrivateIP'}).Options = $true
      }
      if ($CloudBootVolumeId) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootVolumeId'}).VBoxValues = $CloudBootVolumeId
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootVolumeId'}).Options = $true
      }
      if ($CloudOciVcnCompartment) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCNCompartment'}).VBoxValues = $CloudOciVcnCompartment
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCNCompartment'}).Options = $true
      }
      if ($CloudOciSubnetCompartment) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnetCompartment'}).VBoxValues = $CloudOciSubnetCompartment
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnetCompartment'}).Options = $true
      }
      if ($CloudPublicSshKey) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicSSHKey'}).VBoxValues = $CloudPublicSshKey
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicSSHKey'}).Options = $true
      }
      if ($BootingFirmware) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'BootingFirmware'}).VBoxValues = $BootingFirmware
       ($appliancedescriptions | Where-Object {$_.Types -eq 'BootingFirmware'}).Options = $true
      }
      $global:vbox.IVirtualSystemDescription_setFinalValues($ivirtualsystemdescription, $appliancedescriptions.Options, $appliancedescriptions.VBoxValues, $appliancedescriptions.ExtraConfigValues)
     } # foreach $ivirtualsystemdescription in $ivirtualsystemdescriptions
     # export the machine to disk
     Write-Verbose "Writing OVF to disk"
     $imachine.IProgress.Id = $global:vbox.IAppliance_write($iappliance, $OvfFormat, [ExportOptions]::new().ToInt($ExportOptions), (Join-Path -ChildPath "$($imachine.Name).$($Ext)" -Path $FilePath))
     # collect iprogress data
     Write-Verbose "Fetching IProgress data"
     $imachine.IProgress = $imachine.IProgress.Fetch($imachine.IProgress.Id)
     if ($ProgressBar) {Write-Progress -Activity "Exporting VM $($imachine.Name)" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1}
     do {
      # update iprogress data
      $imachine.IProgress = $imachine.IProgress.Update($imachine.IProgress.Id)
      if ($ProgressBar) {
       if ($imachine.IProgress.Percent -lt 20) {
        Write-Progress -Activity "Exporting VM $($imachine.Name)" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1
        if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
       } # end if $imachine.IProgress.Percent < 20
       else {
        Write-Progress -Activity "Exporting VM $($imachine.Name)" -status "$($imachine.IProgress.Description): $($imachine.IProgress.Percent)%" -percentComplete ($imachine.IProgress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.TimeRemaining)
        if ($imachine.IProgress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.OperationDescription)" -status "$($imachine.IProgress.OperationDescription): $($imachine.IProgress.OperationPercent)%" -percentComplete ($imachine.IProgress.OperationPercent) -Id 2 -ParentId 1}
       } # end else
      } # end if $ProgressBar
     } until ($imachine.IProgress.Percent -eq 100 -and $imachine.IProgress.Completed -eq $true) # continue once completed
     if ($imachine.IProgress.ResultCode -ne 0) {Write-Verbose $imachine.IProgress.ErrorInfo}
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     # create an appliance shell
     Write-Verbose "Creating a shell appliance object"
     $iappliance = $global:vbox.CreateAppliance()
     # populate the appliance with the requested virtual machine's settings
     Write-Verbose "Getting the IVirtualSystemDescriptions object reference(s) found by interpereter"
     $ivirtualsystemdescriptions = $imachine.ComObject.ExportTo($iappliance, $FilePath)
     $appliancedescriptions = New-Object IVirtualSystemDescription
     # get an array of iappliance config values to modify before export
     foreach ($ivirtualsystemdescription in $ivirtualsystemdescriptions) {
      # populate the appliance descriptions
      Write-Verbose "Getting appliance descriptions"
      [array]$appliancedescriptions += $appliancedescriptions.FetchCom($ivirtualsystemdescription)
      # remove null rows
      $appliancedescriptions = $appliancedescriptions | Where-Object {$_.Types -ne $null}
      Write-Verbose "Applying requested custom appliance setting(s)"
      if ($OsTypeId) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'OS'}).VBoxValues = $OsTypeId
       ($appliancedescriptions | Where-Object {$_.Types -eq 'OS'}).Options = $true
      }
      if ($Name) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Name'}).VBoxValues = $Name
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Name'}).Options = $true
      }
      if ($PrimaryGroup) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'PrimaryGroup'}).VBoxValues = $PrimaryGroup
       ($appliancedescriptions | Where-Object {$_.Types -eq 'PrimaryGroup'}).Options = $true
      }
      if ($SettingsFile) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'SettingsFile'}).VBoxValues = $SettingsFile
       ($appliancedescriptions | Where-Object {$_.Types -eq 'SettingsFile'}).Options = $true
      }
      if ($BaseFolder) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'BaseFolder'}).VBoxValues = $BaseFolder
       ($appliancedescriptions | Where-Object {$_.Types -eq 'BaseFolder'}).Options = $true
      }
      if ($Description) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Description'}).VBoxValues = $Description
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Description'}).Options = $true
      }
      if ($Cpu) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CPU'}).VBoxValues = $Cpu
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CPU'}).Options = $true
      }
      if ($Memory) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Memory'}).VBoxValues = $Memory
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Memory'}).Options = $true
      }
      if ($UsbController) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'USBController'}).VBoxValues = $USBController
       ($appliancedescriptions | Where-Object {$_.Types -eq 'USBController'}).Options = $true
      }
      if ($NetworkAdapter) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'NetworkAdapter'}).VBoxValues = $NetworkAdapter
       ($appliancedescriptions | Where-Object {$_.Types -eq 'NetworkAdapter'}).Options = $true
      }
      if ($CdRom) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CDROM'}).VBoxValues = $CdRom
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CDROM'}).Options = $true
      }
      if ($HardDiskControllerScsi) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).VBoxValues = $HardDiskControllerScsi
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSCSI'}).Options = $true
      }
      if ($HardDiskControllerIde) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'}).VBoxValues = $HardDiskControllerIde
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerIDE'}).Options = $true
      }
      if ($HardDiskImage) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskImage'}).VBoxValues = $HardDiskImage
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskImage'}).Options = $true
      }
      if ($Product) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Product'}).VBoxValues = $Product
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Product'}).Options = $true
      }
      if ($Vendor) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Vendor'}).VBoxValues = $Vendor
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Vendor'}).Options = $true
      }
      if ($Version) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Version'}).VBoxValues = $Version
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Version'}).Options = $true
      }
      if ($ProductUrl) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'ProductUrl'}).VBoxValues = $ProductUrl
       ($appliancedescriptions | Where-Object {$_.Types -eq 'ProductUrl'}).Options = $true
      }
      if ($VendorUrl) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'VendorUrl'}).VBoxValues = $VendorUrl
       ($appliancedescriptions | Where-Object {$_.Types -eq 'VendorUrl'}).Options = $true
      }
      if ($License) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'License'}).VBoxValues = $License
       ($appliancedescriptions | Where-Object {$_.Types -eq 'License'}).Options = $true
      }
      if ($Miscellaneous) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Miscellaneous'}).VBoxValues = $Miscellaneous
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Miscellaneous'}).Options = $true
      }
      if ($HardDiskControllerSata) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSATA'}).VBoxValues = $HardDiskControllerSata
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSATA'}).Options = $true
      }
      if ($HardDiskControllerSas) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).VBoxValues = $HardDiskControllerSas
       ($appliancedescriptions | Where-Object {$_.Types -eq 'HardDiskControllerSAS'}).Options = $true
      }
      if ($Floppy) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Floppy'}).VBoxValues = $Floppy
       ($appliancedescriptions | Where-Object {$_.Types -eq 'Floppy'}).Options = $true
      }
      if ($SoundCard) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'SoundCard'}).VBoxValues = $SoundCard
       ($appliancedescriptions | Where-Object {$_.Types -eq 'SoundCard'}).Options = $true
      }
      if ($CloudInstanceShape) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceShape'}).VBoxValues = $CloudInstanceShape
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceShape'}).Options = $true
      }
      if ($CloudDomain) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudDomain'}).VBoxValues = $CloudDomain
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudDomain'}).Options = $true
      }
      if ($CloudBootDiskSize) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootDiskSize'}).VBoxValues = $CloudBootDiskSize
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootDiskSize'}).Options = $true
      }
      if ($CloudBucket) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBucket'}).VBoxValues = $CloudBucket
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBucket'}).Options = $true
      }
      if ($CloudOciVcn) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCN'}).VBoxValues = $CloudOciVcn
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCN'}).Options = $true
      }
      if ($CloudPublicIp) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicIP'}).VBoxValues = $CloudPublicIp
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicIP'}).Options = $true
      }
      if ($CloudProfileName) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudProfileName'}).VBoxValues = $CloudProfileName
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudProfileName'}).Options = $true
      }
      if ($CloudOciSubnet) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnet'}).VBoxValues = $CloudOciSubnet
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnet'}).Options = $true
      }
      if ($CloudKeepObject) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudKeepObject'}).VBoxValues = $CloudKeepObject
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudKeepObject'}).Options = $true
      }
      if ($CloudLaunchInstance) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudLaunchInstance'}).VBoxValues = $CloudLaunchInstance
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudLaunchInstance'}).Options = $true
      }
      if ($CloudImageState) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageState'}).VBoxValues = $CloudImageState
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageState'}).Options = $true
      }
      if ($CloudInstanceDisplayName) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceDisplayName'}).VBoxValues = $CloudInstanceDisplayName
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudInstanceDisplayName'}).Options = $true
      }
      if ($CloudImageDisplayName) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageDisplayName'}).VBoxValues = $CloudImageDisplayName
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudImageDisplayName'}).Options = $true
      }
      if ($CloudOciLaunchMode) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCILaunchMode'}).VBoxValues = $CloudOciLaunchMode
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCILaunchMode'}).Options = $true
      }
      if ($CloudPrivateIp) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPrivateIP'}).VBoxValues = $CloudPrivateIp
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPrivateIP'}).Options = $true
      }
      if ($CloudBootVolumeId) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootVolumeId'}).VBoxValues = $CloudBootVolumeId
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudBootVolumeId'}).Options = $true
      }
      if ($CloudOciVcnCompartment) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCNCompartment'}).VBoxValues = $CloudOciVcnCompartment
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCIVCNCompartment'}).Options = $true
      }
      if ($CloudOciSubnetCompartment) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnetCompartment'}).VBoxValues = $CloudOciSubnetCompartment
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudOCISubnetCompartment'}).Options = $true
      }
      if ($CloudPublicSshKey) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicSSHKey'}).VBoxValues = $CloudPublicSshKey
       ($appliancedescriptions | Where-Object {$_.Types -eq 'CloudPublicSSHKey'}).Options = $true
      }
      if ($BootingFirmware) {
       ($appliancedescriptions | Where-Object {$_.Types -eq 'BootingFirmware'}).VBoxValues = $BootingFirmware
       ($appliancedescriptions | Where-Object {$_.Types -eq 'BootingFirmware'}).Options = $true
      }
      Write-Verbose "Converting options to integers"
      foreach ($option in $appliancedescriptions.Options) {[int[]]$finalOptions += $option}
      Write-Verbose "Converting values to strings"
      foreach ($value in $appliancedescriptions.VBoxValues) {[string[]]$finalValues += $value}
      Write-Verbose "Converting extra config values to strings"
      foreach ($extraconfigvalue in $appliancedescriptions.ExtraConfigValues) {[string[]]$finalExtraConfigValues += $extraconfigvalue}
      Write-Verbose "Applying final settings to appliance"
      $ivirtualsystemdescription.SetFinalValues($finalOptions, $finalValues, $finalExtraConfigValues)
     } # foreach $ivirtualsystemdescription in $ivirtualsystemdescriptions
     # export the machine to disk
     Write-Verbose "Writing OVF to disk"
     $imachine.IProgress.Progress = $iappliance.Write($OvfFormat, [ExportOptions]::new().ToInt($ExportOptions), (Join-Path -ChildPath "$($imachine.Name).$($Ext)" -Path $FilePath))
     if ($ProgressBar) {Write-Progress -Activity "Exporting VM $($imachine.Name)" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1}
     do {
      # update iprogress data
      if ($ProgressBar) {
       if ($imachine.IProgress.Progress.Percent -lt 20) {
        Write-Progress -Activity "Exporting VM $($imachine.Name)" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1
        if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
       } # end if $imachine.IProgress.Progress.Percent < 20
       else {
        Write-Progress -Activity "Exporting VM $($imachine.Name)" -status "$($imachine.IProgress.Progress.Description): $($imachine.IProgress.Progress.Percent)%" -percentComplete ($imachine.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imachine.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imachine.IProgress.Progress.TimeRemaining)
        if ($imachine.IProgress.Progress.OperationDescription) {Write-Progress -Activity "$($imachine.IProgress.Progress.OperationDescription)" -status "$($imachine.IProgress.Progress.OperationDescription): $($imachine.IProgress.Progress.OperationPercent)%" -percentComplete ($imachine.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
       } # end else
      } # end if $ProgressBar
     } until ($imachine.IProgress.Progress.Percent -eq 100 -and $imachine.IProgress.Progress.Completed -eq $true) # continue once completed
     if ($imachine.IProgress.Progress.ResultCode -ne 0) {Write-Verbose $imachine.IProgress.Progress.ErrorInfo}
    } # end elseif com
   } # end foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Host "[Error] No matching virtual machines were found using specified parameters" -ForegroundColor Red -BackgroundColor Black;return}
 } # Try
 catch {
  Write-Verbose 'Exception exporting OVF'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine objects'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.IPercent) {$imachine.IPercent = $null}
   } # end foreach $imachine in $imachines
  } # end if $imachines
  if ($ivirtualsystemdescriptions -and $ModuleHost.ToLower() -eq 'websrv') {
   foreach ($ivirtualsystemdescription in $ivirtualsystemdescriptions) {
    $global:vbox.IManagedObjectRef_release($ivirtualsystemdescription)
   }
  }
  if ($iappliance -and $ModuleHost.ToLower() -eq 'websrv') {
   $global:vbox.IManagedObjectRef_release($iappliance)
  }
  $ivirtualsystemdescriptions = $null
  $iappliance = $null
 } # Finally
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Get-VirtualBoxDisk {
<#
.SYNOPSIS
Get VirtualBox disk information
.DESCRIPTION
Retrieve VirtualBox disks by name, format, GUID, machine object, machine name, machine GUID, or all.
.PARAMETER Name
At least one virtual disk name.
.PARAMETER Format
At least one virtual disk format.
.PARAMETER Guid
At least one virtual disk GUID.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER MachineName
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER MachineGuid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Get-VirtualBoxDisk -Name 2016

Name        : 2016 Core.vhd
Description :
Format      : VHD
Size        : 7291584512
LogicalSize : 53687091200
MachineGuid : {7353caa6-8cb6-4066-aec9-6c6a69a001b6}
MachineName : {2016 Core}

Get virtual disk by name
.EXAMPLE
PS C:\> Get-VirtualBoxVM -Name 2016 | Get-VirtualBoxDisk

Name        : 2016 Core.vhd
Description :
Format      : VHD
Size        : 7291584512
LogicalSize : 53687091200
MachineGuid : {7353caa6-8cb6-4066-aec9-6c6a69a001b6}
MachineName : {2016 Core}

Get virtual disk by machine object from pipeline input
.EXAMPLE
PS C:\> Get-VirtualBoxDisk -MachineName 2016

Name        : 2016 Core.vhd
Description :
Format      : VHD
Size        : 7291584512
LogicalSize : 53687091200
MachineGuid : {7353caa6-8cb6-4066-aec9-6c6a69a001b6}
MachineName : {2016 Core}

Get virtual disk by machine name
.EXAMPLE
PS C:\> Get-VirtualBoxDisk -MachineGuid 7353caa6-8cb6-4066-aec9-6c6a69a001b6

Name        : 2016 Core.vhd
Description :
Format      : VHD
Size        : 7291584512
LogicalSize : 53687091200
MachineGuid : {7353caa6-8cb6-4066-aec9-6c6a69a001b6}
MachineName : {2016 Core}

Get virtual disk by machine GUID
.EXAMPLE
PS C:\> Get-VirtualBoxDisk

Name        : GNS3 IOU VM_1.3-disk1.vmdk
Description :
Format      : VMDK
Size        : 1242759168
LogicalSize : 2147483648
MachineGuid : {c9d4dc35-3967-4009-993d-1c23ab4ff22b}
MachineName : {GNS3 IOU VM_1.3}

Name        : turnkey-lamp-disk1.vdi
Description :
Format      : vdi
Size        : 4026531840
LogicalSize : 21474836480
MachineGuid : {a237e4f5-da5a-4fca-b2a6-80f9aea91a9b}
MachineName : {WebSite}

Name        : 2016 Core.vhd
Description :
Format      : VHD
Size        : 7291584512
LogicalSize : 53687091200
MachineGuid : {7353caa6-8cb6-4066-aec9-6c6a69a001b6}
MachineName : {2016 Core}

Name        : Win10.vhd
Description :
Format      : VHD
Size        : 15747268096
LogicalSize : 53687091200
MachineGuid : {15a4c311-3b89-4936-89c7-11d3340ced7a}
MachineName : {Win10}

Get all virtual machine disks in the VirtualBox inventory
.NOTES
NAME        :  Get-VirtualBoxDisk
VERSION     :  1.3
LAST UPDATED:  1/8/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
New-VirtualBoxDisk
.INPUTS
String[]      :  Strings for virtual disk names
String[]      :  Strings for virtual disk formats
Guid[]        :  GUIDs for virtual disk GUIDs
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
.OUTPUTS
VirtualBoxVHD[]
#>
[cmdletbinding(DefaultParameterSetName="Machine")]
Param(
[Parameter(HelpMessage="Enter one or more disk name(s)",
ParameterSetName="Disk",Mandatory=$false,Position=0)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(HelpMessage="Enter one or more disk format(s)",
ParameterSetName="Disk",Mandatory=$false,Position=0)]
[ValidateNotNullorEmpty()]
  [string[]]$Format,
[Parameter(HelpMessage="Enter one or more disk GUID(s)",
ParameterSetName="Disk",Mandatory=$false,Position=0)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$false,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Machine",Mandatory=$false,Position=0)]
  [string[]]$MachineName,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Machine",Mandatory=$false,Position=0)]
  [guid[]]$MachineGuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Format: `"$Format`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - MachineName: `"$MachineName`""
 Write-Verbose "Pipeline - MachineGuid: `"$MachineGuid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 $disks = @()
 $obj = @()
 try {
  # get virtual machine disk inventory
  Write-Verbose "Getting virtual disk inventory"
  if ($ModuleHost.ToLower() -eq 'websrv') {
   foreach ($imediumid in $global:vbox.IVirtualBox_getHardDisks($global:ivbox)) {
    Write-Verbose "Getting disk: $($imediumid)"
    $disk = New-Object VirtualBoxVHD
    $disk.Name = $global:vbox.IMedium_getName($imediumid)
    $disk.GUID = $global:vbox.IMedium_getId($imediumid)
    $disk.Description = $global:vbox.IMedium_getDescription($imediumid)
    $disk.Format = $global:vbox.IMedium_getFormat($imediumid)
    $disk.Size = $global:vbox.IMedium_getSize($imediumid)
    $disk.LogicalSize = $global:vbox.IMedium_getLogicalSize($imediumid)
    $disk.MachineGuid = $global:vbox.IMedium_getMachineIds($imediumid)
    foreach ($machineid in $disk.MachineGuid) {
     foreach ($imachine in ($global:vbox.IVirtualBox_getMachines($global:ivbox))) {
      if (($global:vbox.IMachine_getId($imachine)) -eq $machineid) {
       $disk.MachineName += $global:vbox.IMachine_getName($imachine)
      } # end if $imachine.Guid -eq $machineid
      $disk.MachineName = $disk.MachineName | Where-Object {$_ -ne $null}
     } # foreach $imachine in $imachines
    } # foreach $machineid in $disk.MachineGuid
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
    [VirtualBoxVHD[]]$disks += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;}
   } # end foreach loop inventory
  } # end if websrv
  elseif ($ModuleHost.ToLower() -eq 'com') {
   foreach ($imedium in $vbox.HardDisks) {
    Write-Verbose "Getting disk: $($imedium.Id)"
    $disk = New-Object VirtualBoxVHD
    $disk.Name = $imedium.Name
    $disk.Guid = $imedium.Id
    $disk.Description = $imedium.Description
    $disk.Format = $imedium.Format
    $disk.Size = $imedium.Size
    $disk.LogicalSize = $imedium.LogicalSize
    $disk.MachineGuid = $imedium.MachineIds
    foreach ($machineid in $disk.MachineGuid) {
     foreach ($imachine in $global:vbox.Machines) {
      if ($imachine.Id -eq $machineid) {
       $disk.MachineName += $imachine.Name
      } # end if $imachine.Guid -eq $machineid
      $disk.MachineName = $disk.MachineName | Where-Object {$_ -ne $null}
     } # foreach $imachine in $imachines
    } # foreach $machineid in $disk.MachineGuid
    $disk.State = [MediumState]::new().ToStr($imedium.State)
    $disk.Variant = @()
    foreach ($variant in $imedium.Variant) {
     $disk.Variant += [MediumVariant]::new().ToStrCom($variant)
    }
    $disk.Variant = $disk.Variant | Where-Object {$_ -ne $null}
    $disk.Location = $imedium.Location
    $disk.HostDrive = $imedium.HostDrive
    $disk.Type = [MediumType]::new().ToStr($imedium.Type)
    if ($imedium.Parent) {$disk.Parent = $imedium.Parent.Name}
    if ($imedium.Children) {$disk.Children = $imedium.Children.Name}
    $disk.ComObject = $imedium
    $disk.ReadOnly = $imedium.ReadOnly
    $disk.AutoReset = $imedium.AutoReset
    $disk.LastAccessError = $imedium.LastAccessError
    [VirtualBoxVHD[]]$disks += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;ComObject=$disk.ComObject}
   } # end foreach loop inventory
  } # end elseif com
 } # Try
 catch {
  Write-Verbose 'Exception retrieving virtual disk information'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 if ($PSCmdlet.ParameterSetName -eq "Disk") {
  # filter by disk name
  if ($Name -and $Name -ne '*') {
   foreach ($item in $Name) {
    foreach ($disk in $disks) {
     $matched = $false
     Write-Verbose "Matching $($disk.Name) to $item"
     if ($disk.Name -match $item) {Write-Verbose "Matched $($disk.Name) to $item";$matched = $true}
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;ComObject=$disk.ComObject}}
    } # foreach $disk in $disks
   } # foreach $item in $Name
  } # end if $Name
  # filter by disk format
  elseif ($Format -and $Format -ne '*') {
   foreach ($item in $Format) {
    foreach ($disk in $disks) {
     $matched = $false
     Write-Verbose "Matching $($disk.Format) to $item"
     if ($disk.Format -match $item) {Write-Verbose "Matched $($disk.Format) to $item";$matched = $true}
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;ComObject=$disk.ComObject}}
    } # foreach $disk in $disks
   } # foreach $item in $Format
   $obj = $obj | Where-Object {$_ -ne $null}
  } # end if $Format
  # filter by disk guid
  elseif ($Guid) {
   foreach ($item in $Guid) {
    foreach ($disk in $disks) {
     $matched = $false
     Write-Verbose "Matching $($disk.Guid) to $item"
     if ($disk.Guid -match $item) {Write-Verbose "Matched $($disk.Guid) to $item";$matched = $true}
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;ComObject=$disk.ComObject}}
    } # foreach $disk in $disks
   } # foreach $item in $Guid
  } # end if $Guid
  # no filter
  else {foreach ($disk in $disks) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;ComObject=$disk.ComObject}}}
  Write-Verbose "Found $(($obj | Measure-Object).count) disk(s)"
 }
 elseif ($PSCmdlet.ParameterSetName -eq "Machine") {
  # filter by machine object
  if ($Machine) {
   foreach ($item in $Machine) {
    foreach ($disk in $disks) {
     $matched = $false
     foreach ($vmname in $disk.MachineName) {
      Write-Verbose "Matching $vmname to $($item.Name)"
      if ($vmname -match $item.Name) {Write-Verbose "Matched $vmname to $($item.Name)";$matched = $true}
     } # foreach $vmname in $disk.MachineName
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;ComObject=$disk.ComObject}}
    } # foreach $disk in $disks
   } # foreach $item in $Machine
  } # end if $Machine
  # filter by machine name
  elseif ($MachineName) {
   foreach ($item in $MachineName) {
    foreach ($disk in $disks) {
     $matched = $false
     foreach ($vmname in $disk.MachineName) {
      Write-Verbose "Matching $vmname to $item"
      if ($vmname -match $item) {Write-Verbose "Matched $vmname to $item";$matched = $true}
     } # foreach $vmname in $disk.MachineName
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;ComObject=$disk.ComObject}}
    } # foreach $disk in $disks
   } # foreach $item in $MachineName
  } # end elseif $MachineName
  # filter by machine GUID
  elseif ($MachineGuid) {
   foreach ($item in $MachineGuid) {
    foreach ($disk in $disks) {
     $matched = $false
     foreach ($vmguid in $disk.MachineGuid) {
      Write-Verbose "Matching $vmguid to $item"
      if ($vmguid -eq $item) {Write-Verbose "Matched $vmguid to $item";$matched = $true}
     } # foreach $vmguid in $disk.MachineGuid
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;ComObject=$disk.ComObject}}
    } # foreach $disk in $disks
   } # foreach $item in $MachineGuid
  } # end elseif $MachineGuid
  # no filter
  else {foreach ($disk in $disks) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disk.Name;Guid=$disk.Guid;Description=$disk.Description;Format=$disk.Format;Size=$disk.Size;LogicalSize=$disk.LogicalSize;MachineGuid=$disk.MachineGuid;MachineName=$disk.MachineName;State=$disk.State;Variant=$disk.Variant;Location=$disk.Location;HostDrive=$disk.HostDrive;MediumFormat=$disk.MediumFormat;Type=$disk.Type;Parent=$disk.Parent;Children=$disk.Children;Id=$disk.Id;ReadOnly=$disk.ReadOnly;AutoReset=$disk.AutoReset;LastAccessError=$disk.LastAccessError;ComObject=$disk.ComObject}}}
  Write-Verbose "Found $(($obj | Measure-Object).count) disk(s)"
 }
 if ($obj) {
  # write virtual machines object to the pipeline as an array
  Write-Output ([System.Array]$obj)
 } # end if $obj
 else {
  Write-Verbose "[Warning] No matching virtual disks found."
 } # end else
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function New-VirtualBoxDisk {
<#
.SYNOPSIS
Create VirtualBox disk
.DESCRIPTION
Creates VirtualBox disks. The command will fail if a virtual disk with the same name exists in the VirtualBox inventory.
.PARAMETER Name
The virtual disk name.
.PARAMETER Format
The virtual disk format.
.PARAMETER Location
The location to store the virtual disk. If the path does not exist it will be created.
.PARAMETER AccessMode
Either Readonly or ReadWrite.
.PARAMETER LogicalSize
The size of the virtual disk in bytes.
.PARAMETER VariantType
The variant type of the virtual disk.
.PARAMETER VariantFlag
The variant flag of the virtual disk.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> New-VirtualBoxDisk -AccessMode ReadWrite -Format VMDK -Location C:\Disks -LogicalSize 4194304 -Name TestDisk -VariantFlag Fixed -VariantType Standard -ProgressBar

Create a standard, fixed 4MB virtual disk named "TestDisk.vmdk" in the C:\Disks\ location and display a progress bar
.NOTES
NAME        :  New-VirtualBoxDisk
VERSION     :  1.0
LAST UPDATED:  1/16/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisk
.INPUTS
String        :  String for virtual disk name
String        :  String for virtual disk format
String        :  String for virtual disk location
String        :  String for virtual disk access mode
UInt64        :  UInt64 for virtual disk size
String        :  String for virtual disk variant type
String        :  String for virtual disk variant flag
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(HelpMessage="Enter the virtual disk name",
Mandatory=$true,Position=0)]
[ValidateNotNullorEmpty()]
  [string]$Name,
[Parameter(HelpMessage="Enter the virtual disk format",
Mandatory=$true,Position=1)]
[ValidateNotNullorEmpty()]
[ValidateSet('VMDK','VDI','VHD','Parallels','DMG','QED','QCOW','VHDX','CUE','VBoxIsoMaker','RAW','iSCSI')]
  [string]$Format,
[Parameter(HelpMessage="Enter the virtual disk location",
Mandatory=$true,Position=2)]
  [string]$Location,
[Parameter(HelpMessage="Enter the virtual disk location",
Mandatory=$true,Position=3)]
[ValidateSet('ReadOnly','ReadWrite')]
  [string]$AccessMode,
[Parameter(HelpMessage="Enter the logical size of the virtual disk in bytes",
Mandatory=$true,Position=4)]
  [uint64]$LogicalSize,
[Parameter(HelpMessage="Enter the virtual disk variant type",
Mandatory=$true,Position=5)]
[ValidateSet('Standard','VmdkSplit2G','VmdkRawDisk','VmdkStreamOptimized','VmdkESX','VdiZeroExpand')]
  [string]$VariantType,
[Parameter(HelpMessage="Enter the virtual disk variant flag",
Mandatory=$false,Position=5)]
[ValidateSet('Fixed','Diff','Formatted','NoCreateDir')]
  [string]$VariantFlag,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 if (!(Test-Path $Location)) {
  # create the directory if it doesn't exist
  Write-Verbose "Creating $($Location) directory"
  New-Item -ItemType Directory -Path $Location -Force -Confirm:$false | Write-Verbose
 }
 $existingdisks = Get-VirtualBoxDisk -Name $Name -SkipCheck
 if ($existingdisks) {
  foreach ($existingdisk in $existingdisks) {
   Write-Verbose $existingdisk.Name
   if ($existingdisk.Name -eq "$Name.$Ext") {
    Write-Host "[Error] Hard disk $Name.$Ext already exists. Select another name or format and try again." -ForegroundColor Red -BackgroundColor Black
    return
   }
  }
 }
 try {
  Write-Verbose "Creating virtual disk object"
  $imedium = New-Object VirtualBoxVHD
  if ($ModuleHost.ToLower() -eq 'websrv') {
   Write-Verbose "Creating medium"
   Write-Verbose "Path: `"$(Join-Path -ChildPath "$Name.$Ext" -Path $Location)`""
   Write-Verbose "AccessMode: `"$($AccessMode)`" ($([AccessMode]::new().ToULong($AccessMode)))"
   Write-Verbose "DeviceType: `"$('HardDisk')`" ($([DeviceType]::new().ToULong('HardDisk')))"
   $imedium.Id = $global:vbox.IVirtualBox_createMedium($global:ivbox, $Format, (Join-Path -ChildPath "$Name.$Ext" -Path $Location), [AccessMode]::new().ToULong($AccessMode), [DeviceType]::new().ToULong('HardDisk'))
   Write-Verbose "Creating base storage"
   Write-Verbose "LogicalSize: `"$($LogicalSize)`""
   Write-Verbose "VariantType: `"$($VariantType)`" ($([MediumVariant]::new().ToInt($VariantType)))"
   Write-Verbose "VariantFlag: `"$($VariantFlag)`" ($([MediumVariant]::new().ToInt($VariantFlag)))"
   $imedium.IProgress.Id = $global:vbox.IMedium_createBaseStorage($imedium.Id, $LogicalSize, @([MediumVariant]::new().ToInt($VariantType), [MediumVariant]::new().ToInt($VariantFlag)))
   # collect iprogress data
   Write-Verbose "Fetching IProgress data"
   $imedium.IProgress = $imedium.IProgress.Fetch($imedium.IProgress.Id)
   if ($ProgressBar) {Write-Progress -Activity "Creating virtual disk $($imedium.Name)" -status "$($imedium.IProgress.Description): $($imedium.IProgress.Percent)%" -percentComplete ($imedium.IProgress.Percent) -CurrentOperation "Current Operation: $($imedium.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imedium.IProgress.TimeRemaining)}
   do {
    $mediumstate = $global:vbox.IMedium_getState($imedium.Id)
    # update iprogress data
    $imedium.IProgress = $imedium.IProgress.Update($imedium.IProgress.Id)
    if ($ProgressBar) {Write-Progress -Activity "Creating virtual disk $($imedium.Name)" -status "$($imedium.IProgress.Description): $($imedium.IProgress.Percent)%" -percentComplete ($imedium.IProgress.Percent) -CurrentOperation "Current Operation: $($imedium.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imedium.IProgress.TimeRemaining)}
    if ($ProgressBar) {Write-Progress -Activity "$($imedium.IProgress.OperationDescription)" -status "$($imedium.IProgress.OperationDescription): $($imedium.IProgress.OperationPercent)%" -percentComplete ($imedium.IProgress.OperationPercent) -Id 2 -ParentId 1}
   } until ($imedium.IProgress.Percent -eq 100 -or $mediumstate -match 'NotCreated') # continue once the progress reached 100%
   if ($mediumstate -match 'NotCreated') {Write-Host "[Error] Failed to create base storage" -ForegroundColor Red -BackgroundColor Black}
  } # end if websrv
  elseif ($ModuleHost.ToLower() -eq 'com') {
   Write-Verbose "Creating medium"
   Write-Verbose "Path: `"$(Join-Path -ChildPath "$Name.$Ext" -Path $Location)`""
   Write-Verbose "Format: `"$($Format)`""
   Write-Verbose "AccessMode: `"$($AccessMode)`" ($([AccessMode]::new().ToULong($AccessMode)))"
   Write-Verbose "DeviceType: `"$('HardDisk')`" ($([DeviceType]::new().ToULong('HardDisk')))"
   $newdisk = $global:vbox.CreateMedium($Format, (Join-Path -ChildPath "$Name.$Ext" -Path $Location), [AccessMode]::new().ToULong($AccessMode), [DeviceType]::new().ToULong('HardDisk'))
   $imedium.ComObject = $newdisk
   Write-Verbose "Creating base storage"
   Write-Verbose "LogicalSize: `"$($LogicalSize)`""
   Write-Verbose "VariantType: `"$($VariantType)`" ($([MediumVariant]::new().ToInt($VariantType)))"
   Write-Verbose "VariantFlag: `"$($VariantFlag)`" ($([MediumVariant]::new().ToInt($VariantFlag)))"
   $imedium.IProgress.Progress = $newdisk.CreateBaseStorage($LogicalSize, [int[]]@([MediumVariant]::new().ToInt($VariantType), [MediumVariant]::new().ToInt($VariantFlag)))
   if ($ProgressBar) {Write-Progress -Activity "Creating virtual disk $($imedium.Name)" -status "$($imedium.IProgress.Progress.Description): $($imedium.IProgress.Progress.Percent)%" -percentComplete ($imedium.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imedium.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imedium.IProgress.Progress.TimeRemaining)}
   do {
    # update iprogress data
    if ($ProgressBar) {Write-Progress -Activity "Creating virtual disk $($imedium.Name)" -status "$($imedium.IProgress.Progress.Description): $($imedium.IProgress.Progress.Percent)%" -percentComplete ($imedium.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imedium.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imedium.IProgress.Progress.TimeRemaining)}
    if ($ProgressBar) {Write-Progress -Activity "$($imedium.IProgress.Progress.OperationDescription)" -status "$($imedium.IProgress.Progress.OperationDescription): $($imedium.IProgress.Progress.OperationPercent)%" -percentComplete ($imedium.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
   } until ($imedium.IProgress.Progress.Percent -eq 100 -or $newdisk.State -eq 0) # continue once the progress reached 100%
   if ($newdisk.State -eq 0) {Write-Host "[Error] Failed to create base storage" -ForegroundColor Red -BackgroundColor Black}
  } # end elseif com
 } # Try
 catch {
  Write-Verbose 'Exception creating virtual disk'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Import-VirtualBoxDisk {
<#
.SYNOPSIS
Import VirtualBox disk
.DESCRIPTION
Imports VirtualBox disks. The command will fail if a virtual disk with the same name exists in the VirtualBox inventory.
.PARAMETER FileName
The full path to the virtual disk file.
.PARAMETER AccessMode
Either Readonly or ReadWrite.
.PARAMETER LogicalSize
A switch to request a new disk UUID be created.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Import-VirtualBoxDisk -FileName C:\Disks\TestDisk.vmdk -AccessMode ReadWrite

Import the "C:\Disks\TestDisk.vmdk" disk in Read/Write mode to the VirtualBox inventory
.NOTES
NAME        :  Import-VirtualBoxDisk
VERSION     :  1.1
LAST UPDATED:  2/8/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisk
.INPUTS
String        :  String for virtual disk file path
String        :  String for virtual disk access mode
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(HelpMessage="Enter the full virtual disk path",
Mandatory=$true,Position=2)]
[ValidateScript({Test-Path $_})]
  [string]$FileName,
[Parameter(HelpMessage="Enter the virtual disk access type",
Mandatory=$true,Position=3)]
[ValidateSet('ReadOnly','ReadWrite')]
  [string]$AccessMode,
[Parameter(HelpMessage="Use this switch to request a new disk UUID be created",
Mandatory=$false)]
  [switch]$ForceNewUuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 $existingdisks = Get-VirtualBoxDisk -Name $($FileName.Substring($FileName.LastIndexOf('\')+1)) -SkipCheck
 if ($existingdisks) {
  foreach ($existingdisk in $existingdisks) {
   Write-Verbose $existingdisk.Name
   if ($existingdisk.Name -eq ($FileName.Substring($FileName.LastIndexOf('\')+1))) {
    Write-Host "[Error] Hard disk $($existingdisk.Name) already exists. Select another disk image and try again." -ForegroundColor Red -BackgroundColor Black
    return
   }
  }
 }
 try {
  $imedium = New-Object VirtualBoxVHD
  if ($ModuleHost.ToLower() -eq 'websrv') {
   $imedium.Id = $global:vbox.IVirtualBox_openMedium($global:ivbox, $FileName, [DeviceType]::new().ToULong('HardDisk'), [AccessMode]::new().ToULong($AccessMode), $(if ($ForceNewUuid) {$true} else {$false}))
  } # end if websrv
  elseif ($ModuleHost.ToLower() -eq 'com') {
   $imedium.ComObject = $global:vbox.OpenMedium($FileName, [DeviceType]::new().ToULong('HardDisk'), [AccessMode]::new().ToULong($AccessMode), $(if ($ForceNewUuid) {1} else {0}))
  } # end elseif com
 } # Try
 catch {
  Write-Verbose 'Exception importing virtual disk'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Remove-VirtualBoxDisk {
<#
.SYNOPSIS
Remove VirtualBox disk
.DESCRIPTION
Removes VirtualBox disks. The command will fail if a virtual disk does not exist in the VirtualBox inventory or if the disk is mounted to a machine.
.PARAMETER Disk
At least one virtual disk object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual disk. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual disk. Can be received via pipeline input by name.
.PARAMETER DeleteFromHost
A switch to delete the virtual disk from the host. This cannot be undone.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Remove-VirtualBoxDisk -Name TestDisk.vmdk -DeleteFromHost -ProgressBar -Confirm:$false

Remove the virtual disk named "TestDisk.vmdk" from the VirtualBox inventory, delete it from host, do not confirm the action, and display a progress bar
.NOTES
NAME        :  Remove-VirtualBoxDisk
VERSION     :  1.0
LAST UPDATED:  1/20/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisk
.INPUTS
VirtualBoxVHD[]:  VirtualBoxVHDs for virtual disk objects
String[]       :  Strings for virtual disk names
GUID[]         :  GUIDS for virtual disk GUIDS
.OUTPUTS
None
#>
[cmdletbinding(SupportsShouldProcess,ConfirmImpact='High')]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk object(s)",
Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVHD[]]$Disk,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Use this switch to delete the virtual disk from the host")]
  [switch]$DeleteFromHost,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 Write-Verbose "Pipeline - Disk: `"$Disk`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 Write-Verbose "Controller Name: `"$Controller`""
 Write-Verbose "Controller Port: `"$ControllerPort`""
 Write-Verbose "Controller Slot: `"$ControllerSlot`""
 if (!($Disk -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one disk object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imediums = @()
 if ($Disk) {
  Write-Verbose "Getting disk inventory from Disk(s) object"
  $imediums = $Disk
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting disk inventory from Name(s)"
   $imediums += Get-VirtualBoxDisk -Name $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting disk inventory from GUID(s)"
   $imediums += Get-VirtualBoxDisk -Guid $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 $imediums = $imediums | Where-Object {$_.Format -notmatch 'RAW'}
 if ($imediums) {
  Write-Verbose "[Info] Found disks"
  try {
   foreach ($imedium in $imediums) {
    Write-Verbose "Found disk: $($imedium.Name)"
    if ($imedium.MachineName) {
     foreach ($vmname in $imedium.MachineName) {
     Write-Verbose "Disk attached to VM: $vmname"
      Write-Host "[Error] The disk $($imedium.Name) is still mounted to machine $vmname. Dismount the disk from the machine and try again." -ForegroundColor Red -BackgroundColor Black;return
     } # foreach $vmname in $imedium.MachineName
    } # end if $imedium.MachineName
    if ($DeleteFromHost) {
     if ($PSCmdlet.ShouldProcess("$($imedium.Name)" , "Delete storage medium from host ")) {
      if ($ModuleHost.ToLower() -eq 'websrv') {
       # delete disk from host
       Write-Verbose "Removing disk $($imedium.Name)"
       $imedium.IProgress.Id = $global:vbox.IMedium_deleteStorage($imedium.Id)
       # collect iprogress data
       Write-Verbose "Fetching IProgress data"
       $imedium.IProgress = $imedium.IProgress.Fetch($imedium.IProgress.Id)
       if ($ProgressBar) {Write-Progress -Activity "Removing disk $($imedium.Name) from host" -status "$($imedium.IProgress.Description): $($imedium.IProgress.Percent)%" -percentComplete ($imedium.IProgress.Percent) -CurrentOperation "Current Operation: $($imedium.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imedium.IProgress.TimeRemaining)}
       do {
        # update iprogress data
        $imedium.IProgress = $imedium.IProgress.Update($imedium.IProgress.Id)
        if ($ProgressBar) {Write-Progress -Activity "Removing disk $($imedium.Name) from host" -status "$($imedium.IProgress.Description): $($imedium.IProgress.Percent)%" -percentComplete ($imedium.IProgress.Percent) -CurrentOperation "Current Operation: $($imedium.IProgress.OperationDescription)" -Id 1 -SecondsRemaining ($imedium.IProgress.TimeRemaining)}
        if ($ProgressBar) {Write-Progress -Activity "$($imedium.IProgress.OperationDescription)" -status "$($imedium.IProgress.OperationDescription): $($imedium.IProgress.OperationPercent)%" -percentComplete ($imedium.IProgress.OperationPercent) -Id 2 -ParentId 1}
       } until ($imedium.IProgress.Percent -eq 100) # continue once the progress reaches 100%
      } # end if websrv
      elseif ($ModuleHost.ToLower() -eq 'com') {
       # delete disk from host
       Write-Verbose "Removing disk $($imedium.Name)"
       $imedium.IProgress.Progress = $imedium.ComObject.DeleteStorage()
       if ($ProgressBar) {Write-Progress -Activity "Removing disk $($imedium.Name) from host" -status "$($imedium.IProgress.Progress.Description): $($imedium.IProgress.Progress.Percent)%" -percentComplete ($imedium.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imedium.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imedium.IProgress.Progress.TimeRemaining)}
       do {
        # update iprogress data
        if ($ProgressBar) {Write-Progress -Activity "Removing disk $($imedium.Name) from host" -status "$($imedium.IProgress.Progress.Description): $($imedium.IProgress.Progress.Percent)%" -percentComplete ($imedium.IProgress.Progress.Percent) -CurrentOperation "Current Operation: $($imedium.IProgress.Progress.OperationDescription)" -Id 1 -SecondsRemaining ($imedium.IProgress.Progress.TimeRemaining)}
        if ($ProgressBar) {Write-Progress -Activity "$($imedium.IProgress.Progress.OperationDescription)" -status "$($imedium.IProgress.Progress.OperationDescription): $($imedium.IProgress.Progress.OperationPercent)%" -percentComplete ($imedium.IProgress.Progress.OperationPercent) -Id 2 -ParentId 1}
       } until ($imedium.IProgress.Progress.Percent -eq 100) # continue once the progress reaches 100%
      } # end elseif com
     } # end if $PSCmdlet.ShouldProcess(
     else {Write-Verbose "Operation cancelled by user";return}
    } # end if $DeleteFromHost
    else {
     # close the disk
     Write-Verbose "Removing disk $($imedium.Name) from VirtualBox inventory"
     if ($ModuleHost.ToLower() -eq 'websrv') {
      $global:vbox.IMedium_close($imedium.Id)
     } # end if websrv
     elseif ($ModuleHost.ToLower() -eq 'com') {
      $imedium.ComObject.Close()
     } # end elseif com
    }
   } # foreach $imedium in $imediums
  } # Try
  catch {
   Write-Verbose 'Exception removing virtual disk'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  finally {
   # cleanup
  } # Finally
 } # end if $imediums
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Mount-VirtualBoxDisk {
<#
.SYNOPSIS
Mount VirtualBox disk
.DESCRIPTION
Mounts VirtualBox disks. The command will fail if a virtual disk is already mounted to the specified virtual machine.
.PARAMETER Disk
At least one virtual disk object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual disk. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual disk. Can be received via pipeline input by name.
.PARAMETER MachineName
The name of the virtual machine to mount the disk to. This is a required parameter.
.PARAMETER Controller
The name of the storage controller to mount the disk to. This is a required parameter.
.PARAMETER ControllerPort
The port of the storage controller to mount the disk to. This is a required parameter.
.PARAMETER ControllerSlot
The slot of the storage controller to mount the disk to. This is a required parameter.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Mount-VirtualBoxDisk -Name TestDisk -MachineName Win10 -Controller SATA -ControllerPort 0 -ControllerSlot 0

Mounts the virtual disk named "TestDisk.vmdk" to the Win10 virtual machine SATA controller on port 0 slot 0
.NOTES
NAME        :  Mount-VirtualBoxDisk
VERSION     :  1.0
LAST UPDATED:  1/20/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisk
.INPUTS
VirtualBoxVHD[]:  VirtualBoxVHDs for virtual disk objects
String[]       :  Strings for virtual disk names
GUID[]         :  GUIDS for virtual disk GUIDS
String         :  String for virtual machine name
String         :  String for controller name
Int            :  Integer for controller port
Int            :  Integer for controller slot
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk object(s)",
Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVHD[]]$Disk,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(Mandatory=$false,HelpMessage="Enter a virtual machine object to dismount the disk from")]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(Mandatory=$false,HelpMessage="Enter the name of the virtual machine to dismount the disk from")]
[ValidateNotNullorEmpty()]
  [string]$MachineName,
[Parameter(Mandatory=$false,HelpMessage="Enter the GUID of the virtual machine to dismount the disk from")]
[ValidateNotNullorEmpty()]
  [guid]$MachineGuid,
[Parameter(Mandatory=$true,HelpMessage="Enter the name of the controller to mount the disk to",
ParameterSetName="HardDisk")]
[ValidateNotNullorEmpty()]
  [string]$Controller,
[Parameter(Mandatory=$true,HelpMessage="Enter the port number to mount the disk to",
ParameterSetName="HardDisk")]
[ValidateNotNullorEmpty()]
  [int]$ControllerPort,
[Parameter(Mandatory=$true,HelpMessage="Enter the slot number to mount the disk to",
ParameterSetName="HardDisk")]
[ValidateNotNullorEmpty()]
  [int]$ControllerSlot,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 Write-Verbose "Pipeline - Disk: `"$Disk`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 Write-Verbose "Machine: `"$Machine`""
 Write-Verbose "Machine Name: `"$MachineName`""
 Write-Verbose "Machine Guid: `"$MachineGuid`""
 Write-Verbose "Controller Name: `"$Controller`""
 Write-Verbose "Controller Port: `"$ControllerPort`""
 Write-Verbose "Controller Slot: `"$ControllerSlot`""
 if (!($Disk -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one disk object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 if (!($Machine -or $MachineName -or $MachineGuid)) {Write-Host "[Error] You must supply at least one machine object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imediums = @()
 if ($Disk) {
  Write-Verbose "Getting disk inventory from Disk(s) object"
  $imediums = $Disk
  $imediums = $imediums | Where-Object {$_ -ne $null}
 } # get disk inventory (by $Disk)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting disk inventory from Name(s)"
   $imediums += Get-VirtualBoxDisk -Name $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 } # get disk inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting disk inventory from GUID(s)"
   $imediums += Get-VirtualBoxDisk -Guid $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 } # get vm inventory (by $Guid)
 $imediums = $imediums | Where-Object {$_.Format -notmatch 'RAW'}
 if ($imediums) {
  Write-Verbose "[Info] Found disks"
  try {
   foreach ($imedium in $imediums) {
    Write-Verbose "Found disk: $($imedium.Name)"
    if ($imedium.MachineName) {
     # make sure it's not already attached to the requested vm
     foreach ($vmname in $imedium.MachineName) {
      Write-Verbose "Disk attached to VM: $vmname"
      if (Get-VirtualBoxDisk -MachineName $vmname -SkipCheck) {Write-Host "[Error] The disk $($imedium.Name) is already mounted to machine $($imachine.Name)." -ForegroundColor Red -BackgroundColor Black;return}
     } # foreach $vmname in $imedium.MachineName
    } # end if $imedium.MachineName
    if ($Machine) {
     $imachines = $Machine
    } # get vm inventory (by $Machine)
    elseif ($MachineName) {
     $imachines = Get-VirtualBoxVM -Name $MachineName -SkipCheck
    } # get vm inventory (by $MachineName)
    elseif ($MachineGuid) {
     $imachines = Get-VirtualBoxVM -Guid $MachineGuid -SkipCheck
    } # get vm inventory (by $MachineGuid)
    if ($imachines) {
     foreach ($imachine in $imachines) {
      if ($imachine.State -ne 'PoweredOff') {Write-Host "[Error] The machine $($imachine.Name) is not powered off. Hotswap is not supported at this time. Power the machine off and try again." -ForegroundColor Red -BackgroundColor Black;return}
      if ($ModuleHost.ToLower() -eq 'websrv') {
       #$istoragecontrollers = New-Object IStorageController
       #$istoragecontrollers = $istoragecontrollers.Fetch($imachine.Id)
       foreach ($istoragecontroller in $imachine.IStorageControllers) {
        if ($istoragecontroller.Name -eq $Controller) {
         if ($ControllerPort -lt 0 -or $ControllerPort -gt $istoragecontroller.PortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough available ports. Specify a new port number and try again." -ForegroundColor Red -BackgroundColor Black;return}
         if ($ControllerSlot -lt 0 -or $ControllerSlot -gt $istoragecontroller.MaxDevicesPerPortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough slots available on the requseted port. Specify a new slot number and try again." -ForegroundColor Red -BackgroundColor Black;return}
         $controllerfound = $true
        } # end if $istoragecontroller.Name -eq $Controller
        if (!$controllerfound) {Write-Host "[Error] The controller $($istoragecontroller.Name) was not found. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black;return}
       } # foreach $istoragecontroller in $imachine.IStorageControllers
       Write-Verbose "Getting write lock on machine $($imachine.Name)"
       $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
       # create a new machine object
       $mmachine = New-Object VirtualBoxVM
       # get the mutable machine object
       Write-Verbose "Getting the mutable machine object"
       $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
       $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
       # attach the disk
       Write-Verbose "Mounting disk $($imedium.Name) to machine $($imachine.Name)"
       $global:vbox.IMachine_attachDevice($mmachine.Id, $Controller, $ControllerPort, $ControllerSlot, [DeviceType]::new().ToULong('HardDisk'), $imedium.Id)
       # save new settings
       Write-Verbose "Saving new settings"
       $global:vbox.IMachine_saveSettings($mmachine.Id)
       # unlock machine session
       Write-Verbose "Unlocking machine session"
       $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      } # end if websrv
      elseif ($ModuleHost.ToLower() -eq 'com') {
       $istoragecontrollers = $imachine.ComObject.StorageControllers
       foreach ($istoragecontroller in $istoragecontrollers) {
        if ($istoragecontroller.Name -eq $Controller) {
         if ($ControllerPort -lt 0 -or $ControllerPort -gt $istoragecontroller.PortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough available ports. Specify a new port number and try again." -ForegroundColor Red -BackgroundColor Black;return}
         if ($ControllerSlot -lt 0 -or $ControllerSlot -gt $istoragecontroller.MaxDevicesPerPortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough slots available on the requseted port. Specify a new slot number and try again." -ForegroundColor Red -BackgroundColor Black;return}
         $controllerfound = $true
        } # end if $istoragecontroller.Name -eq $Controller
        if (!$controllerfound) {Write-Host "[Error] The controller $($istoragecontroller.Name) was not found. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black;return}
       } # foreach $istoragecontroller in $istoragecontrollers
       Write-Verbose "Getting write lock on machine $($imachine.Name)"
       $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
       # create a new machine object
       $mmachine = New-Object VirtualBoxVM
       # get the mutable machine object
       Write-Verbose "Getting the mutable machine object"
       $mmachine.ComObject = $imachine.ISession.Session.Machine
       $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
       # wait for the disk to become available
       Write-Verbose "Waiting for the disk to become available"
       do {
       } until ($imedium.ComObject.State -eq 1)
       # attach the disk
       Write-Verbose "Mounting disk $($imedium.Name) to machine $($imachine.Name)"
       $mmachine.ComObject.AttachDevice($Controller, $ControllerPort, $ControllerSlot, [DeviceType]::new().ToULong('HardDisk'), $imedium.ComObject)
       # save new settings
       Write-Verbose "Saving new settings"
       $mmachine.ComObject.SaveSettings()
       # unlock machine session
       Write-Verbose "Unlocking machine session"
       $imachine.ISession.Session.UnlockMachine()
      } # end elseif com
     } # foreach $imachine in $imachines
    } # end if $imachines
    else {Write-Host "[Error] No machines found using specified filters." -ForegroundColor Red -BackgroundColor Black;return}
   } # foreach $imedium in $imediums
  } # Try
  catch {
   Write-Verbose 'Exception mounting virtual disk'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  finally {
   # release mutable machine objects if they exist
   if ($mmachine) {
    if ($mmachine.ISession.Id) {
     # release mutable session object
     Write-Verbose "Releasing mutable session object"
     $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
    }
    if ($mmachine.ISession.Session) {
     if ($mmachine.ISession.Session.State -gt 1) {
      $mmachine.ISession.Session.UnlockMachine()
     } # end if $mmachine.ISession.Session locked
    } # end if $mmachine.ISession.Session
    if ($mmachine.Id) {
     # release mutable object
     Write-Verbose "Releasing mutable object"
     $global:vbox.IManagedObjectRef_release($mmachine.Id)
    }
   }
   # obligatory session unlock
   Write-Verbose 'Cleaning up machine sessions'
   if ($imachines) {
    foreach ($imachine in $imachines) {
     if ($imachine.ISession.Id) {
      if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
       Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
       $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      } # end if session state not unlocked
     } # end if $imachine.ISession.Id
     if ($imachine.ISession.Session) {
      if ($imachine.ISession.Session.State -gt 1) {
       $imachine.ISession.Session.UnlockMachine()
      } # end if $imachine.ISession.Session locked
     } # end if $imachine.ISession.Session
     if ($imachine.IConsole) {
      # release the iconsole session
      Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
      $global:vbox.IManagedObjectRef_release($imachine.IConsole)
     } # end if $imachine.IConsole
     #$imachine.ISession.Id = $null
     $imachine.IConsole = $null
     if ($imachine.IPercent) {$imachine.IPercent = $null}
     $imachine.MSession = $null
     $imachine.MConsole = $null
     $imachine.MMachine = $null
    } # end foreach $imachine in $imachines
   } # end if $imachines
  } # Finally
 } # end if $imediums
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Dismount-VirtualBoxDisk {
<#
.SYNOPSIS
Dismount VirtualBox disk
.DESCRIPTION
Dismounts VirtualBox disks. The command will fail if the virtual disk is not attached to the specified virtual machine.
.PARAMETER Disk
At least one virtual disk object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual disk. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual disk. Can be received via pipeline input by name.
.PARAMETER MachineName
The name of the virtual machine to dismount the disk from. This is a required parameter.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Dismount-VirtualBoxDisk -Name TestDisk -MachineName Win10 -Controller SATA -ControllerPort 0 -ControllerSlot 0

Dismounts the virtual disk named "TestDisk.vmdk" from the Win10 virtual machine SATA controller on port 0 slot 0
.NOTES
NAME        :  Dismount-VirtualBoxDisk
VERSION     :  1.0
LAST UPDATED:  1/20/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisk
.INPUTS
VirtualBoxVHD[]:  VirtualBoxVHDs for virtual disk objects
String[]       :  Strings for virtual disk names
GUID[]         :  GUIDS for virtual disk GUIDS
String         :  String for virtual machine name
String         :  String for controller name
Int            :  Integer for controller port
Int            :  Integer for controller slot
.OUTPUTS
None
#>
[cmdletbinding(SupportsShouldProcess,ConfirmImpact='High')]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk object(s)",
Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVHD[]]$Disk,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(Mandatory=$false,HelpMessage="Enter a virtual machine object to dismount the disk from")]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(Mandatory=$false,HelpMessage="Enter the name of the virtual machine to dismount the disk from")]
[ValidateNotNullorEmpty()]
  [string]$MachineName,
[Parameter(Mandatory=$false,HelpMessage="Enter the GUID of the virtual machine to dismount the disk from")]
[ValidateNotNullorEmpty()]
  [guid]$MachineGuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 Write-Verbose "Pipeline - Disk: `"$Disk`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Disk -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one disk object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imediums = @()
 if ($Disk) {
  Write-Verbose "Getting disk inventory from Disk(s) object"
  $imediums = $Disk
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting disk inventory from Name(s)"
   $imediums += Get-VirtualBoxDisk -Name $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting disk inventory from GUID(s)"
   $imediums += Get-VirtualBoxDisk -Guid $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 $imediums = $imediums | Where-Object {$_.Format -notmatch 'RAW'}
 if ($imediums) {
  Write-Verbose "[Info] Found disks"
  try {
   foreach ($imedium in $imediums) {
    Write-Verbose "Found disk: $($imedium.Name)"
    if ($imedium.MachineGuid) {
     foreach ($vmids in $imedium.MachineGuid) {
      Write-Verbose "Disk attached to VM: $vmname"
      if ($Machine) {
       $imachines = $Machine
      } # get vm inventory (by $Machine)
      elseif ($MachineName) {
       $imachines = Get-VirtualBoxVM -Name $MachineName -SkipCheck
      } # get vm inventory (by $MachineName)
      elseif ($MachineGuid) {
       $imachines = Get-VirtualBoxVM -Guid $MachineGuid -SkipCheck
      } # get vm inventory (by $MachineGuid)
      elseif ($Machine -or $MachineName -or $MachineGuid -and !$imachines) {
       Write-Verbose "[Warning] No machines found using provided filters. Matching all machines."
       $imachines = Get-VirtualBoxVM -SkipCheck
      } # get vm inventory (fallback)
      if (!$imachines) {$imachines = Get-VirtualBoxVM -SkipCheck}
      if ($imachines) {
       foreach ($imachine in $imachines) {
        if ($imachine.State -ne 'PoweredOff') {Write-Host "[Error] The machine $($imachine.Name) is not powered off. Hotswap is not supported at this time. Power the machine off and try again." -ForegroundColor Red -BackgroundColor Black}
        if ($PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Dismount storage medium $($imedium.Name) ")) {
         if ($ModuleHost.ToLower() -eq 'websrv') {
          Write-Verbose "Getting medium attachment information"
          $imediumattachment = $global:vbox.IMachine_getMediumAttachments($imachine.Id) | Where-Object {$_.machine -match $imachine.Id} | Where-Object {$_.Medium -match $imedium.Id}
          Write-Verbose "Getting write lock on machine $($imachine.Name)"
          $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
          # create a new machine object
          $mmachine = New-Object VirtualBoxVM
          # get the mutable machine object
          Write-Verbose "Getting the mutable machine object"
          $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
          $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
          Write-Verbose "Attempting to unmount disk $($imedium.Name) from machine: $($imachine.Name)"
          $global:vbox.IMachine_detachDevice($mmachine.Id, $imediumattachment.controller, $imediumattachment.port, $imediumattachment.device)
          # save new settings
          Write-Verbose "Saving new settings"
          $global:vbox.IMachine_saveSettings($mmachine.Id)
          # unlock machine session
          Write-Verbose "Unlocking machine session"
          $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
         } # end if websrv
         elseif ($ModuleHost.ToLower() -eq 'com') {
          Write-Verbose "Getting medium attachment information"
          $imediumattachment = ($global:vbox.Machines | Where-Object {$_.Id -match $imachine.Guid}).MediumAttachments | Where-Object {$_.Medium.Id -match $imedium.Guid}
          Write-Verbose "Getting write lock on machine $($imachine.Name)"
          $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
          # create a new machine object
          $mmachine = New-Object VirtualBoxVM
          # get the mutable machine object
          Write-Verbose "Getting the mutable machine object"
          $mmachine.ComObject = $imachine.ISession.Session.Machine
          $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
          Write-Verbose "Attempting to unmount disk $($imedium.Name) from machine: $($imachine.Name)"
          Write-Verbose "Controller: `"$($imediumattachment.Controller)`""
          Write-Verbose "Port: `"$($imediumattachment.Port)`""
          Write-Verbose "Device: `"$($imediumattachment.Device)`""
          $mmachine.ComObject.DetachDevice($imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device)
          # save new settings
          Write-Verbose "Saving new settings"
          $mmachine.ComObject.SaveSettings()
          # unlock machine session
          Write-Verbose "Unlocking machine session"
          $imachine.ISession.Session.UnlockMachine()
         } # end elseif com
        } # end if $PSCmdlet.ShouldProcess(
       } # foreach $imachine in $imachines
      } # end if $imachines
     } # foreach $vmname in $imedium.MachineGuid
    } # end if $imedium.MachineGuid
   } # foreach $imedium in $imediums
  } # Try
  catch {
   Write-Verbose 'Exception dismounting virtual disk'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  finally {
   # release mutable machine objects if they exist
   if ($mmachine) {
    if ($mmachine.ISession.Id) {
     # release mutable session object
     Write-Verbose "Releasing mutable session object"
     $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
    }
    if ($mmachine.ISession.Session) {
     if ($mmachine.ISession.Session.State -gt 1) {
      $mmachine.ISession.Session.UnlockMachine()
     } # end if $mmachine.ISession.Session locked
    } # end if $mmachine.ISession.Session
    if ($mmachine.Id) {
     # release mutable object
     Write-Verbose "Releasing mutable object"
     $global:vbox.IManagedObjectRef_release($mmachine.Id)
    }
   }
   # obligatory session unlock
   Write-Verbose 'Cleaning up machine sessions'
   if ($imachines) {
    foreach ($imachine in $imachines) {
     if ($imachine.ISession.Id) {
      if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
       Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
       $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      } # end if session state not unlocked
     } # end if $imachine.ISession.Id
     if ($imachine.ISession.Session) {
      if ($imachine.ISession.Session.State -gt 1) {
       $imachine.ISession.Session.UnlockMachine()
      } # end if $imachine.ISession.Session locked
     } # end if $imachine.ISession.Session
     if ($imachine.IConsole) {
      # release the iconsole session
      Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
      $global:vbox.IManagedObjectRef_release($imachine.IConsole)
     } # end if $imachine.IConsole
     #$imachine.ISession.Id = $null
     $imachine.IConsole = $null
     if ($imachine.IPercent) {$imachine.IPercent = $null}
     $imachine.MSession = $null
     $imachine.MConsole = $null
     $imachine.MMachine = $null
    } # end foreach $imachine in $imachines
   } # end if $imachines
  } # Finally
 } # end if $imediums
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Edit-VirtualBoxDisk {
<#
.SYNOPSIS
Edit VirtualBox disk
.DESCRIPTION
Modifies VirtualBox disks. After a disk is modified, an updated virtual disk object will be returned via the pipeline. The command will fail if the virtual disk is not attached to the specified virtual machine.
.PARAMETER Disk
At least one virtual disk object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual disk. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual disk. Can be received via pipeline input by name.
.PARAMETER MachineName
The name of the virtual machine to dismount the disk from. This is a required parameter.
.PARAMETER AutoDiscard
Specify whether or not the TRIM is enabled for the virtual disk. This is silently ignored if it is not supported by the specified controller and/or drive.
.PARAMETER BandwidthGroup
Specify a bandwidth group for the virtual disk. Specify a blank string to clear the bandwidth group.
.PARAMETER NonRotational
Specify whether or not the SSD flag is enabled for the virtual disk. This is silently ignored if it is not supported by the specified controller and/or drive.
.PARAMETER HotPluggable
Specify whether or not the Hot-pluggable flag is enabled for the virtual disk. This is silently ignored if it is not supported by the specified controller and/or drive.
.PARAMETER TemporaryEject
Specify whether or not guest-triggered eject results in unmounting the virtual disk.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Edit-VirtualBoxDisk -Name TestDisk -MachineName Win10 -NonRotational $true

Enable the SSD flag for the virtual disk named "TestDisk.vmdk" attached to the Win10 virtual machine
.NOTES
NAME        :  Edit-VirtualBoxDisk
VERSION     :  1.0
LAST UPDATED:  2/13/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisk
.INPUTS
VirtualBoxVHD[]:  VirtualBoxVHDs for virtual disk objects
String[]       :  Strings for virtual disk names
GUID[]         :  GUIDS for virtual disk GUIDS
VirtualBoxVM[] :  VirtualBoxVHDs for virtual machine objects
String         :  String for virtual machine name
GUID[]         :  GUIDS for virtual machine GUIDS
Bool           :  Bool for SSD flag
.OUTPUTS
VirtualBoxVHD[]
#>
[cmdletbinding()]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk object(s)",
Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVHD[]]$Disk,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disk GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(Mandatory=$false,HelpMessage="Enter a virtual machine object to dismount the disk from")]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(Mandatory=$false,HelpMessage="Enter the name of the virtual machine to dismount the disk from")]
[ValidateNotNullorEmpty()]
  [string]$MachineName,
[Parameter(Mandatory=$false,HelpMessage="Enter the GUID of the virtual machine to dismount the disk from")]
[ValidateNotNullorEmpty()]
  [guid]$MachineGuid,
[Parameter(Mandatory=$false,HelpMessage="Specify if the virtual disk supports TRIM")]
[ValidateNotNullorEmpty()]
  [bool]$AutoDiscard,
[Parameter(Mandatory=$false,HelpMessage="Specify the virtual disk bandwidth group")]
[ValidateNotNullorEmpty()]
  [string]$BandwidthGroup,
[Parameter(Mandatory=$false,HelpMessage="Specify if the virtual disk is SSD")]
[ValidateNotNullorEmpty()]
  [bool]$NonRotational,
[Parameter(Mandatory=$false,HelpMessage="Specify if the virtual disk is hot-pluggable")]
[ValidateNotNullorEmpty()]
  [bool]$HotPluggable,
[Parameter(Mandatory=$false,HelpMessage="Specify if the virtual disk is guest ejectable")]
[ValidateNotNullorEmpty()]
  [bool]$TemporaryEject,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 Write-Verbose "Pipeline - Disk: `"$Disk`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Disk -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one disk object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imediums = @()
 if ($Disk) {
  Write-Verbose "Getting disk inventory from Disk(s) object"
  $imediums = $Disk
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting disk inventory from Name(s)"
   $imediums += Get-VirtualBoxDisk -Name $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting disk inventory from GUID(s)"
   $imediums += Get-VirtualBoxDisk -Guid $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 $imediums = $imediums | Where-Object {$_.Format -notmatch 'RAW'}
 if ($imediums) {
  Write-Verbose "[Info] Found disks"
  try {
   foreach ($imedium in $imediums) {
    Write-Verbose "Found disk: $($imedium.Name)"
    if ($imedium.MachineGuid) {
     foreach ($vmids in $imedium.MachineGuid) {
      Write-Verbose "Disk attached to VM: $vmname"
      if ($Machine) {
       $imachines = $Machine
      } # get vm inventory (by $Machine)
      elseif ($MachineName) {
       $imachines = Get-VirtualBoxVM -Name $MachineName -SkipCheck
      } # get vm inventory (by $MachineName)
      elseif ($MachineGuid) {
       $imachines = Get-VirtualBoxVM -Guid $MachineGuid -SkipCheck
      } # get vm inventory (by $MachineGuid)
      elseif ($Machine -or $MachineName -or $MachineGuid -and !$imachines) {
       Write-Verbose "[Warning] No machines found using provided filters. Matching all machines."
       $imachines = Get-VirtualBoxVM -SkipCheck
      } # get vm inventory (fallback)
      if (!$imachines) {$imachines = Get-VirtualBoxVM -SkipCheck}
      if ($imachines) {
       foreach ($imachine in $imachines) {
        if ($imachine.State -ne 'PoweredOff') {Write-Host "[Error] The machine $($imachine.Name) is not powered off. Power the machine off and try again." -ForegroundColor Red -BackgroundColor Black}
        if ($ModuleHost.ToLower() -eq 'websrv') {
         Write-Verbose "Getting medium attachment information"
         $imediumattachment = $global:vbox.IMachine_getMediumAttachments($imachine.Id) | Where-Object {$_.machine -match $imachine.Id} | Where-Object {$_.Medium -match $imedium.Id}
         Write-Verbose "Getting write lock on machine $($imachine.Name)"
         $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
         # create a new machine object
         $mmachine = New-Object VirtualBoxVM
         # get the mutable machine object
         Write-Verbose "Getting the mutable machine object"
         $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
         $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
         try {
          # set requested settings here
          if ($AutoDiscard -ne $null) {$vbox.IMachine_setAutoDiscardForDevice($mmachine.Id, $imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, $AutoDiscard)}
          if ($BandwidthGroup -ne $null) {
           if ($BandwidthGroup -eq '') {$vbox.IMachine_setNoBandwidthGroupForDevice($mmachine.Id, $imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device)}
           else {$vbox.IMachine_setBandwidthGroupForDevice($mmachine.Id, $imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, $BandwidthGroup)}
          }
          if ($NonRotational -ne $null) {$vbox.IMachine_nonRotationalDevice($mmachine.Id, $imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, $NonRotational)}
          if ($HotPluggable -ne $null) {$vbox.IMachine_setHotPluggableForDevice($mmachine.Id, $imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, $HotPluggable)}
          if ($TemporaryEject -ne $null) {$vbox.IMachine_temporaryEjectDevice($mmachine.Id, $imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, $TemporaryEject)}
         } # Try
         catch {
          Write-Verbose 'Exception applying new virtual disk settings'
          Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
          Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
         } # Catch
         # save new settings
         Write-Verbose "Saving new settings"
         $global:vbox.IMachine_saveSettings($mmachine.Id)
         # unlock machine session
         Write-Verbose "Unlocking machine session"
         $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
        } # end if websrv
        elseif ($ModuleHost.ToLower() -eq 'com') {
         Write-Verbose "Getting medium attachment information"
         $imediumattachment = ($global:vbox.Machines | Where-Object {$_.Id -match $imachine.Guid}).MediumAttachments | Where-Object {$_.Medium.Id -match $imedium.Guid}
         Write-Verbose "Getting write lock on machine $($imachine.Name)"
         $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
         # create a new machine object
         $mmachine = New-Object VirtualBoxVM
         # get the mutable machine object
         Write-Verbose "Getting the mutable machine object"
         $mmachine.ComObject = $imachine.ISession.Session.Machine
         $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
         Write-Verbose "Controller: `"$($imediumattachment.Controller)`""
         Write-Verbose "Port: `"$($imediumattachment.Port)`""
         Write-Verbose "Device: `"$($imediumattachment.Device)`""
         try {
          # set requested settings here
          if ($AutoDiscard -ne $null) {$mmachine.ComObject.SetAutoDiscardForDevice($imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, [int]$AutoDiscard)}
          if ($BandwidthGroup -ne $null) {
           if ($BandwidthGroup -eq '') {$mmachine.ComObject.SetNoBandwidthGroupForDevice($imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device)}
           else {$mmachine.ComObject.SetBandwidthGroupForDevice($imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, $BandwidthGroup)}
          }
          if ($NonRotational -ne $null) {$mmachine.ComObject.NonRotationalDevice($imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, [int]$NonRotational)}
          if ($HotPluggable -ne $null) {$mmachine.ComObject.SetHotPluggableForDevice($imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, [int]$HotPluggable)}
          if ($TemporaryEject -ne $null) {$mmachine.ComObject.TemporaryEjectDevice($imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, [int]$TemporaryEject)}
         } # Try
         catch {
          Write-Verbose 'Exception applying new virtual disk settings'
          Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
          Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
         } # Catch
         # save new settings
         Write-Verbose "Saving new settings"
         $mmachine.ComObject.SaveSettings()
         # unlock machine session
         Write-Verbose "Unlocking machine session"
         $imachine.ISession.Session.UnlockMachine()
        } # end elseif com
        # write the updated disk to the pipeline
        Write-Output (Get-VirtualBoxDisk -Guid $imedium.Guid)
       } # foreach $imachine in $imachines
      } # end if $imachines
     } # foreach $vmname in $imedium.MachineGuid
    } # end if $imedium.MachineGuid
   } # foreach $imedium in $imediums
  } # Try
  catch {
   Write-Verbose 'Exception editing virtual disk'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  finally {
   # release mutable machine objects if they exist
   if ($mmachine) {
    if ($mmachine.ISession.Id) {
     # release mutable session object
     Write-Verbose "Releasing mutable session object"
     $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
    }
    if ($mmachine.ISession.Session) {
     if ($mmachine.ISession.Session.State -gt 1) {
      $mmachine.ISession.Session.UnlockMachine()
     } # end if $mmachine.ISession.Session locked
    } # end if $mmachine.ISession.Session
    if ($mmachine.Id) {
     # release mutable object
     Write-Verbose "Releasing mutable object"
     $global:vbox.IManagedObjectRef_release($mmachine.Id)
    }
   }
   # obligatory session unlock
   Write-Verbose 'Cleaning up machine sessions'
   if ($imachines) {
    foreach ($imachine in $imachines) {
     if ($imachine.ISession.Id) {
      if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
       Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
       $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      } # end if session state not unlocked
     } # end if $imachine.ISession.Id
     if ($imachine.ISession.Session) {
      if ($imachine.ISession.Session.State -gt 1) {
       $imachine.ISession.Session.UnlockMachine()
      } # end if $imachine.ISession.Session locked
     } # end if $imachine.ISession.Session
     if ($imachine.IConsole) {
      # release the iconsole session
      Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
      $global:vbox.IManagedObjectRef_release($imachine.IConsole)
     } # end if $imachine.IConsole
     #$imachine.ISession.Id = $null
     $imachine.IConsole = $null
     if ($imachine.IPercent) {$imachine.IPercent = $null}
     $imachine.MSession = $null
     $imachine.MConsole = $null
     $imachine.MMachine = $null
    } # end foreach $imachine in $imachines
   } # end if $imachines
  } # Finally
 } # end if $imediums
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Get-VirtualBoxDisc {
<#
.SYNOPSIS
Get VirtualBox disc information
.DESCRIPTION
Retrieve VirtualBox discs by name, GUID, machine object, machine name, machine GUID, or all.
.PARAMETER Name
At least one virtual disc name.
.PARAMETER Guid
At least one virtual disc GUID.
.PARAMETER Machine
At least one virtual machine object. Can be received via pipeline input.
.PARAMETER MachineName
The name of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER MachineGuid
The GUID of at least one virtual machine. Can be received via pipeline input by name.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Get-VirtualBoxDisc -Name vboxguestadd

Name        : VBoxGuestAdditions.iso
Description :
Format      : RAW
Size        : 59516928
LogicalSize : 59516928
MachineGuid : {7353caa6-8cb6-4066-aec9-6c6a69a001b6, 15a4c311-3b89-4936-89c7-11d3340ced7a}
MachineName : {2016 Core, Win10}

Get virtual disc by Name
.EXAMPLE
PS C:\> Get-VirtualBoxVM -Name 2016 | Get-VirtualBoxDisc

Name        : VBoxGuestAdditions.iso
Description :
Format      : RAW
Size        : 59516928
LogicalSize : 59516928
MachineGuid : {7353caa6-8cb6-4066-aec9-6c6a69a001b6, 15a4c311-3b89-4936-89c7-11d3340ced7a}
MachineName : {2016 Core, Win10}

Get virtual disc by machine object from pipeline input
.EXAMPLE
PS C:\> Get-VirtualBoxDisc -MachineName 2016

Name        : VBoxGuestAdditions.iso
Description :
Format      : RAW
Size        : 59516928
LogicalSize : 59516928
MachineGuid : {7353caa6-8cb6-4066-aec9-6c6a69a001b6, 15a4c311-3b89-4936-89c7-11d3340ced7a}
MachineName : {2016 Core, Win10}

Get virtual disc by machine name
.EXAMPLE
PS C:\> Get-VirtualBoxDisc -MachineGuid c9d4dc35-3967-4009-993d-1c23ab4ff22b

Name        : Hiren's.BootCD.15.2.iso
Description :
Format      : RAW
Size        : 623890432
LogicalSize : 623890432
MachineGuid : {c9d4dc35-3967-4009-993d-1c23ab4ff22b}
MachineName : {GNS3 IOU VM_1.3}

Get virtual disc by machine GUID
.EXAMPLE
PS C:\> Get-VirtualBoxDisc

Name        : VBoxGuestAdditions.iso
Description :
Format      : RAW
Size        : 59516928
LogicalSize : 59516928
MachineGuid : {7353caa6-8cb6-4066-aec9-6c6a69a001b6, 15a4c311-3b89-4936-89c7-11d3340ced7a}
MachineName : {2016 Core, Win10}

Name        : Hiren's.BootCD.15.2.iso
Description :
Format      : RAW
Size        : 623890432
LogicalSize : 623890432
MachineGuid : {c9d4dc35-3967-4009-993d-1c23ab4ff22b}
MachineName : {GNS3 IOU VM_1.3}

Get all virtual discs in the VirtualBox inventory
.NOTES
NAME        :  Get-VirtualBoxDisc
VERSION     :  1.0
LAST UPDATED:  2/8/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
New-VirtualBoxDisc
.INPUTS
String[]      :  Strings for virtual disc names
Guid[]        :  VirtualBoxVMs for virtual machine objects
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
.OUTPUTS
VirtualBoxVHD[]
#>
[cmdletbinding(DefaultParameterSetName="Machine")]
Param(
[Parameter(HelpMessage="Enter one or more disc name(s)",
ParameterSetName="Disc",Mandatory=$false,Position=0)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(HelpMessage="Enter one or more disc GUID(s)",
ParameterSetName="Disc",Mandatory=$false,Position=0)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine object(s)",
ParameterSetName="Machine",Mandatory=$false,Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
ParameterSetName="Machine",Mandatory=$false,Position=0)]
  [string[]]$MachineName,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
ParameterSetName="Machine",Mandatory=$false,Position=0)]
  [guid[]]$MachineGuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - MachineName: `"$MachineName`""
 Write-Verbose "Pipeline - MachineGuid: `"$MachineGuid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 $discs = @()
 $obj = @()
 try {
  # get virtual machine disc inventory
  Write-Verbose "Getting virtual disc inventory"
  if ($ModuleHost.ToLower() -eq 'websrv') {
   foreach ($imediumid in $global:vbox.IVirtualBox_getDVDImages($global:ivbox)) {
    Write-Verbose "Getting disc: $($imediumid)"
    $disc = New-Object VirtualBoxVHD
    $disc.Name = $global:vbox.IMedium_getName($imediumid)
    $disc.GUID = $global:vbox.IMedium_getId($imediumid)
    $disc.Description = $global:vbox.IMedium_getDescription($imediumid)
    $disc.Format = $global:vbox.IMedium_getFormat($imediumid)
    $disc.Size = $global:vbox.IMedium_getSize($imediumid)
    $disc.LogicalSize = $global:vbox.IMedium_getLogicalSize($imediumid)
    $disc.MachineGuid = $global:vbox.IMedium_getMachineIds($imediumid)
    foreach ($machineid in $disc.MachineGuid) {
     foreach ($imachine in ($global:vbox.IVirtualBox_getMachines($global:ivbox))) {
      if (($global:vbox.IMachine_getId($imachine)) -eq $machineid) {
       $disc.MachineName += $global:vbox.IMachine_getName($imachine)
      } # end if $imachine.Guid -eq $machineid
      $disc.MachineName = $disc.MachineName | Where-Object {$_ -ne $null}
     } # foreach $imachine in $imachines
    } # foreach $machineid in $disc.MachineGuid
    $disc.State = $global:vbox.IMedium_getState($imediumid)
    $disc.Variant = $global:vbox.IMedium_getVariant($imediumid)
    $disc.Location = $global:vbox.IMedium_getLocation($imediumid)
    $disc.HostDrive = $global:vbox.IMedium_getHostDrive($imediumid)
    $disc.MediumFormat = $global:vbox.IMedium_getMediumFormat($imediumid)
    $disc.Type = $global:vbox.IMedium_getType($imediumid)
    $disc.Parent = $global:vbox.IMedium_getParent($imediumid)
    $disc.Children = $global:vbox.IMedium_getChildren($imediumid)
    $disc.Id = $imediumid
    $disc.ReadOnly = $global:vbox.IMedium_getReadOnly($imediumid)
    $disc.AutoReset = $global:vbox.IMedium_getAutoReset($imediumid)
    $disc.LastAccessError = $global:vbox.IMedium_getLastAccessError($imediumid)
    [VirtualBoxVHD[]]$discs += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;}
   } # end foreach loop inventory
  } # end if websrv
  elseif ($ModuleHost.ToLower() -eq 'com') {
   foreach ($imedium in $vbox.DVDImages) {
    Write-Verbose "Getting disc: $($imedium.Id)"
    $disc = New-Object VirtualBoxVHD
    $disc.Name = $imedium.Name
    $disc.Guid = $imedium.Id
    $disc.Description = $imedium.Description
    $disc.Format = $imedium.Format
    $disc.Size = $imedium.Size
    $disc.LogicalSize = $imedium.LogicalSize
    $disc.MachineGuid = $imedium.MachineIds
    foreach ($machineid in $disc.MachineGuid) {
     foreach ($imachine in $global:vbox.Machines) {
      if ($imachine.Id -eq $machineid) {
       $disc.MachineName += $imachine.Name
      } # end if $imachine.Guid -eq $machineid
      $disc.MachineName = $disc.MachineName | Where-Object {$_ -ne $null}
     } # foreach $imachine in $imachines
    } # foreach $machineid in $disc.MachineGuid
    $disc.State = [MediumState]::new().ToStr($imedium.State)
    foreach ($variant in $imedium.Variant) {
     $disc.Variant += [MediumVariant]::new().ToStrCom($variant)
    } # foreach $variant in $imedium.Variant
    $disc.Location = $imedium.Location
    $disc.HostDrive = $imedium.HostDrive
    $disc.MediumFormat = $imedium.MediumFormat.Name
    $disc.Type = [MediumType]::new().ToStr($imedium.Type)
    if ($imedium.Parent) {$disc.Parent = $imedium.Parent.Name}
    if ($imedium.Children) {$disc.Children = $imedium.Children.Name}
    $disc.ComObject = $imedium
    $disc.ReadOnly = $imedium.ReadOnly
    $disc.AutoReset = $imedium.AutoReset
    $disc.LastAccessError = $imedium.LastAccessError
    [VirtualBoxVHD[]]$discs += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;ComObject=$disc.ComObject}
   } # end foreach loop inventory
  } # end elseif com
 } # Try
 catch {
  Write-Verbose 'Exception retrieving virtual disc information'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 if ($PSCmdlet.ParameterSetName -eq "Disc") {
  # filter by disc name
  if ($Name -and $Name -ne '*') {
   foreach ($item in $Name) {
    foreach ($disc in $discs) {
     $matched = $false
     Write-Verbose "Matching $($disc.Name) to $item"
     if ($disc.Name -match $item) {Write-Verbose "Matched $($disc.Name) to $item";$matched = $true}
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;ComObject=$disc.ComObject}}
    } # foreach $disc in $discs
   } # foreach $item in $Name
  } # end if $Name
  # filter by disc format
  elseif ($Format -and $Format -ne '*') {
   foreach ($item in $Format) {
    foreach ($disc in $discs) {
     $matched = $false
     Write-Verbose "Matching $($disc.Format) to $item"
     if ($disc.Format -match $item) {Write-Verbose "Matched $($disc.Format) to $item";$matched = $true}
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;ComObject=$disc.ComObject}}
    } # foreach $disc in $discs
   } # foreach $item in $Format
   $obj = $obj | Where-Object {$_ -ne $null}
  } # end if $Format
  # filter by disc guid
  elseif ($Guid) {
   foreach ($item in $Guid) {
    foreach ($disc in $discs) {
     $matched = $false
     Write-Verbose "Matching $($disc.Guid) to $item"
     if ($disc.Guid -match $item) {Write-Verbose "Matched $($disc.Guid) to $item";$matched = $true}
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;ComObject=$disc.ComObject}}
    } # foreach $disc in $discs
   } # foreach $item in $Guid
  } # end if $Guid
  # no filter
  else {foreach ($disc in $discs) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;ComObject=$disc.ComObject}}}
  Write-Verbose "Found $(($obj | Measure-Object).count) disc(s)"
 }
 elseif ($PSCmdlet.ParameterSetName -eq "Machine") {
  # filter by machine object
  if ($Machine) {
   foreach ($item in $Machine) {
    foreach ($disc in $discs) {
     $matched = $false
     foreach ($vmname in $disc.MachineName) {
      Write-Verbose "Matching $vmname to $($item.Name)"
      if ($vmname -match $item.Name) {Write-Verbose "Matched $vmname to $($item.Name)";$matched = $true}
     } # foreach $vmname in $disc.MachineName
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;ComObject=$disc.ComObject}}
    } # foreach $disc in $discs
   } # foreach $item in $Machine
  } # end if $Machine
  # filter by machine name
  elseif ($MachineName) {
   foreach ($item in $MachineName) {
    foreach ($disc in $discs) {
     $matched = $false
     foreach ($vmname in $disc.MachineName) {
      Write-Verbose "Matching $vmname to $item"
      if ($vmname -match $item) {Write-Verbose "Matched $vmname to $item";$matched = $true}
     } # foreach $vmname in $disc.MachineName
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;ComObject=$disc.ComObject}}
    } # foreach $disc in $discs
   } # foreach $item in $MachineName
  } # end elseif $MachineName
  # filter by machine GUID
  elseif ($MachineGuid) {
   foreach ($item in $MachineGuid) {
    foreach ($disc in $discs) {
     $matched = $false
     foreach ($vmguid in $disc.MachineGuid) {
      Write-Verbose "Matching $vmguid to $item"
      if ($vmguid -eq $item) {Write-Verbose "Matched $vmguid to $item";$matched = $true}
     } # foreach $vmguid in $disc.MachineGuid
     if ($matched -eq $true) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;ComObject=$disc.ComObject}}
    } # foreach $disc in $discs
   } # foreach $item in $MachineGuid
  } # end elseif $MachineGuid
  # no filter
  else {foreach ($disc in $discs) {[VirtualBoxVHD[]]$obj += [VirtualBoxVHD]@{Name=$disc.Name;Guid=$disc.Guid;Description=$disc.Description;Format=$disc.Format;Size=$disc.Size;LogicalSize=$disc.LogicalSize;MachineGuid=$disc.MachineGuid;MachineName=$disc.MachineName;State=$disc.State;Variant=$disc.Variant;Location=$disc.Location;HostDrive=$disc.HostDrive;MediumFormat=$disc.MediumFormat;Type=$disc.Type;Parent=$disc.Parent;Children=$disc.Children;Id=$disc.Id;ReadOnly=$disc.ReadOnly;AutoReset=$disc.AutoReset;LastAccessError=$disc.LastAccessError;ComObject=$disc.ComObject}}}
  Write-Verbose "Found $(($obj | Measure-Object).count) disc(s)"
 }
 if ($obj) {
  # write virtual machines object to the pipeline as an array
  Write-Output ([System.Array]$obj)
 } # end if $obj
 else {
  Write-Verbose "[Warning] No matching virtual discs found."
 } # end else
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Import-VirtualBoxDisc {
<#
.SYNOPSIS
Import VirtualBox disc
.DESCRIPTION
Imports VirtualBox discs. The command will fail if a virtual disc with the same name exists in the VirtualBox inventory.
.PARAMETER FileName
The full path to the virtual disc file.
.PARAMETER ForceNewUuid
A switch to request a new disc UUID be created.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Import-VirtualBoxDisc -FileName C:\Discs\TestDisc.iso

Import the "C:\Discs\TestDisc.iso" disc to the VirtualBox inventory
.NOTES
NAME        :  Import-VirtualBoxDisc
VERSION     :  1.0
LAST UPDATED:  2/8/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisc
.INPUTS
String        :  String for virtual disc file path
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(HelpMessage="Enter the full virtual disc path",
Mandatory=$true,Position=2)]
[ValidateScript({Test-Path $_})]
  [string]$FileName,
[Parameter(HelpMessage="Use this switch to request a new disc UUID be created",
Mandatory=$false)]
  [switch]$ForceNewUuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 $existingdiscs = Get-VirtualBoxDisc -Name $($FileName.Substring($FileName.LastIndexOf('\')+1)) -SkipCheck
 if ($existingdiscs) {
  foreach ($existingdisc in $existingdiscs) {
   Write-Verbose $existingdisc.Name
   if ($existingdisc.Name -eq ($FileName.Substring($FileName.LastIndexOf('\')+1))) {
    Write-Host "[Error] Disc $($existingdisc.Name) already exists. Select another disc image and try again." -ForegroundColor Red -BackgroundColor Black
    return
   }
  }
 }
 try {
  $imedium = New-Object VirtualBoxVHD
  if ($ModuleHost.ToLower() -eq 'websrv') {
   $imedium.Id = $global:vbox.IVirtualBox_openMedium($global:ivbox, $FileName, [DeviceType]::new().ToULong('DVD'), [AccessMode]::new().ToULong('ReadOnly'), $(if ($ForceNewUuid) {$true} else {$false}))
  } # end if websrv
  elseif ($ModuleHost.ToLower() -eq 'com') {
   $imedium.ComObject = $global:vbox.OpenMedium($FileName, [DeviceType]::new().ToULong('DVD'), [AccessMode]::new().ToULong('ReadOnly'), $(if ($ForceNewUuid) {1} else {0}))
  } # end elseif com
 } # Try
 catch {
  Write-Verbose 'Exception importing virtual disc'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Remove-VirtualBoxDisc {
<#
.SYNOPSIS
Remove VirtualBox disc
.DESCRIPTION
Removes VirtualBox discs. The command will fail if a virtual disc does not exist in the VirtualBox inventory or if the disc is mounted to a machine.
.PARAMETER Disc
At least one virtual disc object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Remove-VirtualBoxDisc -Name TestDisc.iso -ProgressBar -Confirm:$false

Remove the virtual disc named "TestDisc.iso" from the VirtualBox inventory, do not confirm the action, and display a progress bar
.NOTES
NAME        :  Remove-VirtualBoxDisc
VERSION     :  1.0
LAST UPDATED:  2/8/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisc
.INPUTS
VirtualBoxVHD[]:  VirtualBoxVHDs for virtual disc objects
String[]       :  Strings for virtual disc names
GUID[]         :  GUIDS for virtual disc GUIDS
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disc object(s)",
Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVHD[]]$Disc,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disc name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disc GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Use this switch to display a progress bar")]
  [switch]$ProgressBar,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 Write-Verbose "Pipeline - Disc: `"$Disc`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 Write-Verbose "Controller Name: `"$Controller`""
 Write-Verbose "Controller Port: `"$ControllerPort`""
 Write-Verbose "Controller Slot: `"$ControllerSlot`""
 if (!($Disc -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one disc object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imediums = @()
 if ($Disc) {
  Write-Verbose "Getting disc inventory from Disc(s) object"
  $imediums = $Disc
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting disc inventory from Name(s)"
   $imediums += Get-VirtualBoxDisc -Name $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting disc inventory from GUID(s)"
   $imediums += Get-VirtualBoxDisc -Guid $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 $imediums = $imediums | Where-Object {$_.Format -match 'RAW'}
 if ($imediums) {
  Write-Verbose "[Info] Found discs"
  try {
   foreach ($imedium in $imediums) {
    Write-Verbose "Found disc: $($imedium.Name)"
    if ($imedium.MachineName) {
     foreach ($vmname in $imedium.MachineName) {
     Write-Verbose "Disc attached to VM: $vmname"
      Write-Host "[Error] The disc $($imedium.Name) is still mounted to machine $vmname. Dismount the disc from the machine and try again." -ForegroundColor Red -BackgroundColor Black;return
     } # foreach $vmname in $imedium.MachineName
    } # end if $imedium.MachineName
    # close the disc
    Write-Verbose "Removing disc $($imedium.Name) from VirtualBox inventory"
    if ($ModuleHost.ToLower() -eq 'websrv') {
     $global:vbox.IMedium_close($imedium.Id)
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     $imedium.ComObject.Close()
    } # end elseif com
   } # foreach $imedium in $imediums
  } # Try
  catch {
   Write-Verbose 'Exception removing virtual disc'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  finally {
   # cleanup
  } # Finally
 } # end if $imediums
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Mount-VirtualBoxDisc {
<#
.SYNOPSIS
Mount VirtualBox disc
.DESCRIPTION
Mounts VirtualBox discs. The command will fail if a virtual disc is already mounted to the specified virtual machine.
.PARAMETER Disc
At least one virtual disc object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER MachineName
The name of the virtual machine to mount the disc to. This is a required parameter.
.PARAMETER Controller
The name of the storage controller to mount the disc to. This is a required parameter.
.PARAMETER ControllerPort
The port of the storage controller to mount the disc to. This is a required parameter.
.PARAMETER ControllerSlot
The slot of the storage controller to mount the disc to. This is a required parameter.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Mount-VirtualBoxDisc -Name TestDisc -MachineName Win10 -Controller SATA -ControllerPort 0 -ControllerSlot 0

Mounts the virtual disc named "TestDisc.vmdk" to the Win10 virtual machine SATA controller on port 0 slot 0
.NOTES
NAME        :  Mount-VirtualBoxDisc
VERSION     :  1.0
LAST UPDATED:  1/20/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisc
.INPUTS
VirtualBoxVHD[]:  VirtualBoxVHDs for virtual disc objects
String[]       :  Strings for virtual disc names
GUID[]         :  GUIDS for virtual disc GUIDS
String         :  String for virtual machine name
String         :  String for controller name
Int            :  Integer for controller port
Int            :  Integer for controller slot
.OUTPUTS
None
#>
[cmdletbinding()]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disc object(s)",
Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVHD[]]$Disc,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disc name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disc GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(Mandatory=$false,HelpMessage="Enter a virtual machine object to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(Mandatory=$false,HelpMessage="Enter the name of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [string]$MachineName,
[Parameter(Mandatory=$false,HelpMessage="Enter the GUID of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [guid]$MachineGuid,
[Parameter(Mandatory=$true,HelpMessage="Enter the name of the controller to mount the disc to",
ParameterSetName="HardDisc")]
[ValidateNotNullorEmpty()]
  [string]$Controller,
[Parameter(Mandatory=$true,HelpMessage="Enter the port number to mount the disc to",
ParameterSetName="HardDisc")]
[ValidateNotNullorEmpty()]
  [int]$ControllerPort,
[Parameter(Mandatory=$true,HelpMessage="Enter the slot number to mount the disc to",
ParameterSetName="HardDisc")]
[ValidateNotNullorEmpty()]
  [int]$ControllerSlot,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 Write-Verbose "Pipeline - Disc: `"$Disc`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 Write-Verbose "Machine: `"$Machine`""
 Write-Verbose "Machine Name: `"$MachineName`""
 Write-Verbose "Machine Guid: `"$MachineGuid`""
 Write-Verbose "Controller Name: `"$Controller`""
 Write-Verbose "Controller Port: `"$ControllerPort`""
 Write-Verbose "Controller Slot: `"$ControllerSlot`""
 if (!($Disc -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one disc object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 if (!($Machine -or $MachineName -or $MachineGuid)) {Write-Host "[Error] You must supply at least one machine object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imediums = @()
 if ($Disc) {
  Write-Verbose "Getting disc inventory from Disc(s) object"
  $imediums = $Disc
  $imediums = $imediums | Where-Object {$_ -ne $null}
 } # get disc inventory (by $Disc)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting disc inventory from Name(s)"
   $imediums += Get-VirtualBoxDisc -Name $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 } # get disc inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting disc inventory from GUID(s)"
   $imediums += Get-VirtualBoxDisc -Guid $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 } # get vm inventory (by $Guid)
 if ($imediums) {
  Write-Verbose "[Info] Found discs"
  try {
   foreach ($imedium in $imediums) {
    Write-Verbose "Found disc: $($imedium.Name)"
    <#
    if ($imedium.MachineName) {
     # make sure it's not already attached to the requested vm
     foreach ($vmname in $imedium.MachineName) {
      Write-Verbose "Disc attached to VM: $vmname"
      $imachine | Select-Object *
      if (Get-VirtualBoxDisc -MachineName $vmname -SkipCheck) {Write-Host "[Error] The disc $($imedium.Name) is already mounted to the $($imachine.Name) machine." -ForegroundColor Red -BackgroundColor Black;return}
     } # foreach $vmname in $imedium.MachineName
    } # end if $imedium.MachineName
    #>
    if ($Machine) {
     $imachines = $Machine
    } # get vm inventory (by $Machine)
    elseif ($MachineName) {
     $imachines = Get-VirtualBoxVM -Name $MachineName -SkipCheck
    } # get vm inventory (by $MachineName)
    elseif ($MachineGuid) {
     $imachines = Get-VirtualBoxVM -Guid $MachineGuid -SkipCheck
    } # get vm inventory (by $MachineGuid)
    if ($imachines) {
     foreach ($imachine in $imachines) {
      if ($imachine.State -ne 'PoweredOff') {Write-Host "[Error] The machine $($imachine.Name) is not powered off. Hotswap is not supported at this time. Power the machine off and try again." -ForegroundColor Red -BackgroundColor Black;return}
      if ($ModuleHost.ToLower() -eq 'websrv') {
       <#
       $istoragecontrollers = $imachine.ComObject.StorageControllers
       foreach ($istoragecontroller in $istoragecontrollers) {
        if ($istoragecontroller.Name -eq $Controller) {
         if ($ControllerPort -lt 0 -or $ControllerPort -gt $istoragecontroller.PortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough available ports. Specify a new port number and try again." -ForegroundColor Red -BackgroundColor Black;return}
         if ($ControllerSlot -lt 0 -or $ControllerSlot -gt $istoragecontroller.MaxDevicesPerPortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough slots available on the requseted port. Specify a new slot number and try again." -ForegroundColor Red -BackgroundColor Black;return}
         $controllerfound = $true
        } # end if $istoragecontroller.Name -eq $Controller
        if (!$controllerfound) {Write-Host "[Error] The controller $($istoragecontroller.Name) was not found. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black;return}
       } # foreach $istoragecontroller in $istoragecontrollers
       #>
       $storagecontrollers = Get-VirtualBoxVMStorageController -Name $Controller -Machine $imachine -SkipCheck
       if (!$storagecontrollers) {Write-Host "[Error] The controller $Controller was not found on the $($imachine.Name) machine. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black}
       else {
        foreach ($storagecontroller in $storagecontrollers) {
         Write-Verbose "Getting write lock on machine $($imachine.Name)"
         $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
         # create a new machine object
         $mmachine = New-Object VirtualBoxVM
         # get the mutable machine object
         Write-Verbose "Getting the mutable machine object"
         $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
         $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
         # attach the disc
         Write-Verbose "Mounting disc $($imedium.Name) to machine $($imachine.Name)"
         $global:vbox.IMachine_mountMedium($mmachine.Id, $storagecontroller.Name, $ControllerPort, $ControllerSlot, $imedium.Id, $false)
         # save new settings
         Write-Verbose "Saving new settings"
         $global:vbox.IMachine_saveSettings($mmachine.Id)
         # unlock machine session
         Write-Verbose "Unlocking machine session"
         $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
        } # foreach $storagecontroller in $storagecontrollers
       } # end else
      } # end if websrv
      elseif ($ModuleHost.ToLower() -eq 'com') {
       <#
       $istoragecontrollers = $imachine.ComObject.StorageControllers
       foreach ($istoragecontroller in $istoragecontrollers) {
        if ($istoragecontroller.Name -eq $Controller) {
         if ($ControllerPort -lt 0 -or $ControllerPort -gt $istoragecontroller.PortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough available ports. Specify a new port number and try again." -ForegroundColor Red -BackgroundColor Black;return}
         if ($ControllerSlot -lt 0 -or $ControllerSlot -gt $istoragecontroller.MaxDevicesPerPortCount) {Write-Host "[Error] The controller $($istoragecontroller.Name) does not have enough slots available on the requseted port. Specify a new slot number and try again." -ForegroundColor Red -BackgroundColor Black;return}
         $controllerfound = $true
        } # end if $istoragecontroller.Name -eq $Controller
        if (!$controllerfound) {Write-Host "[Error] The controller $($istoragecontroller.Name) was not found. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black;return}
       } # foreach $istoragecontroller in $istoragecontrollers
       #>
       $storagecontrollers = Get-VirtualBoxVMStorageController -Name $Controller -Machine $imachine -SkipCheck
       if (!$storagecontrollers) {Write-Host "[Error] The controller $Controller was not found on the $($imachine.Name) machine. Specify an existing controller name and try again." -ForegroundColor Red -BackgroundColor Black}
       else {
        foreach ($storagecontroller in $storagecontrollers) {
         Write-Verbose "Getting write lock on machine $($imachine.Name)"
         $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
         # create a new machine object
         $mmachine = New-Object VirtualBoxVM
         # get the mutable machine object
         Write-Verbose "Getting the mutable machine object"
         $mmachine.ComObject = $imachine.ISession.Session.Machine
         $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
         # wait for the disc to become available
         Write-Verbose "Waiting for the disc to become available"
         do {} until ($imedium.ComObject.State -eq 1)
         # attach the disc
         Write-Verbose "Mounting disc $($imedium.Name) to machine $($imachine.Name)"
         $mmachine.ComObject.MountMedium($storagecontroller.Name, $ControllerPort, $ControllerSlot, $imedium.ComObject, $false)
         # save new settings
         Write-Verbose "Saving new settings"
         $mmachine.ComObject.SaveSettings()
         # unlock machine session
         Write-Verbose "Unlocking machine session"
         $imachine.ISession.Session.UnlockMachine()
        } # foreach $storagecontroller in $storagecontrollers
       } # end else
      } # end elseif com
     } # foreach $imachine in $imachines
    } # end if $imachines
    else {Write-Host "[Error] No machines found using specified filters." -ForegroundColor Red -BackgroundColor Black;return}
   } # foreach $imedium in $imediums
  } # Try
  catch {
   Write-Verbose 'Exception mounting virtual disc'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  finally {
   # release mutable machine objects if they exist
   if ($mmachine) {
    if ($mmachine.ISession.Id) {
     # release mutable session object
     Write-Verbose "Releasing mutable session object"
     $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
    }
    if ($mmachine.ISession.Session) {
     if ($mmachine.ISession.Session.State -gt 1) {
      $mmachine.ISession.Session.UnlockMachine()
     } # end if $mmachine.ISession.Session locked
    } # end if $mmachine.ISession.Session
    if ($mmachine.Id) {
     # release mutable object
     Write-Verbose "Releasing mutable object"
     $global:vbox.IManagedObjectRef_release($mmachine.Id)
    }
   }
   # obligatory session unlock
   Write-Verbose 'Cleaning up machine sessions'
   if ($imachines) {
    foreach ($imachine in $imachines) {
     if ($imachine.ISession.Id) {
      if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
       Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
       $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      } # end if session state not unlocked
     } # end if $imachine.ISession.Id
     if ($imachine.ISession.Session) {
      if ($imachine.ISession.Session.State -gt 1) {
       $imachine.ISession.Session.UnlockMachine()
      } # end if $imachine.ISession.Session locked
     } # end if $imachine.ISession.Session
     if ($imachine.IConsole) {
      # release the iconsole session
      Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
      $global:vbox.IManagedObjectRef_release($imachine.IConsole)
     } # end if $imachine.IConsole
     #$imachine.ISession.Id = $null
     $imachine.IConsole = $null
     if ($imachine.IPercent) {$imachine.IPercent = $null}
     $imachine.MSession = $null
     $imachine.MConsole = $null
     $imachine.MMachine = $null
    } # end foreach $imachine in $imachines
   } # end if $imachines
  } # Finally
 } # end if $imediums
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Dismount-VirtualBoxDisc {
<#
.SYNOPSIS
Dismount VirtualBox disc
.DESCRIPTION
Dismounts VirtualBox discs. The command will fail if the virtual disc is not attached to the specified virtual machine.
.PARAMETER Disc
At least one virtual disc object. Can be received via pipeline input.
.PARAMETER Name
The name of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER Guid
The GUID of at least one virtual disc. Can be received via pipeline input by name.
.PARAMETER MachineName
The name of the virtual machine to dismount the disc from. This is a required parameter.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Dismount-VirtualBoxDisc -Name TestDisc -MachineName Win10 -Controller SATA -ControllerPort 0 -ControllerSlot 0

Dismounts the virtual disc named "TestDisc.vmdk" from the Win10 virtual machine SATA controller on port 0 slot 0
.NOTES
NAME        :  Dismount-VirtualBoxDisc
VERSION     :  1.0
LAST UPDATED:  1/20/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Get-VirtualBoxDisc
.INPUTS
VirtualBoxVHD[]:  VirtualBoxVHDs for virtual disc objects
String[]       :  Strings for virtual disc names
GUID[]         :  GUIDS for virtual disc GUIDS
String         :  String for virtual machine name
String         :  String for controller name
Int            :  Integer for controller port
Int            :  Integer for controller slot
.OUTPUTS
None
#>
[cmdletbinding(SupportsShouldProcess,ConfirmImpact='High')]
Param(
[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disc object(s)",
Position=0)]
[ValidateNotNullorEmpty()]
  [VirtualBoxVHD[]]$Disc,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disc name(s)")]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual disc GUID(s)")]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(Mandatory=$false,HelpMessage="Enter a virtual machine object to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [VirtualBoxVM]$Machine,
[Parameter(Mandatory=$false,HelpMessage="Enter the name of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [string]$MachineName,
[Parameter(Mandatory=$false,HelpMessage="Enter the GUID of the virtual machine to dismount the disc from")]
[ValidateNotNullorEmpty()]
  [guid]$MachineGuid,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
 # get extensions supported by the selected format
 $Ext = ($global:mediumformatspso | Where-Object {$_.Name -match $Format}).Extensions
 # get the last of the extensions and use it
 $Ext = $Ext[$Ext.GetUpperBound(0)]
} # Begin
Process {
 Write-Verbose "Pipeline - Disc: `"$Disc`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Disc -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one disc object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imediums = @()
 if ($Disc) {
  Write-Verbose "Getting disc inventory from Disc(s) object"
  $imediums = $Disc
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting disc inventory from Name(s)"
   $imediums += Get-VirtualBoxDisc -Name $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting disc inventory from GUID(s)"
   $imediums += Get-VirtualBoxDisc -Guid $item -SkipCheck
  }
  $imediums = $imediums | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 if ($imediums) {
  Write-Verbose "[Info] Found discs"
  try {
   foreach ($imedium in $imediums) {
    Write-Verbose "Found disc: $($imedium.Name)"
    if ($imedium.MachineGuid) {
     foreach ($vmids in $imedium.MachineGuid) {
      Write-Verbose "Disc attached to VM: $vmname"
      if ($Machine) {
       $imachines = $Machine
      } # get vm inventory (by $Machine)
      elseif ($MachineName) {
       $imachines = Get-VirtualBoxVM -Name $MachineName -SkipCheck
      } # get vm inventory (by $MachineName)
      elseif ($MachineGuid) {
       $imachines = Get-VirtualBoxVM -Guid $MachineGuid -SkipCheck
      } # get vm inventory (by $MachineGuid)
      elseif ($Machine -or $MachineName -or $MachineGuid -and !$imachines) {
       Write-Verbose "[Warning] No machines found using provided filters. Matching all machines."
       $imachines = Get-VirtualBoxVM -SkipCheck
      } # get vm inventory (fallback)
      if (!$imachines) {$imachines = Get-VirtualBoxVM -SkipCheck}
      if ($imachines) {
       foreach ($imachine in $imachines) {
        if ($imachine.State -ne 'PoweredOff') {Write-Host "[Error] The machine $($imachine.Name) is not powered off. Hotswap is not supported at this time. Power the machine off and try again." -ForegroundColor Red -BackgroundColor Black}
        if ($PSCmdlet.ShouldProcess("$($imachine.Name) virtual machine" , "Dismount storage medium $($imedium.Name) ")) {
         if ($ModuleHost.ToLower() -eq 'websrv') {
          Write-Verbose "Getting medium attachment information"
          $imediumattachments = $global:vbox.IMachine_getMediumAttachments($imachine.Id) | Where-Object {$_.machine -match $imachine.Id} | Where-Object {$_.Medium -match $imedium.Id}
          foreach ($imediumattachment in $imediumattachments) {
           Write-Verbose "Getting write lock on machine $($imachine.Name)"
           $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Write'))
           # create a new machine object
           $mmachine = New-Object VirtualBoxVM
           # get the mutable machine object
           Write-Verbose "Getting the mutable machine object"
           $mmachine.Id = $global:vbox.ISession_getMachine($imachine.ISession.Id)
           $mmachine.ISession.Id = $global:vbox.IWebsessionManager_getSessionObject($global:ivbox)
           Write-Verbose "Attempting to unmount disc $($imedium.Name) from machine: $($imachine.Name)"
           $global:vbox.IMachine_unmountMedium($mmachine.Id, $imediumattachment.controller, $imediumattachment.port, $imediumattachment.device, $false)
           # save new settings
           Write-Verbose "Saving new settings"
           $global:vbox.IMachine_saveSettings($mmachine.Id)
           # unlock machine session
           Write-Verbose "Unlocking machine session"
           $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
          } # foreach $imediumattachment in $imediumattachments
         } # end if websrv
         elseif ($ModuleHost.ToLower() -eq 'com') {
          Write-Verbose "Getting medium attachment information"
          $imediumattachments = ($global:vbox.Machines | Where-Object {$_.Id -match $imachine.Guid}).MediumAttachments | Where-Object {$_.Medium.Id -match $imedium.Guid}
          foreach ($imediumattachment in $imediumattachments) {
           Write-Verbose "Getting write lock on machine $($imachine.Name)"
           $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Write'))
           # create a new machine object
           $mmachine = New-Object VirtualBoxVM
           # get the mutable machine object
           Write-Verbose "Getting the mutable machine object"
           $mmachine.ComObject = $imachine.ISession.Session.Machine
           $mmachine.ISession.Session = New-Object -ComObject VirtualBox.Session
           Write-Verbose "Attempting to unmount disc $($imedium.Name) from machine: $($imachine.Name)"
           Write-Verbose "Controller: `"$($imediumattachment.Controller)`""
           Write-Verbose "Port: `"$($imediumattachment.Port)`""
           Write-Verbose "Device: `"$($imediumattachment.Device)`""
           $mmachine.ComObject.UnmountMedium($imediumattachment.Controller, $imediumattachment.Port, $imediumattachment.Device, $false)
           # save new settings
           Write-Verbose "Saving new settings"
           $mmachine.ComObject.SaveSettings()
           # unlock machine session
           Write-Verbose "Unlocking machine session"
           $imachine.ISession.Session.UnlockMachine()
          } # foreach $imediumattachment in $imediumattachments
         } # end elseif com
        } # end if $PSCmdlet.ShouldProcess(
       } # foreach $imachine in $imachines
      } # end if $imachines
     } # foreach $vmname in $imedium.MachineGuid
    } # end if $imedium.MachineGuid
   } # foreach $imedium in $imediums
  } # Try
  catch {
   Write-Verbose 'Exception dismounting virtual disc'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  finally {
   # release mutable machine objects if they exist
   if ($mmachine) {
    if ($mmachine.ISession.Id) {
     # release mutable session object
     Write-Verbose "Releasing mutable session object"
     $global:vbox.IManagedObjectRef_release($mmachine.ISession.Id)
    }
    if ($mmachine.ISession.Session) {
     if ($mmachine.ISession.Session.State -gt 1) {
      $mmachine.ISession.Session.UnlockMachine()
     } # end if $mmachine.ISession.Session locked
    } # end if $mmachine.ISession.Session
    if ($mmachine.Id) {
     # release mutable object
     Write-Verbose "Releasing mutable object"
     $global:vbox.IManagedObjectRef_release($mmachine.Id)
    }
   }
   # obligatory session unlock
   Write-Verbose 'Cleaning up machine sessions'
   if ($imachines) {
    foreach ($imachine in $imachines) {
     if ($imachine.ISession.Id) {
      if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
       Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
       $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
      } # end if session state not unlocked
     } # end if $imachine.ISession.Id
     if ($imachine.ISession.Session) {
      if ($imachine.ISession.Session.State -gt 1) {
       $imachine.ISession.Session.UnlockMachine()
      } # end if $imachine.ISession.Session locked
     } # end if $imachine.ISession.Session
     if ($imachine.IConsole) {
      # release the iconsole session
      Write-verbose "Releasing the IConsole session for VM $($imachine.Name)"
      $global:vbox.IManagedObjectRef_release($imachine.IConsole)
     } # end if $imachine.IConsole
     #$imachine.ISession.Id = $null
     $imachine.IConsole = $null
     if ($imachine.IPercent) {$imachine.IPercent = $null}
     $imachine.MSession = $null
     $imachine.MConsole = $null
     $imachine.MMachine = $null
    } # end foreach $imachine in $imachines
   } # end if $imachines
  } # Finally
 } # end if $imediums
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
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
.PARAMETER Timeout
The time before the process will timeout in milliseconds. Default value is 0 which will disable timeout period.
.PARAMETER Credential
Administrator/Root credentials for the machine.
.PARAMETER StdOut
A switch to send StdOut to the pipeline.
.PARAMETER StdErr
A switch to display StdErr to the screen.
.PARAMETER NoWait
A switch to skip waiting for process completion. StdOut and StdErr switches will be ignored if this switch is used. Warning: this will launch the process and immediately exit. If the process does not terminate successfully, you will need to do so manually from within the guest OS.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Submit-VirtualBoxVMProcess Win10 'cmd.exe' '/c','shutdown','/s','/f' -Credential $credentials
Runs cmd.exe in the Win10 virtual machine guest OS with the argument list "/c shutdown /s /f"
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Where-Object {$_.GuestOS -match 'windows'} | Submit-VirtualBoxVMProcess -PathToExecutable 'C:\\Windows\\System32\\gpupdate.exe' -Credential $credentials
Runs gpupdate.exe on all running virtual machines with a Windows guest OS
.NOTES
NAME        :  Submit-VirtualBoxVMProcess
VERSION     :  1.2
LAST UPDATED:  1/22/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Submit-VirtualBoxVMPowerShellScript
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for process to create
String[]      :  Strings for arguments to process
Uint32        :  uint32 for timeout in ms
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
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
Mandatory=$true,ParameterSetName="Name",Position=0)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
Mandatory=$true,ParameterSetName="Guid",Position=0)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(HelpMessage="Enter the full path to the executable",
Position=1,Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string]$PathToExecutable,
[Parameter(HelpMessage="Enter an array of arguments to use when creating the process",
Position=2)]
  [string[]]$Arguments,
[Parameter(HelpMessage="Enter the timeout in milliseconds",
Position=3,Mandatory=$false)]
  [uint32]$Timeout = 0,
[Parameter(Mandatory=$true,
HelpMessage="Enter the credentials to login to the guest OS")]
  [pscredential]$Credential,
[Parameter(HelpMessage="Use this switch to write StdOut to the pipeline")]
  [switch]$StdOut,
[Parameter(HelpMessage="Use this switch to write StdErr to the screen")]
  [switch]$StdErr,
[Parameter(HelpMessage="Use this switch ONLY when needed")]
  [switch]$NoWait,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 if ($Arguments) {$Arguments = $PathToExecutable,$Arguments}
 $command = "$($PathToExecutable) -- $($Arguments)"
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 try {
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($ModuleHost.ToLower() -eq 'websrv') {
     Write-verbose "Locking the machine session"
     $global:vbox.IMachine_lockMachine($imachine.Id, $imachine.ISession.Id, [LockType]::new().ToInt('Shared'))
     # create iconsole session to vm
     Write-verbose "Creating IConsole session to the machine"
     $imachine.IConsole = $global:vbox.ISession_getConsole($imachine.ISession.Id)
     # create iconsole guest session to vm
     Write-verbose "Creating IConsole guest session to the machine"
     $imachine.IConsoleGuest = $global:vbox.IConsole_getGuest($imachine.IConsole)
     # create a guest session
     Write-Verbose "Creating a guest console session"
     $imachine.IGuestSession = $global:vbox.IGuest_createSession($imachine.IConsoleGuest, $Credential.GetNetworkCredential().UserName, $Credential.GetNetworkCredential().Password, $Credential.GetNetworkCredential().Domain, "WsPsLaunchProcess_$([datetime]::Now)")
     # wait 10 seconds for the session to be created successfully
     Write-Verbose "Waiting for guest console to establish successfully (timeout: 10s)"
     $iguestsessionstatus = $global:vbox.IGuestSession_waitFor($imachine.IGuestSession, [GuestSessionWaitForFlag]::new().ToULong('Start'), 10000)
     Write-Verbose "Guest console status: $iguestsessionstatus"
     # create the process in the guest machine and send it a list of arguments
     Write-Verbose "Sending `"$command`" command (timeout: $($Timeout)ms)"
     if ($global:vbox.IGuestSession_fsObjExists($imachine.IGuestSession, $PathToExecutable, 1)) {
      $iguestprocess = $global:vbox.IGuestSession_processCreate($imachine.IGuestSession, $PathToExecutable, $Arguments, [array]@(), [array]@([ProcessCreateFlag]::new().ToInt('WaitForStdOut'), [ProcessCreateFlag]::new().ToInt('WaitForStdErr')), $Timeout)
     }
     else {Write-Host "[Error] Executable specified ($PathToExecutable) does not exist on the guest. Check the path and try again." -ForegroundColor Red -BackgroundColor Black;return}
     if (!$NoWait) {
      # create event source
      Write-Verbose "Creating event source"
      $ieventsource = $global:vbox.IConsole_getEventSource($imachine.IConsole)
      # create event listener
      Write-Verbose "Creating event listener"
      $ieventlistener = $global:vbox.IEventSource_createListener($ieventsource)
      # register event listener
      Write-Verbose "Registering event listener"
      $global:vbox.IEventSource_registerListener($ieventsource, $ieventlistener, [VBoxEventType]::new().ToInt('Any'), $false)
      try {
       # wait for process creation
       Write-Verbose "Waiting for guest process to be created (timeout: 10s)"
       $processwaitresult = $global:vbox.IProcess_waitFor($iguestprocess, [ProcessWaitForFlag]::new().ToULong('Start'), 10000)
       Write-Verbose "Process wait result: $($processwaitresult)"
       # gather extra data
       $guestprocessexecutablepath = $global:vbox.IProcess_getExecutablePath($iguestprocess)
       $guestprocesspid = $global:vbox.IProcess_getPID($iguestprocess)
       $guestprocessarguments = $global:vbox.IProcess_getArguments($iguestprocess)
       Write-Verbose "Launched process: `"$($guestprocessexecutablepath)`""
       Write-Verbose "Process PID: `"$($guestprocesspid)`""
       Write-Verbose "Guest syntax: `"$($guestprocessarguments)`""
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
        # this is returning WaitFlagNotSupported - waiting for stdout and stderr is not currently implemented - leaving this for when it does work and waitfor terminate still works
        $processwaitresult = $global:vbox.IProcess_waitForArray($iguestprocess, @([ProcessWaitForFlag]::new().ToULong('StdOut'),[ProcessWaitForFlag]::new().ToULong('StdErr'),[ProcessWaitForFlag]::new().ToULong('Terminate')), 200)
        if ($StdOut) {
         # read guest process stdout
         [char[]]$readstdout = $global:vbox.IProcess_read($iguestprocess, [Handle]::new().ToULong('StdOut'), (64 * 1024), 0)
         if ($readstdout) {
          # write stdout to pipeline
          Write-Verbose "Writing StdOut to pipeline"
          Write-Output ([Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($readstdout -join ''))).TrimEnd("`r`n")
         } # end if $readstdout
        } # end if $StdOut
        # write stderr to the host as error text if it contains anything
        if ($StdErr) {
         # read guest process stderr
         [char[]]$readstderr = $global:vbox.IProcess_read($iguestprocess, [Handle]::new().ToULong('StdErr'), (64 * 1024), 0)
         if ($readstderr) {
          # write stderr to pipeline
          Write-Verbose "Writing StdErr to pipeline"
          $origFgC = [System.Console]::ForegroundColor
          [System.Console]::ForegroundColor = 'Red'
          Write-Output ([Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($readstderr -join ''))).TrimEnd("`r`n")
          [System.Console]::ForegroundColor = $origFgC
         } # end if $readstderr
        } # end if $StdErr
        $iprocessstatus = $global:vbox.IProcess_getStatus($iguestprocess)
        # note the process status to look for abnormal return
        if ($iprocessstatus -notmatch 'Start') {
         if ($iprocessstatus -eq 'TerminatedNormally') {Write-Verbose 'Process terminated normally'}
         else {Write-Verbose "Process status: $($iprocessstatus)"}
        } # end if $iprocessstatus -notmatch 'Start'
        if ($processwaitresult -match 'Timeout') {Write-Verbose "Process timed out"}
       } until ($iprocessstatus.toString() -match 'Terminated' -or $processwaitresult -match 'Timeout')
      } # Try
      catch {
       Write-Verbose 'Exception while running process in guest machine'
       Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
       Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
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
     } # end if not NoWait
    } # end if websrv
    elseif ($ModuleHost.ToLower() -eq 'com') {
     Write-verbose "Locking the machine session"
     $imachine.ComObject.LockMachine($imachine.ISession.Session, [LockType]::new().ToInt('Shared'))
     # create a guest session
     Write-Verbose "Creating a guest console session"
     $imachine.ISession.Session.Console.Guest.CreateSession($Credential.GetNetworkCredential().UserName, $Credential.GetNetworkCredential().Password, $Credential.GetNetworkCredential().Domain, "ComPsLaunchProcess_$([datetime]::Now)") | Out-Null
     # wait 10 seconds for the session to be created successfully
     Write-Verbose "Waiting for guest console to establish successfully (timeout: 10s)"
     $iguestsessionstatus = $imachine.ISession.Session.Console.Guest.Sessions.WaitFor(1, 10000)
     Write-Verbose "Guest console status: $($iguestsessionstatus)"
     Write-Verbose "Guest console status: $($iguestsessionstatus.GetType())"
     if ($iguestsessionstatus.GetType().Name -eq 'Int32') {Write-Verbose "Guest console status: $([GuestSessionWaitForFlag]::new().ToStr($iguestsessionstatus))"}
     elseif ($iguestsessionstatus.GetType().Name -eq 'Object[]') {Write-Verbose "Guest console status: $([GuestSessionWaitForFlag]::new().ToStr($iguestsessionstatus.GetUpperBound(0)))"}
     # create the process in the guest machine and send it a list of arguments
     Write-Verbose "Sending `"$command`" command (timeout: $($Timeout)ms)"
     if ($imachine.ISession.Session.Console.Guest.Sessions.FsObjExists($PathToExecutable, 1) -eq 1) {
      if ($StdOut -and $StdErr) {$iguestprocess = $imachine.ISession.Session.Console.Guest.Sessions.ProcessCreate($PathToExecutable, [string[]]@($Arguments), [string[]]@(), [int[]]@([ProcessCreateFlag]::new().ToIntCom('WaitForStdOut'), [ProcessCreateFlag]::new().ToIntCom('WaitForStdErr')), $Timeout)}
      elseif ($StdOut) {$iguestprocess = $imachine.ISession.Session.Console.Guest.Sessions.ProcessCreate($PathToExecutable, [string[]]@($Arguments), [string[]]@(), [int[]]@([ProcessCreateFlag]::new().ToIntCom('WaitForStdOut')), $Timeout)}
      elseif ($StdErr) {$iguestprocess = $imachine.ISession.Session.Console.Guest.Sessions.ProcessCreate($PathToExecutable, [string[]]@($Arguments), [string[]]@(), [int[]]@([ProcessCreateFlag]::new().ToIntCom('WaitForStdErr')), $Timeout)}
      else {$iguestprocess = $imachine.ISession.Session.Console.Guest.Sessions.ProcessCreate($PathToExecutable, [string[]]@($Arguments), [string[]]@(), [int[]]@([ProcessCreateFlag]::new().ToIntCom('Hidden')), $Timeout)}
     }
     else {Write-Host "[Error] Executable specified ($PathToExecutable) does not exist on the guest. Check the path and try again." -ForegroundColor Red -BackgroundColor Black;return}
     if (!$NoWait) {
      # create event listener
      Write-Verbose "Creating event listener"
      $ieventlistener = $imachine.ISession.Session.Console.EventSource.CreateListener()
      # register event listener
      Write-Verbose "Registering event listener"
      $imachine.ISession.Session.Console.EventSource.RegisterListener($ieventlistener, [int[]]@([VBoxEventType]::new().ToInt('Any')), $false)
      try {
       # wait for process creation
       Write-Verbose "Waiting for guest process to be created (timeout: 10s)"
       $processwaitresult = $imachine.ISession.Session.Console.Guest.Sessions.Processes.WaitFor([ProcessWaitForFlag]::new().ToULong('Start'), 10000)
       Write-Verbose "Process wait result: $($processwaitresult)"
       # gather extra data
       $guestprocessexecutablepath = $imachine.ISession.Session.Console.Guest.Sessions.Processes.ExecutablePath
       $guestprocesspid = $imachine.ISession.Session.Console.Guest.Sessions.Processes.PID
       $guestprocessarguments = $imachine.ISession.Session.Console.Guest.Sessions.Processes.Arguments
       Write-Verbose "Launched process: `"$($guestprocessexecutablepath)`""
       Write-Verbose "Process PID: `"$($guestprocesspid)`""
       Write-Verbose "Guest syntax: `"$($guestprocessarguments)`""
       $ieventsublistener = $null
       do {
        # get new events
        $ievent = $imachine.ISession.Session.Console.EventSource.GetEvent($ieventlistener, 200)
        if ($ievent) {
         # process new event
         Write-Verbose "Encountered event: $($ievent)"
         $imachine.ISession.Session.Console.EventSource.EventProcessed($ieventlistener, $ievent)
        } # end if $ievent -ne ''
        # this is returning WaitFlagNotSupported - waiting for stdout and stderr is not currently implemented - leaving this for when it does work and waitfor terminate still works
        $processwaitresult = $imachine.ISession.Session.Console.Guest.Sessions.Processes.WaitForArray([int[]]@([ProcessWaitForFlag]::new().ToULong('StdOut'), [ProcessWaitForFlag]::new().ToULong('StdErr'), [ProcessWaitForFlag]::new().ToULong('Terminate')), 200)
        if ($StdOut) {
         # read guest process stdout
         [byte[]]$readstdout = $imachine.ISession.Session.Console.Guest.Sessions.Processes.Read([Handle]::new().ToULong('StdOut'), (64 * 1024), 0)
         if ($readstdout) {
          # write stdout to pipeline
          Write-Verbose "Writing StdOut to pipeline"
          Write-Output ([System.Text.Encoding]::ASCII.GetString($readstdout)).TrimEnd("`r`n")
         } # end if $readstdout
        } # end if $StdOut
        # write stderr to the host as error text if it contains anything
        if ($StdErr) {
         # read guest process stderr
         [byte[]]$readstderr = $imachine.ISession.Session.Console.Guest.Sessions.Processes.Read([Handle]::new().ToULong('StdErr'), (64 * 1024), 0)
         if ($readstderr) {
          # write stderr to pipeline
          Write-Verbose "Writing StdErr to pipeline"
          $origFgC = [System.Console]::ForegroundColor
          [System.Console]::ForegroundColor = 'Red'
          Write-Output ([System.Text.Encoding]::ASCII.GetString($readstderr)).TrimEnd("`r`n")
          [System.Console]::ForegroundColor = $origFgC
         } # end if $readstderr
        } # end if $StdErr
        $iprocessstatus = $imachine.ISession.Session.Console.Guest.Sessions.Processes.Status
        # note the process status to look for abnormal return
        if ($iprocessstatus.ToString() -notmatch '100') {
         if ($iprocessstatus.ToString() -match '500') {Write-Verbose 'Process terminated normally'}
         else {Write-Verbose "Process status: $($iprocessstatus)"}
        } # end if $iprocessstatus -notmatch 'Start'
        if ($processwaitresult -eq 5) {Write-Verbose "Process timed out"}
       } until ($iprocessstatus -gt 200 -or $processwaitresult -eq 5 -or $processwaitresult -eq 2)
      } # Try
      catch {
       Write-Verbose 'Exception while running process in guest machine'
       Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
       Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
      } # Catch
      finally {
       # unregister listener object
       Write-Verbose 'Unregistering listener'
       $imachine.ISession.Session.Console.EventSource.UnregisterListener($ieventlistener)
       if (!$imachine.ISession.Session.Console.Guest.Sessions.Processes.Status.toString().contains('Terminated')) {
        # kill guest process if it hasn't ended yet
        Write-Verbose 'Terminating guest process'
        $imachine.ISession.Session.Console.Guest.Sessions.Processes.Terminate()
       } # end if process hasn't terminated
      } # Finally
     } # end if not NoWait
    } # end elseif com
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Host "[Error] No matching virtual machines were found using specified parameters" -ForegroundColor Red -BackgroundColor Black;return}
 } # Try
 catch {
  Write-Verbose 'Exception running process in guest machine'
  Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
  Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
 } # Catch
 finally {
  # obligatory session unlock
  Write-Verbose 'Cleaning up machine sessions'
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if ($imachine.ISession.Id) {
     if ($global:vbox.ISession_getState($imachine.ISession.Id) -eq 'Locked') {
      Write-Verbose "Unlocking ISession for VM $($imachine.Name)"
      $global:vbox.ISession_unlockMachine($imachine.ISession.Id)
     } # end if session state not unlocked
    } # end if $imachine.ISession.Id
    if ($imachine.IConsole) {
     # release the iconsole session object
     Write-verbose "Releasing the IConsole session object for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsole)
    } # end if $imachine.IConsole
    # next 2 ifs only for in-guest sessions
    if ($imachine.IGuestSession -and !$NoWait) {
     # close the iconsole session
     Write-verbose "Closing the IGuestSession session for VM $($imachine.Name)"
     $global:vbox.IGuestSession_close($imachine.IGuestSession)
     # release the iconsole session
     Write-verbose "Releasing the IGuestSession object for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IGuestSession)
    } # end if $imachine.IConsole and not bypass
    if ($imachine.ISession.Session -and !$NoWait) {
     if ($imachine.ISession.Session.Console.Guest.Sessions) {
      $imachine.ISession.Session.Console.Guest.Sessions.Close()
     } # end if guest session
     if ($imachine.ISession.Session.State -gt 1) {
      $imachine.ISession.Session.UnlockMachine()
     } # end if machine session locked
    } # end if machine session and not bypass
    if ($imachine.IConsoleGuest) {
     # release the iconsole session
     Write-verbose "Releasing the IConsoleGuest object for VM $($imachine.Name)"
     $global:vbox.IManagedObjectRef_release($imachine.IConsoleGuest)
    } # end if $imachine.IConsole
    #$imachine.ISession.Id = $null
    $imachine.IConsole = $null
    if ($imachine.IPercent) {$imachine.IPercent = $null}
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
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
Function Submit-VirtualBoxVMPowerShellScript {
<#
.SYNOPSIS
Start a guest virtual machine PowerShell script
.DESCRIPTION
Will start the requested PowerShell script, with optional arguments, in the guest operating system.
.PARAMETER Machine
At least one running virtual machine object. Can be received via pipeline input.
.PARAMETER Name
The Name of at least one running virtual machine.
.PARAMETER GUID
The GUID of at least one running virtual machine.
.PARAMETER ScriptBlock
The PowerShell script to be run in the guest.
.PARAMETER Timeout
The time before the script will timeout in milliseconds. Default value is 0 which will disable timeout period.
.PARAMETER Credential
Administrator/Root credentials for the machine.
.PARAMETER StdOut
A switch to send StdOut to the pipeline.
.PARAMETER StdErr
A switch to display StdErr to the screen.
.PARAMETER NoWait
A switch to skip waiting for PowerShell completion. StdOut and StdErr switches will be ignored if this switch is used. Warning: this will launch the PowerShell scriptblock and immediately exit. If PowerShell does not terminate successfully, you will need to do so manually from within the guest OS.
.PARAMETER SkipCheck
A switch to skip web service update. (Only for use between functions - DO NOT USE)
.EXAMPLE
PS C:\> Submit-VirtualBoxVMPowerShellScript Win10 '& cmd.exe /c shutdown /s /f' -Credential $credentials
Runs '& cmd.exe /c shutdown /s /f' in the Win10 virtual machine guest OS with PowerShell
.EXAMPLE
PS C:\> Get-VirtualBoxVM -State Running | Where-Object {$_.GuestOS -match 'windows'} | Submit-VirtualBoxVMPowerShellScript -ScriptBlock '& cmd /c gpupdate.exe /force' -Credential $credentials
Runs '& cmd /c gpupdate.exe /force' with PowerShell on all running virtual machines with a Windows guest OS
.NOTES
NAME        :  Submit-VirtualBoxVMPowerShellScript
VERSION     :  1.1
LAST UPDATED:  1/22/2020
AUTHOR      :  Andrew Brehm
EDITOR      :  SmithersTheOracle
.LINK
Submit-VirtualBoxVMPowerShellScript
.INPUTS
VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
String[]      :  Strings for virtual machine names
Guid[]        :  GUIDs for virtual machine GUIDs
String        :  String for scriptblock to be run
Uint32        :  uint32 for timeout in ms
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
  [VirtualBoxVM[]]$Machine,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine name(s)",
Mandatory=$true,ParameterSetName="Name",Position=0)]
[ValidateNotNullorEmpty()]
  [string[]]$Name,
[Parameter(ValueFromPipelineByPropertyName=$true,
HelpMessage="Enter one or more virtual machine GUID(s)",
Mandatory=$true,ParameterSetName="Guid",Position=0)]
[ValidateNotNullorEmpty()]
  [guid[]]$Guid,
[Parameter(Position=1,Mandatory=$true)]
[ValidateNotNullorEmpty()]
  [string]$ScriptBlock,
[Parameter(HelpMessage="Enter the timeout in milliseconds",
Position=2,Mandatory=$false)]
  [uint32]$Timeout = 0,
[Parameter(Mandatory=$true,
HelpMessage="Enter the credentials to login to the guest OS")]
  [pscredential]$Credential,
[Parameter(HelpMessage="Use this switch to write StdOut to the pipeline")]
  [switch]$StdOut,
[Parameter(HelpMessage="Use this switch to write StdErr to the screen")]
  [switch]$StdErr,
[Parameter(HelpMessage="Use this switch ONLY when needed")]
  [switch]$NoWait,
[Parameter(HelpMessage="Use this switch to skip service update (for development use)")]
  [switch]$SkipCheck
) # Param
Begin {
 Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 if ($ModuleHost.ToLower() -eq 'websrv') {
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  if (!$global:ivbox) {Start-VirtualBoxSession}
 } # end if websrv
} # Begin
Process {
 Write-Verbose "Pipeline - Machine: `"$Machine`""
 Write-Verbose "Pipeline - Name: `"$Name`""
 Write-Verbose "Pipeline - Guid: `"$Guid`""
 Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
 if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
 # initialize $imachines array
 $imachines = @()
 if ($Machine) {
  Write-Verbose "Getting VM inventory from Machine(s)"
  $imachines = $Machine
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Machine)
 elseif ($Name) {
  foreach ($item in $Name) {
   Write-Verbose "Getting VM inventory from Name(s)"
   $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Name)
 elseif ($Guid) {
  foreach ($item in $Guid) {
   Write-Verbose "Getting VM inventory from GUID(s)"
   $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
  }
  $imachines = $imachines | Where-Object {$_ -ne $null}
 }# get vm inventory (by $Guid)
 if ($imachines) {
  foreach ($imachine in $imachines) {
   Write-Verbose "Submitting PowerShell command to the $($imachine.Name) machine"
   if ($NoWait) {Submit-VirtualBoxVMProcess -Machine $imachine -PathToExecutable "cmd.exe" -Arguments "/c","powershell","-ExecutionPolicy","Bypass","-Command",$ScriptBlock -Credential $Credential -Timeout $Timeout -NoWait -SkipCheck}
   elseif ($StdOut -and $StdErr) {Submit-VirtualBoxVMProcess -Machine $imachine -PathToExecutable "cmd.exe" -Arguments "/c","powershell","-ExecutionPolicy","Bypass","-Command",$ScriptBlock -Credential -Timeout $Timeout $Credential -StdOut -StdErr -SkipCheck}
   elseif ($StdOut) {Submit-VirtualBoxVMProcess -Machine $Machine -PathToExecutable "cmd.exe" -Arguments "/c","powershell","-ExecutionPolicy","Bypass","-Command",$ScriptBlock -Credential $Credential -Timeout $Timeout -StdOut -SkipCheck}
   elseif ($StdErr) {Submit-VirtualBoxVMProcess -Machine $Machine -PathToExecutable "cmd.exe" -Arguments "/c","powershell","-ExecutionPolicy","Bypass","-Command",$ScriptBlock -Credential $Credential -Timeout $Timeout -StdErr -SkipCheck}
   else {Submit-VirtualBoxVMProcess -Machine $Machine -PathToExecutable "cmd.exe" -Arguments "/c","powershell","-ExecutionPolicy","Bypass","-Command",$ScriptBlock -Credential $Credential -Timeout $Timeout -SkipCheck}
  }
 } # end if $imachines
 else {Write-Host "[Error] No matching virtual machines were found using specified parameters" -ForegroundColor Red -BackgroundColor Black;return}
} # Process
End {
 Write-Verbose "Ending $($MyInvocation.MyCommand)"
} # End
} # end function
if ($ModuleHost.ToLower() -eq 'websrv') {
 Function Start-VirtualBoxSession {
 <#
 .SYNOPSIS
 Starts a VirtualBox Web Service session and populates the $global:ivbox managed object reference
 .DESCRIPTION
 Create a PowerShell managed object reference to the VirtualBox Web Service managed object.
 .PARAMETER Protocol
 The protocol of the VirtualBox Web Service. Default is http.
 .PARAMETER Address
 The domain name or IP address of the VirtualBox Web Service. Default is localhost.
 .PARAMETER Port
 The TCP port of the VirtualBox Web Service. Default is 18083.
 .PARAMETER Force
 A switch to force updating global properties.
 .EXAMPLE
 PS C:\> Start-VirtualBoxSession -Protocol "http" -Address "localhost" -Port "18083" -Credential $Credential
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
   [string]$Address = "localhost",
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
  Write-Verbose "Beginning $($MyInvocation.MyCommand)"
  # refresh vboxwebsrv variable
  if (!$SkipCheck -or !(Get-Process 'VBoxWebSrv')) {$global:vboxwebsrvtask = Update-VirtualBoxWebSrv}
  # start the websrvtask if it's not running
  if ($global:vboxwebsrvtask.Status -and $global:vboxwebsrvtask.Status -ne 'Running') {Start-VirtualBoxWebSrv}
  # set the target web service url
  $global:vbox.Url = "$($Protocol)://$($Address):$($Port)"
  # save the host address
  $global:hostaddress = $Address
  # if a session already exists, stop it
  if ($global:ivbox) {Stop-VirtualBoxSession}
 } # Begin
 Process {
  try {
   # login to web service
   Write-Verbose 'Creating the VirtualBox Web Service session ($global:ivbox)'
   $global:ivbox = $global:vbox.IWebsessionManager_logon($Credential.GetNetworkCredential().UserName,$Credential.GetNetworkCredential().Password)
   $apiversion = ($vbox.IVirtualBox_getAPIVersion($ivbox)).Replace('_','.')
   if ($apiversion -lt 6.1) {Write-Host "[Error] Minimum VirtualBox API version required for this module is `"6.1`". Installed version is `"$apiversion`"." -ForegroundColor Yellow -BackgroundColor Black;return}
   if ($global:ivbox) {
    if (!$global:guestostype -or $Force) {
     try {
      # get guest OS type IDs
      Write-Verbose 'Fetching guest OS type data ($global:guestostype)'
      if ($global:guestostype) {if ((Get-Variable -Name guestostype -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name guestostype -Option None -Scope Global -Force}}
      $global:guestostype = $global:vbox.IVirtualBox_getGuestOSTypes($global:ivbox)
      if ((Get-Variable -Name guestostype -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name guestostype -Description "A list of VirtualBox guest OS type codes." -Option ReadOnly -Scope Global -Force}
     } # Try
     catch {
      Write-Verbose 'Exception fetching guest OS type data'
      Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
      Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
     } # Catch
    }
    if (!$global:isystemproperties -or $Force) {
     try {
      # create a local copy of capabilities for quick reference
      Write-Verbose 'Fetching system properties object ($global:isystemproperties)'
      $global:isystemproperties = $global:vbox.IVirtualBox_getSystemProperties($global:ivbox)
     } # Try
     catch {
      Write-Verbose 'Exception fetching system properties'
      Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
      Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
     } # Catch
     <#
     try {
      # get a local copy of device types by storage bus
      Write-Verbose 'Fetching device types by storage bus ($global:devicetypesforstoragebus)'
      $global:devicetypesforstoragebus = $global:vbox.ISystemProperties_getDeviceTypesForStorageBus($global:isystemproperties, 1)
      for ($i=2;$i-lt8;$i++) {
       $global:devicetypesforstoragebus += $global:vbox.ISystemProperties_getDeviceTypesForStorageBus($global:isystemproperties, $i)
      }
     }
     catch {
      Write-Verbose 'Exception fetching device types for storage buses'
      Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
      Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
     }
     #> # disabling since this isn't really useful data
     try {
      Write-Verbose 'Fetching supported system properties ($global:systempropertiessupported)'
      if ($global:systempropertiessupported) {if ((Get-Variable -Name systempropertiessupported -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name systempropertiessupported -Option None -Scope Global -Force}}
      $global:systempropertiessupported.Fetch()
      if ((Get-Variable -Name systempropertiessupported -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name systempropertiessupported -Description "A list of VirtualBox properties currently supported by your system." -Option ReadOnly -Scope Global -Force}
     } # Try
     catch {
      Write-Verbose 'Exception fetching supported system properties'
      Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
      Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
     } # Catch
     try {
      Write-Verbose 'Fetching supported medium formats ($global:mediumformats)'
      if ($global:mediumformats) {if ((Get-Variable -Name mediumformats -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name mediumformats -Option None -Scope Global -Force}}
      $global:mediumformats.Fetch()
      if ((Get-Variable -Name mediumformats -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name mediumformats -Description "A list of medium formats." -Option ReadOnly -Scope Global -Force}
      # get a human readable copy
      Write-Verbose 'Fetching medium format PSO ($global:mediumformatspso)'
      if ($global:mediumformatspso) {if ((Get-Variable -Name mediumformatspso -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name mediumformatspso -Option None -Scope Global -Force}}
      $global:mediumformatspso = $mediumformatspso.FetchObject($global:vbox.ISystemProperties_getMediumFormats($global:isystemproperties))
      if ((Get-Variable -Name mediumformatspso -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name mediumformatspso -Description "A list of medium formats in PowerShell-readable format." -Option ReadOnly -Scope Global -Force}
     } # Try
     catch {
      Write-Verbose 'Exception fetching supported medium formats'
      Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
      Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
     } # Catch
    }
   }
  }
  catch {
   Write-Verbose 'Exception creating the VirtualBox Web Service session'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  }
 } # Process
 End {
  Write-Verbose "Ending $($MyInvocation.MyCommand)"
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
  Write-Verbose "Beginning $($MyInvocation.MyCommand)"
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
    Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
    Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
   }
  }
 } # Process
 End {
  Write-Verbose "Ending $($MyInvocation.MyCommand)"
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
  Write-Verbose "Beginning $($MyInvocation.MyCommand)"
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
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  }
 } # Process
 End {
  Write-Verbose "Ending $($MyInvocation.MyCommand)"
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
  Write-Verbose "Beginning $($MyInvocation.MyCommand)"
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
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # end catch
 } # Process
 End {
  Write-Verbose "Ending $($MyInvocation.MyCommand)"
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
  Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 } # Begin
 Process {
  # restart the web service task
  Stop-VirtualBoxWebSrv
  Start-VirtualBoxWebSrv
 } # Process
 End {
  Write-Verbose "Ending $($MyInvocation.MyCommand)"
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
  Write-Verbose "Beginning $($MyInvocation.MyCommand)"
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
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # end catch
  if (!$vboxwebsrvtask) {Write-Host '[Error] Failed to update $vboxwebsrvtask' -ForegroundColor Red -BackgroundColor Black;return}
  return $vboxwebsrvtask
 } # Process
 End {
  Write-Verbose $vboxwebsrvtask
  Write-Verbose "Ending $($MyInvocation.MyCommand)"
 } # End
 } # end function
} # end if websrv
elseif ($ModuleHost.ToLower() -eq 'com') {
 Function Show-VirtualBoxVM {
 <#
 .SYNOPSIS
 Display a virtual machine console window
 .DESCRIPTION
 Displays a virtual machine console window and powers it on if needed.
 .PARAMETER Machine
 At least one virtual machine object. Can be received via pipeline input.
 .PARAMETER Name
 The name of at least one virtual machine. Can be received via pipeline input by name.
 .PARAMETER Guid
 The GUID of at least one virtual machine. Can be received via pipeline input by name.
 .EXAMPLE
 PS C:\> Get-VirtualBoxVM -State Running | Show-VirtualBoxVM
 Display a console window for all running virtual machines
 .EXAMPLE
 PS C:\> Show-VirtualBoxVM -Name "2016"
 Display a console window for the "2016 Core" virtual machine
 .EXAMPLE
 PS C:\> Show-VirtualBoxVM -Guid 7353caa6-8cb6-4066-aec9-6c6a69a001b6
 Display a console window for the virtual machine with GUID 7353caa6-8cb6-4066-aec9-6c6a69a001b6
 .NOTES
 NAME        :  Show-VirtualBoxVM
 VERSION     :  1.0
 LAST UPDATED:  1/24/2020
 AUTHOR      :  Andrew Brehm
 EDITOR      :  SmithersTheOracle
 .LINK
 None
 .INPUTS
 VirtualBoxVM[]:  VirtualBoxVMs for virtual machine objects
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
   [VirtualBoxVM[]]$Machine,
 [Parameter(ValueFromPipelineByPropertyName=$true,
 HelpMessage="Enter one or more virtual machine name(s)",
 ParameterSetName="Name",Mandatory=$true)]
 [ValidateNotNullorEmpty()]
   [string[]]$Name,
 [Parameter(ValueFromPipelineByPropertyName=$true,
 HelpMessage="Enter one or more virtual machine GUID(s)",
 ParameterSetName="Guid",Mandatory=$true)]
 [ValidateNotNullorEmpty()]
   [guid[]]$Guid
 ) # Param
 Begin {
  Write-Verbose "Beginning $($MyInvocation.MyCommand)"
 } # Begin
 Process {
  Write-Verbose "Pipeline - Machine: `"$Machine`""
  Write-Verbose "Pipeline - Name: `"$Name`""
  Write-Verbose "Pipeline - Guid: `"$Guid`""
  Write-Verbose "ParameterSetName: `"$($PSCmdlet.ParameterSetName)`""
  if (!($Machine -or $Name -or $Guid)) {Write-Host "[Error] You must supply at least one VM object, name, or GUID." -ForegroundColor Red -BackgroundColor Black;return}
  # initialize $imachines array
  $imachines = @()
  if ($Machine) {
   Write-Verbose "Getting VM inventory from Machine(s)"
   $imachines = $Machine
   $imachines = $imachines | Where-Object {$_ -ne $null}
  }# get vm inventory (by $Machine)
  elseif ($Name) {
   foreach ($item in $Name) {
    Write-Verbose "Getting VM inventory from Name(s)"
    $imachines += Get-VirtualBoxVM -Name $item -SkipCheck
   }
   $imachines = $imachines | Where-Object {$_ -ne $null}
  }# get vm inventory (by $Name)
  elseif ($Guid) {
   foreach ($item in $Guid) {
    Write-Verbose "Getting VM inventory from GUID(s)"
    $imachines += Get-VirtualBoxVM -Guid $item -SkipCheck
   }
   $imachines = $imachines | Where-Object {$_ -ne $null}
  }# get vm inventory (by $Guid)
  if ($imachines) {
   foreach ($imachine in $imachines) {
    if (Test-Path "$($env:VBOX_MSI_INSTALL_PATH)VBoxSDL.exe") {
     Write-Verbose "Launching console window for `"$($imachine.Name)`""
     Start-Process -FilePath "$($env:VBOX_MSI_INSTALL_PATH)VBoxSDL.exe" -ArgumentList ("--startvm `"$($imachine.Name)`" --separate") -WindowStyle Hidden
    } # end if VBoxSDL.exe exists
    else {Write-Host "[Error] VBoxSDL.exe not found. Ensure VirtualBox is installed on this machine and try again.";return}
   } # foreach $imachine in $imachines
  } # end if $imachines
  else {Write-Verbose "[Warning] No matching virtual machines were found using specified parameters"}
 } # Process
 End {
  Write-Verbose "Ending $($MyInvocation.MyCommand)"
 } # End
 } # end function
} # end elseif com
#########################################################################################
# Entry
Write-Verbose "Initializing VirtualBox environment"
if ($ModuleHost.ToLower() -eq 'websrv') {
 if (!(Get-Process 'VBoxWebSrv')) {
  if (Test-Path "$($env:VBOX_MSI_INSTALL_PATH)VBoxWebSrv.exe") {
   Start-VirtualBoxWebSrv
  }
  else {Write-Host "[Error] VBoxWebSrv not found." -ForegroundColor Red -BackgroundColor Black;return}
 } # end if VBoxWebSrv check
 # get the global reference to the virtualbox web service object
 Write-Verbose 'Creating the VirtualBox Web Service object ($global:vbox)'
 if (!$vbox) {$global:vbox = New-WebServiceProxy -Uri "$($env:VBOX_MSI_INSTALL_PATH)sdk\bindings\webservice\vboxwebService.wsdl" -Namespace "VirtualBox" -Class "VirtualBoxWebSrv"}
 if ($WebSrvCredential) {Start-VirtualBoxSession -Protocol $WebSrvProtocol -Address $WebSrvAddress -Port $WebSrvPort -Credential $WebSrvCredential}
 if ($ivbox) {
  try {
   # get guest OS type IDs
   Write-Verbose 'Fetching guest OS type data ($global:guestostype)'
   if ($guestostype) {if ((Get-Variable -Name guestostype -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name guestostype -Option None -Scope Global -Force}}
   $global:guestostype = $vbox.IVirtualBox_getGuestOSTypes($ivbox)
   if ((Get-Variable -Name guestostype -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name guestostype -Description "A list of VirtualBox guest OS type codes." -Option ReadOnly -Scope Global -Force}
  } # Try
  catch {
   Write-Verbose 'Exception fetching guest OS type data'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  try {
   # get system properties interface reference
   Write-Verbose 'Fetching system properties object ($global:isystemproperties)'
   $global:isystemproperties = $vbox.IVirtualBox_getSystemProperties($ivbox)
  } # Try
  catch {
   Write-Verbose 'Exception fetching system properties'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  try {
   Write-Verbose 'Fetching supported system properties ($global:systempropertiessupported)'
   if ($systempropertiessupported) {if ((Get-Variable -Name systempropertiessupported -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name systempropertiessupported -Option None -Scope Global -Force}}
   $global:systempropertiessupported.Fetch()
   if ((Get-Variable -Name systempropertiessupported -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name systempropertiessupported -Description "A list of VirtualBox properties currently supported by your system." -Option ReadOnly -Scope Global -Force}
  } # Try
  catch {
   Write-Verbose 'Exception fetching supported system properties'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  try {
   Write-Verbose 'Fetching supported medium formats ($global:mediumformats)'
   if ($mediumformats) {if ((Get-Variable -Name mediumformats -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name mediumformats -Option None -Scope Global -Force}}
   $global:mediumformats.Fetch()
   if ((Get-Variable -Name mediumformats -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name mediumformats -Description "A list of medium formats." -Option ReadOnly -Scope Global -Force}
   # get a human readable copy
   Write-Verbose 'Fetching medium format PSO ($global:mediumformatspso)'
   if ($mediumformatspso) {if ((Get-Variable -Name mediumformatspso -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name mediumformatspso -Option None -Scope Global -Force}}
   $global:mediumformatspso = $global:mediumformatspso.FetchObject($vbox.ISystemProperties_getMediumFormats($isystemproperties))
   if ((Get-Variable -Name mediumformatspso -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name mediumformatspso -Description "A list of medium formats in PowerShell-readable format." -Option ReadOnly -Scope Global -Force}
  } # Try
  catch {
   Write-Verbose 'Exception fetching supported medium formats'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
 } # end if $ivbox
 # get the web service task
 Write-Verbose "Updating VirtualBox WebSrv status"
 $vboxwebsrvtask = Update-VirtualBoxWebSrv
} # end if websrv
elseif ($ModuleHost.ToLower() -eq 'com') {
 $ErrorActionPreference = 'SilentlyContinue'
 # create vbox app
 Write-Verbose 'Creating the VirtualBox COM object ($global:vbox)'
 $vbox = New-Object -ComObject "VirtualBox.VirtualBox"
 $ErrorActionPreference = 'Continue'
 if ($vbox) {
  if ($vbox.APIVersion.Replace('_','.') -lt 6.1) {Write-Host "[Error] Minimum VirtualBox API version required for this module is `"6.1`". Installed version is `"$($vbox.APIVersion.Replace('_','.'))`"." -ForegroundColor Yellow -BackgroundColor Black;return}
  try {
   # get guest OS type IDs
   Write-Verbose 'Fetching guest OS type data ($global:guestostype)'
   if ($guestostype) {if ((Get-Variable -Name guestostype -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name guestostype -Option None -Scope Global -Force}}
   $global:guestostype = $vbox.GuestOSTypes
   if ((Get-Variable -Name guestostype -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name guestostype -Description "A list of VirtualBox guest OS type codes." -Option ReadOnly -Scope Global -Force}
  } # Try
  catch {
   Write-Verbose 'Exception fetching guest OS type data'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  try {
   Write-Verbose 'Fetching supported system properties ($global:systempropertiessupported)'
   if ($systempropertiessupported) {if ((Get-Variable -Name systempropertiessupported -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name systempropertiessupported -Option None -Scope Global -Force}}
   foreach ($SupportedParavirtProvider in $vbox.SystemProperties.SupportedParavirtProviders) {$global:systempropertiessupported.ParavirtProviders += [ParavirtProvider]::new().ToStr($SupportedParavirtProvider)}
   foreach ($SupportedClipboardMode in $vbox.SystemProperties.SupportedClipboardModes) {$global:systempropertiessupported.ClipboardModes += [ClipboardMode]::new().ToStr($SupportedClipboardMode)}
   foreach ($SupportedDnDMode in $vbox.SystemProperties.SupportedDnDModes) {$global:systempropertiessupported.DndModes += [DnDMode]::new().ToStr($SupportedDnDMode)}
   foreach ($SupportedFirmwareType in $vbox.SystemProperties.SupportedFirmwareTypes) {$global:systempropertiessupported.FirmwareTypes += [FirmwareType]::new().ToStr($SupportedFirmwareType)}
   foreach ($SupportedPointingHIDType in $vbox.SystemProperties.SupportedPointingHIDTypes) {$global:systempropertiessupported.PointingHidTypes += [PointingHIDType]::new().ToStr($SupportedPointingHIDType)}
   foreach ($SupportedKeyboardHIDType in $vbox.SystemProperties.SupportedKeyboardHIDTypes) {$global:systempropertiessupported.KeyboardHidTypes += [KeyboardHIDType]::new().ToStr($SupportedKeyboardHIDType)}
   foreach ($SupportedVFSType in $vbox.SystemProperties.SupportedVFSTypes) {$global:systempropertiessupported.VfsTypes += [VFSType]::new().ToStr($SupportedVFSType)}
   foreach ($SupportedImportOption in $vbox.SystemProperties.SupportedImportOptions) {$global:systempropertiessupported.ImportOptions += [ImportOptions]::new().ToStr($SupportedImportOption)}
   foreach ($SupportedExportOption in $vbox.SystemProperties.SupportedExportOptions) {$global:systempropertiessupported.ExportOptions += [ExportOptions]::new().ToStr($SupportedExportOption)}
   foreach ($SupportedRecordingAudioCodec in $vbox.SystemProperties.SupportedRecordingAudioCodecs) {$global:systempropertiessupported.RecordingAudioCodecs += [RecordingAudioCodec]::new().ToStr($SupportedRecordingAudioCodec)}
   foreach ($SupportedRecordingVideoCodec in $vbox.SystemProperties.SupportedRecordingVideoCodecs) {$global:systempropertiessupported.RecordingVideoCodecs += [RecordingVideoCodec]::new().ToStr($SupportedRecordingVideoCodec)}
   foreach ($SupportedRecordingVSMethod in $vbox.SystemProperties.SupportedRecordingVSMethods) {$global:systempropertiessupported.RecordingVsMethods += [RecordingVideoScalingMethod]::new().ToStr($SupportedRecordingVSMethod)}
   foreach ($SupportedRecordingVRCMode in $vbox.SystemProperties.SupportedRecordingVRCModes) {$global:systempropertiessupported.RecordingVrcModes += [RecordingVideoRateControlMode]::new().ToStr($SupportedRecordingVRCMode)}
   foreach ($SupportedGraphicsControllerType in $vbox.SystemProperties.SupportedGraphicsControllerTypes) {$global:systempropertiessupported.GraphicsControllerTypes += [GraphicsControllerType]::new().ToStr($SupportedGraphicsControllerType)}
   foreach ($SupportedCloneOption in $vbox.SystemProperties.SupportedCloneOptions) {$global:systempropertiessupported.CloneOptions += [CloneOptions]::new().ToStr($SupportedCloneOption)}
   foreach ($SupportedAutostopType in $vbox.SystemProperties.SupportedAutostopTypes) {$global:systempropertiessupported.AutostopTypes += [AutostopType]::new().ToStr($SupportedAutostopType)}
   foreach ($SupportedVMProcPriority in $vbox.SystemProperties.SupportedVMProcPriorities) {$global:systempropertiessupported.VmProcPriorities += [VMProcPriority]::new().ToStr($SupportedVMProcPriority)}
   foreach ($SupportedNetworkAttachmentType in $vbox.SystemProperties.SupportedNetworkAttachmentTypes) {$global:systempropertiessupported.NetworkAttachmentTypes += [NetworkAttachmentType]::new().ToStr($SupportedNetworkAttachmentType)}
   foreach ($SupportedNetworkAdapterType in $vbox.SystemProperties.SupportedNetworkAdapterTypes) {$global:systempropertiessupported.NetworkAdapterTypes += [NetworkAdapterType]::new().ToStr($SupportedNetworkAdapterType)}
   foreach ($SupportedPortMode in $vbox.SystemProperties.SupportedPortModes) {$global:systempropertiessupported.PortModes += [PortMode]::new().ToStr($SupportedPortMode)}
   foreach ($SupportedUartType in $vbox.SystemProperties.SupportedUartTypes) {$global:systempropertiessupported.UartTypes += [UartType]::new().ToStr($SupportedUartType)}
   foreach ($SupportedUSBControllerType in $vbox.SystemProperties.SupportedUSBControllerTypes) {$global:systempropertiessupported.UsbControllerTypes += [USBControllerType]::new().ToStr($SupportedUSBControllerType)}
   foreach ($SupportedAudioDriverType in $vbox.SystemProperties.SupportedAudioDriverTypes) {$global:systempropertiessupported.AudioDriverTypes += [AudioDriverType]::new().ToStr($SupportedAudioDriverType)}
   foreach ($SupportedAudioControllerType in $vbox.SystemProperties.SupportedAudioControllerTypes) {$global:systempropertiessupported.AudioControllerTypes += [AudioControllerType]::new().ToStr($SupportedAudioControllerType)}
   foreach ($SupportedStorageBus in $vbox.SystemProperties.SupportedStorageBuses) {$global:systempropertiessupported.StorageBuses += [StorageBus]::new().ToStr($SupportedStorageBus)}
   foreach ($SupportedStorageControllerType in $vbox.SystemProperties.SupportedStorageControllerTypes) {$global:systempropertiessupported.StorageControllerTypes += [StorageControllerType]::new().ToStr($SupportedStorageControllerType)}
   foreach ($SupportedChipsetType in $vbox.SystemProperties.SupportedChipsetTypes) {$global:systempropertiessupported.ChipsetTypes += [ChipsetType]::new().ToStr($SupportedChipsetType)}
   $global:systempropertiessupported.MinGuestRam = $vbox.SystemProperties.MinGuestRAM
   $global:systempropertiessupported.MaxGuestRam = $vbox.SystemProperties.MaxGuestRAM
   $global:systempropertiessupported.MinGuestVRam = $vbox.SystemProperties.MinGuestVRAM
   $global:systempropertiessupported.MaxGuestVRam = $vbox.SystemProperties.MaxGuestVRAM
   $global:systempropertiessupported.MinGuestCPUCount = $vbox.SystemProperties.MinGuestCPUCount
   $global:systempropertiessupported.MaxGuestCPUCount = $vbox.SystemProperties.MaxGuestCPUCount
   if ((Get-Variable -Name systempropertiessupported -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name systempropertiessupported -Description "A list of VirtualBox properties currently supported by your system." -Option ReadOnly -Scope Global -Force}
  } # Try
  catch {
   Write-Verbose 'Exception fetching supported system properties'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
  try {
   Write-Verbose 'Fetching supported medium formats ($global:mediumformats)'
   if ($mediumformats) {if ((Get-Variable -Name mediumformats -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name mediumformats -Option None -Scope Global -Force}}
   $global:mediumformats = Get-Content -Raw -Path "C:\Program Files\Oracle\VirtualBox\sdk\MediumFormat.json" | ConvertFrom-Json
   if ((Get-Variable -Name mediumformats -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name mediumformats -Description "A list of medium formats." -Option ReadOnly -Scope Global -Force}
   # get a human readable copy
   Write-Verbose 'Fetching medium format PSO ($global:mediumformatspso)'
   if ($mediumformatspso) {if ((Get-Variable -Name mediumformatspso -Scope Global).Options -match 'ReadOnly') {Set-Variable -Name mediumformatspso -Option None -Scope Global -Force}}
   $global:mediumformatspso = Get-Content -Raw -Path "C:\Program Files\Oracle\VirtualBox\sdk\MediumFormatPso.json" | ConvertFrom-Json
   if ((Get-Variable -Name mediumformatspso -Scope Global).Options -notmatch 'ReadOnly') {Set-Variable -Name mediumformatspso -Description "A list of medium formats in PowerShell-readable format." -Option ReadOnly -Scope Global -Force}
  } # Try
  catch {
   Write-Verbose 'Exception fetching supported medium formats'
   Write-Verbose "Stack trace output: $($_.ScriptStackTrace)"
   Write-Host "[Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
  } # Catch
 } # end if $vbox
} # end elseif com
if (!$vbox) {Write-Host '[Error] The VirtualBoxPS module did not load successfully.' -ForegroundColor Red -BackgroundColor Black;return (Remove-Module -Name VirtualBoxPS -Force)}
# define aliases
New-Alias -Name gvboxvm -Value Get-VirtualBoxVM
New-Alias -Name ssvboxvm -Value Suspend-VirtualBoxVM
New-Alias -Name ruvboxvm -Value Resume-VirtualBoxVM
New-Alias -Name savboxvm -Value Start-VirtualBoxVM
New-Alias -Name spvboxvm -Value Stop-VirtualBoxVM
New-Alias -Name nvboxvm -Value New-VirtualBoxVM
New-Alias -Name rvboxvm -Value Remove-VirtualBoxVM
New-Alias -Name ipvboxvm -Value Import-VirtualBoxVM
New-Alias -Name edvboxvm -Value Edit-VirtualBoxVM
New-Alias -Name gvboxvmsc -Value Get-VirtualBoxVMStorageController
New-Alias -Name nvboxvmsc -Value New-VirtualBoxVMStorageController
New-Alias -Name nvboxvmod -Value New-VirtualBoxVMOpticalDrive
New-Alias -Name rvboxvmod -Value Remove-VirtualBoxVMOpticalDrive
New-Alias -Name svboxvmgp -Value Set-VirtualBoxVMGuestProperty
New-Alias -Name rvboxvmgp -Value Remove-VirtualBoxVMGuestProperty
New-Alias -Name evboxvmvrde -Value Enable-VirtualBoxVMVRDEServer
New-Alias -Name dvboxvmvrde -Value Disable-VirtualBoxVMVRDEServer
New-Alias -Name edvboxvmvrde -Value Edit-VirtualBoxVMVRDEServer
New-Alias -Name ccvboxvmvrde -Value Connect-VirtualBoxVMVRDEServer
New-Alias -Name ipvboxovf -Value Import-VirtualBoxOVF
New-Alias -Name epvboxovf -Value Export-VirtualBoxOVF
New-Alias -Name gvboxd -Value Get-VirtualBoxDisk
New-Alias -Name nvboxd -Value New-VirtualBoxDisk
New-Alias -Name ipvboxd -Value Import-VirtualBoxDisk
New-Alias -Name rvboxd -Value Remove-VirtualBoxDisk
New-Alias -Name mtvboxd -Value Mount-VirtualBoxDisk
New-Alias -Name dmvboxd -Value Dismount-VirtualBoxDisk
New-Alias -Name edvboxd -Value Edit-VirtualBoxDisk
New-Alias -Name gvboxdvd -Value Get-VirtualBoxDisc
New-Alias -Name ipvboxdvd -Value Import-VirtualBoxDisc
New-Alias -Name rvboxdvd -Value Remove-VirtualBoxDisc
New-Alias -Name mtvboxdvd -Value Mount-VirtualBoxDisc
New-Alias -Name dmvboxdvd -Value Dismount-VirtualBoxDisc
New-Alias -Name sbvboxvmp -Value Submit-VirtualBoxVMProcess
New-Alias -Name sbvboxvmpss -Value Submit-VirtualBoxVMPowerShellScript
if ($ModuleHost.ToLower() -eq 'websrv') {
 New-Alias -Name savboxs -Value Start-VirtualBoxSession
 New-Alias -Name spvboxs -Value Stop-VirtualBoxSession
 New-Alias -Name savboxws -Value Start-VirtualBoxWebSrv
 New-Alias -Name spvboxws -Value Stop-VirtualBoxWebSrv
 New-Alias -Name rtvboxws -Value Restart-VirtualBoxWebSrv
 New-Alias -Name udvboxws -Value Update-VirtualBoxWebSrv
} # end if websrv
elseif ($ModuleHost.ToLower() -eq 'com') {
 New-Alias -Name shvboxvm -Value Show-VirtualBoxVM
} # end elseif com
# export module members
Export-ModuleMember -Alias * -Function * -Variable @('vbox','vboxerror')