#include-once
#include <WinAPI.au3>

; Constants
;~ Global Const $FILE_SHARE_READ = 0x1
;~ Global Const $FILE_SHARE_WRITE = 0x2
;~ Global Const $OPEN_EXISTING = 3
;~ Global Const $INVALID_HANDLE_VALUE = Ptr(0xFFFFFFFF) ; invalid ptr value
Global Const $IOCTL_STORAGE_GET_DEVICE_NUMBER = 0x2D1080
Global Const $DRIVE_REMOVABLE = 2
Global Const $DRIVE_FIXED = 3
Global Const $DRIVE_CDROM = 5
Global Const $DIGCF_PRESENT = 0x2
Global Const $DIGCF_DEVICEINTERFACE = 0x10
Global Const $CR_SUCCESS = 0x0
Global Const $DN_REMOVABLE = 0x4000
Global Const $CM_REMOVE_UI_OK = 0x0
Global Const $CM_REMOVE_UI_NOT_OK = 0x1
Global Const $CM_REMOVE_NO_RESTART = 0x2
Global Const $CM_SETUP_DEVNODE_READY = 0x0
Global Const $CM_SETUP_DEVNODE_RESET = 0x4
Global Const $CR_ACCESS_DENIED = 0x33
Global Const $PNP_VetoTypeUnknown = 0 ; Name is unspecified
Global Const $PNP_VetoLegacyDevice = 1 ; Name is an Instance Path
Global Const $PNP_VetoPendingClose = 2 ; Name is an Instance Path
Global Const $PNP_VetoWindowsApp = 3 ; Name is a Module
Global Const $PNP_VetoWindowsService = 4 ; Name is a Service
Global Const $PNP_VetoOutstandingOpen = 5 ; Name is an Instance Path
Global Const $PNP_VetoDevice = 6 ; Name is an Instance Path
Global Const $PNP_VetoDriver = 7 ; Name is a Driver Service Name
Global Const $PNP_VetoIllegalDeviceRequest = 8 ; Name is an Instance Path
Global Const $PNP_VetoInsufficientPower = 9 ; Name is unspecified
Global Const $PNP_VetoNonDisableable = 10 ; Name is an Instance Path
Global Const $PNP_VetoLegacyDriver = 11 ; Name is a Service
Global Const $PNP_VetoInsufficientRights = 12 ; Name is unspecified

; Structures
Global Const $STORAGE_DEVICE_NUMBER = "ulong DeviceType;ulong DeviceNumber;ulong PartitionNumber"
Global Const $SP_DEV_BUF = "byte[2052]"
Global Const $SP_DEVICE_INTERFACE_DETAIL_DATA = "dword cbSize;wchar DevicePath[1024]" ; created at SP_DEV_BUF ptr
Global Const $SP_DEVICE_INTERFACE_DATA = "dword cbSize;byte InterfaceClassGuid[16];dword Flags;ulong_ptr Reserved" ; GUID struct = 16 bytes
Global Const $SP_DEVINFO_DATA = "dword cbSize;byte ClassGuid[16];dword DevInst;ulong_ptr Reserved"

; GUIDs
; GUID_DEVINTERFACE_DISK =      0x53f56307L, 0xb6bf, 0x11d0, 0x94, 0xf2, 0x00, 0xa0, 0xc9, 0x1e, 0xfb, 0x8b
; GUID_DEVINTERFACE_CDROM =     0x53f56308L, 0xb6bf, 0x11d0, 0x94, 0xf2, 0x00, 0xa0, 0xc9, 0x1e, 0xfb, 0x8b
; GUID_DEVINTERFACE_FLOPPY =    0x53f56311L, 0xb6bf, 0x11d0, 0x94, 0xf2, 0x00, 0xa0, 0xc9, 0x1e, 0xfb, 0x8b
Global Const $guidDisk = DllStructCreate($tagGUID)
DllStructSetData($guidDisk, "Data1", 0x53f56307)
DllStructSetData($guidDisk, "Data2", 0xb6bf)
DllStructSetData($guidDisk, "Data3", 0x11d0)
DllStructSetData($guidDisk, "Data4", Binary("0x94f200a0c91efb8b"))
Global Const $guidCDROM = DllStructCreate($tagGUID)
DllStructSetData($guidCDROM, "Data1", 0x53f56308)
DllStructSetData($guidCDROM, "Data2", 0xb6bf)
DllStructSetData($guidCDROM, "Data3", 0x11d0)
DllStructSetData($guidCDROM, "Data4", Binary("0x94f200a0c91efb8b"))
Global Const $guidFloppy = DllStructCreate($tagGUID)
DllStructSetData($guidFloppy, "Data1", 0x53f56311)
DllStructSetData($guidFloppy, "Data2", 0xb6bf)
DllStructSetData($guidFloppy, "Data3", 0x11d0)
DllStructSetData($guidFloppy, "Data4", Binary("0x94f200a0c91efb8b"))

Func _IsUSBHDD(ByRef $query)
    ; if drive type is fixed (3), and device path contains 'usbstor', and parent device ID contains 'USB', and IsRemovable
    Return (($query[1] = 3) And StringInStr($query[3], "usbstor") And StringInStr($query[7], "usb") And $query[8])
EndFunc ;==>_IsUSBHDD

Func _QueryDrive($drive, $fEject = False)
    ; drive letter only!
    $drive = StringLeft($drive, 1)

    Local $queryarray[9]
    ; [0] = device number (int)
    ; [1] = drive type (int)
    ; [2] = dos device name (string)
    ; [3] = device path (string)
    ; [4] = device ID (string)
    ; [5] = device instance (int)
    ; [6] = device instance parent (int)
    ; [7] = parent device ID (string)
    ; [8] = IsRemovable (boolean)

    ; "X:\" -> for GetDriveType
    Local $szRootPath = $drive & ":\"
    ; "X:" -> for QueryDosDevice
    Local $szDevicePath = $drive & ":"
    ; "\\.\X:" -> to open the volume
    Local $szVolumeAccessPath = "\\.\" & $drive & ":"

    Local $DeviceNumber = -1

    Local $hVolume = DllCall("kernel32.dll", "ptr", "CreateFileW", "wstr", $szVolumeAccessPath, "dword", 0, _
            "dword", BitOR($FILE_SHARE_READ, $FILE_SHARE_WRITE), "ptr", 0, "dword", $OPEN_EXISTING, _
            "dword", 0, "ptr", 0)
    $hVolume = $hVolume[0]
    If $hVolume <> $INVALID_HANDLE_VALUE Then
;~ ConsoleWrite("hVolume: " & $hVolume & @CRLF)

        Local $sdn = DllStructCreate($STORAGE_DEVICE_NUMBER)
        Local $res = DllCall("kernel32.dll", "int", "DeviceIoControl", "ptr", $hVolume, "dword", $IOCTL_STORAGE_GET_DEVICE_NUMBER, _
                "ptr", 0, "dword", 0, "ptr", DllStructGetPtr($sdn), "dword", DllStructGetSize($sdn), _
                "dword*", 0, "ptr", 0)
        If $res[0] Then
            $DeviceNumber = DllStructGetData($sdn, "DeviceNumber")
;~ ConsoleWrite("DeviceNumber: " & $DeviceNumber & @CRLF)
            $queryarray[0] = $DeviceNumber

            $res = DllCall("kernel32.dll", "int", "CloseHandle", "ptr", $hVolume)
            If $res[0] And $DeviceNumber <> -1 Then
                Local $DriveType = DllCall("kernel32.dll", "uint", "GetDriveTypeW", "wstr", $szRootPath)
                $DriveType = $DriveType[0]
;~ ConsoleWrite("Drive type: " & $DriveType & @CRLF)
                $queryarray[1] = $DriveType

                ; get the dos device name (like \device\floppy0)
                ; to decide if it's a floppy or not
                $res = DllCall("kernel32.dll", "dword", "QueryDosDeviceW", "wstr", $szDevicePath, "wstr", "", "dword", 260)
                If $res[0] Then
                    Local $szDosDeviceName = $res[2]
;~ ConsoleWrite("DosDeviceName: " & $szDosDeviceName & @CRLF)
                    $queryarray[2] = $szDosDeviceName

                    Local $DevInst = _GetDrivesDevInstByDeviceNumber($DeviceNumber, $DriveType, $szDosDeviceName, $queryarray)
;~ ConsoleWrite("DevInst: " & $DevInst & @CRLF)
                    If $DevInst <> 0 Then
                        $queryarray[4] = $DevInst
                        $res = DllCall("setupapi.dll", "dword", "CM_Get_Device_IDW", "ptr", $DevInst, "wstr", "", "ulong", 1024, "ulong", 0)
;~ ConsoleWrite("DeviceID: " & $res[2] & @CRLF)
                        $queryarray[5] = $res[2]
                        Local $EjectSuccess = _DevInstQuery($DevInst, $queryarray, $fEject)
                    EndIf
                EndIf
            EndIf
        EndIf
    EndIf

    If $fEject Then
        Return $EjectSuccess
    Else
        Return $queryarray
    EndIf
EndFunc ;==>_QueryDrive

Func _RestartDrive($queryarray)
Local $res = DllCall("setupapi.dll", "dword", "CM_Setup_DevNode", "dword", $queryarray[6], "ulong", 0)
Return SetError(Number(Not ($res[0] = 0)), 0, ($res[0] = 0))
EndFunc ;==>_RestartDrive

Func _GetDrivesDevInstByDeviceNumber($DeviceNumber, $DriveType, $szDosDeviceName, ByRef $queryarray)
    Local $IsFloppy = (StringInStr($szDosDeviceName, "\floppy") <> 0)
;~ ConsoleWrite("Is floppy: " & $IsFloppy & @CRLF)

    Local $GUID
    Switch $DriveType
        Case $DRIVE_REMOVABLE
            If $IsFloppy Then
                $GUID = DllStructGetPtr($guidFloppy)
            Else
                $GUID = DllStructGetPtr($guidDisk)
            EndIf
        Case $DRIVE_FIXED
            $GUID = DllStructGetPtr($guidDisk)
        Case $DRIVE_CDROM
            $GUID = DllStructGetPtr($guidCDROM)
        Case Default
            Return 0
    EndSwitch

    ; Get device interface info set handle
    ; for all devices attached to system
    Local $hDevInfo = DllCall("setupapi.dll", "ptr", "SetupDiGetClassDevsW", "ptr", $GUID, "ptr", 0, "hwnd", 0, _
            "dword", BitOR($DIGCF_PRESENT, $DIGCF_DEVICEINTERFACE))
    $hDevInfo = $hDevInfo[0]
    If $hDevInfo <> $INVALID_HANDLE_VALUE Then
;~ ConsoleWrite("hDevInfo: " & $hDevInfo & @CRLF)

        ; Retrieve a context structure for a device interface
        ; of a device information set.
        Local $dwIndex = 0
        Local $bRet

        Local $buf = DllStructCreate($SP_DEV_BUF)
        Local $pspdidd = DllStructCreate($SP_DEVICE_INTERFACE_DETAIL_DATA, DllStructGetPtr($buf))
        Local $cb_spdidd = 6 ; size of fixed part of structure
        If @AutoItX64 Then $cb_spdidd = 8 ; fix for x64
        Local $spdid = DllStructCreate($SP_DEVICE_INTERFACE_DATA)
        Local $spdd = DllStructCreate($SP_DEVINFO_DATA)

        DllStructSetData($spdid, "cbSize", DllStructGetSize($spdid))

        While True
            $bRet = DllCall("setupapi.dll", "int", "SetupDiEnumDeviceInterfaces", "ptr", $hDevInfo, "ptr", 0, _
                    "ptr", $GUID, "dword", $dwIndex, "ptr", DllStructGetPtr($spdid))
            If Not $bRet[0] Then ExitLoop

            Local $res = DllCall("setupapi.dll", "int", "SetupDiGetDeviceInterfaceDetailW", "ptr", $hDevInfo, "ptr", DllStructGetPtr($spdid), _
                    "ptr", 0, "dword", 0, "dword*", 0, "ptr", 0)
            Local $dwSize = $res[5]
;~ ConsoleWrite("dwSize: " & $dwSize & @CRLF)

            If $dwSize <> 0 And $dwSize <= DllStructGetSize($buf) Then
                DllStructSetData($pspdidd, "cbSize", $cb_spdidd)
                _ZeroMemory(DllStructGetPtr($spdd), DllStructGetSize($spdd))
                DllStructSetData($spdd, "cbSize", DllStructGetSize($spdd))

                $res = DllCall("setupapi.dll", "int", "SetupDiGetDeviceInterfaceDetailW", "ptr", $hDevInfo, "ptr", DllStructGetPtr($spdid), _
                        "ptr", DllStructGetPtr($pspdidd), "dword", $dwSize, "dword*", 0, "ptr", DllStructGetPtr($spdd))

                If $res[0] Then
;~ ConsoleWrite("DevicePath: " & DllStructGetData($pspdidd, "DevicePath") & @CRLF)
                    $queryarray[3] = DllStructGetData($pspdidd, "DevicePath")

                    Local $hDrive = DllCall("kernel32.dll", "ptr", "CreateFileW", "wstr", DllStructGetData($pspdidd, "DevicePath"), "dword", 0, _
                            "dword", BitOR($FILE_SHARE_READ, $FILE_SHARE_WRITE), "ptr", 0, "dword", $OPEN_EXISTING, _
                            "dword", 0, "ptr", 0)
                    $hDrive = $hDrive[0]
;~ ConsoleWrite("hDrive: " & $hDrive & @CRLF)

                    If $hDrive <> $INVALID_HANDLE_VALUE Then
                        Local $sdn2 = DllStructCreate($STORAGE_DEVICE_NUMBER)
                        $res = DllCall("kernel32.dll", "int", "DeviceIoControl", "ptr", $hDrive, "dword", $IOCTL_STORAGE_GET_DEVICE_NUMBER, _
                                "ptr", 0, "dword", 0, "ptr", DllStructGetPtr($sdn2), "dword", DllStructGetSize($sdn2), _
                                "dword*", 0, "ptr", 0)

                        If $res[0] Then
                            If $DeviceNumber = DllStructGetData($sdn2, "DeviceNumber") Then
                                $res = DllCall("kernel32.dll", "int", "CloseHandle", "ptr", $hDrive)
;~ If Not $res[0] Then ConsoleWrite("Error closing volume: " & $hDrive & @CRLF)
                                $res = DllCall("setupapi.dll", "int", "SetupDiDestroyDeviceInfoList", "ptr", $hDevInfo)
;~ If Not $res[0] Then ConsoleWrite("Destroy error." & @CRLF)
                                Return DllStructGetData($spdd, "DevInst")
                            EndIf
                        EndIf

                        $res = DllCall("kernel32.dll", "int", "CloseHandle", "ptr", $hDrive)
;~ If Not $res[0] Then ConsoleWrite("Error closing volume: " & $hDrive & @CRLF)
                    EndIf
                EndIf
            EndIf
            $dwIndex += 1
        WEnd

        $res = DllCall("setupapi.dll", "int", "SetupDiDestroyDeviceInfoList", "ptr", $hDevInfo)
;~ If Not $res[0] Then ConsoleWrite("Destroy error." & @CRLF)
    EndIf

    Return 0
EndFunc ;==>_GetDrivesDevInstByDeviceNumber

Func _DevInstQuery($DevInst, ByRef $queryarray, $fEject)
    ; get drives's parent, e.g. the USB bridge,
    ; the SATA port, an IDE channel with two drives, etc.
    Local $res = DllCall("setupapi.dll", "dword", "CM_Get_Parent", "dword*", 0, "dword", $DevInst, "ulong", 0)
    If $res[0] = $CR_SUCCESS Then
        Local $DevInstParent = $res[1]
;~ ConsoleWrite("DevInstParent: " & $DevInstParent & @CRLF)
        $queryarray[6] = $DevInstParent

        $res = DllCall("setupapi.dll", "dword", "CM_Get_Device_IDW", "ptr", $DevInstParent, "wstr", "", "ulong", 1024, "ulong", 0)
;~ ConsoleWrite("Parent DeviceID: " & $res[2] & @CRLF)
        $queryarray[7] = $res[2]

        $res = DllCall("setupapi.dll", "dword", "CM_Get_DevNode_Status", "ulong*", 0, "ulong*", 0, "ptr", $DevInstParent, "ulong", 0)
        Local $IsRemovable = (BitAND($res[1], $DN_REMOVABLE) <> 0)
;~ ConsoleWrite("IsRemovable: " & $IsRemovable & @CRLF)
        $queryarray[8] = $IsRemovable
    EndIf

    If $fEject Then
        Local $bSuccess = False
        For $tries = 1 To 3
            ; sometimes we need some tries...
;~          ConsoleWrite("Try: " & $tries & @CRLF)
            $res = DllCall("setupapi.dll", "dword", "CM_Query_And_Remove_SubTreeW", "dword", $DevInstParent, _
                    "dword*", 0, "wstr", "", "ulong", 260, "ulong", $CM_REMOVE_UI_OK)
            If $res[0] = $CR_ACCESS_DENIED Then
;~              ConsoleWrite("Trying CM_Request_Device_Eject..." & @CRLF)
                $res = DllCall("setupapi.dll", "dword", "CM_Request_Device_EjectW", "dword", $DevInstParent, _
                        "dword*", 0, "wstr", "", "ulong", 260, "ulong", 0)
            EndIf
;~          ConsoleWrite("VetoType: " & $res[2] & @CRLF) ; success when type = 0
;~          ConsoleWrite("VetoName: " & $res[3] & @CRLF) ; name will be blank on success

            $bSuccess = (($res[0] = $CR_SUCCESS) And ($res[2] = $PNP_VetoTypeUnknown))
            If $bSuccess Then ExitLoop

            Sleep(500) ; required to give the next tries a chance
        Next
        Return $bSuccess
    Else
        Return 0
    EndIf
EndFunc ;==>_DevInstQuery

Func _ZeroMemory($ptr, $size)
    DllCall("kernel32.dll", "none", "RtlZeroMemory", "ptr", $ptr, "ulong_ptr", $size)
EndFunc ;==>_ZeroMemory

Func Usage()
	Fail("usage:  eject_usb <drive leter>")
EndFunc

Func Fail($message)
	ConsoleWrite($message & @CRLF)
	Exit(1)
EndFunc

if $CmdLine[0] <> 1 Then
	Usage()
EndIf
$drive = $CmdLine[1]

; Global $drive = InputBox("Eject Drive", "Enter drive (letter only):", "", " M1", 200, 130)
;If @error Then Exit
$drive = StringUpper(StringLeft($drive, 1))
If Not FileExists($drive & ":\") Then
	Fail("Error:  Drive not found")
EndIf

Global $driveInfo[9] = ["Device Number", "Drive Type", "DOS Device Name", "Device Path", "Device ID", "Device Instance", "Device Instance Parent", _
"Parent Device ID", "IsRemovable"]
Global $driveArray = _QueryDrive($drive)
For $i = 0 To UBound($driveArray) - 1
ConsoleWrite("-" & $driveInfo[$i] & ": " & $driveArray[$i] & @CRLF)
Next
ConsoleWrite("-Is USBHDD: " & _IsUSBHDD($driveArray) & @CRLF)
$EjectSuccess = _QueryDrive($drive, True)
ConsoleWrite("Ejecting drive <" & $drive & ":> - " & $EjectSuccess & @CRLF)
;If (6 = MsgBox(4 + 32, "Eject?", "Eject this drive?")) Then ConsoleWrite("Ejecting drive <" & $drive & ":> - " & _QueryDrive($drive, True) & @CRLF)
;If (6 = MsgBox(4 + 32, "Restart?", "Restart this drive?")) Then ConsoleWrite("Restarting drive <" & $drive & ":> - " & _RestartDrive($driveArray) & @CRLF)