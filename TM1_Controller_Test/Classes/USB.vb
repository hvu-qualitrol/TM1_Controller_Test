Imports System
Imports System.IO
Imports System.IO.File
Imports System.Runtime.InteropServices

Module EjectUSB
    <DllImport("kernel32.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function GetVolumeNameForVolumeMountPoint(ByVal lpszVolumeMountPoint As String, ByVal lpszVolumeName As IntPtr, ByVal cchBufferLength As Integer) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function SetupDiGetClassDevs(ByRef ClassGuid As Guid, ByVal Enumerator As IntPtr, ByVal hWndParent As Integer, ByVal Flags As Integer) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function SetupDiDestroyDeviceInfoList(ByVal DeviceInfoSet As Integer) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function SetupDiEnumDeviceInterfaces(ByVal DeviceInfoSet As Integer, ByVal DeviceInfoData As IntPtr, ByRef Guid As Guid, ByVal MemberIndex As Integer, ByVal DeviceInterfaceData As IntPtr) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function SetupDiEnumDeviceInfo(ByVal DeviceInfoSet As Integer, ByVal MemberIndex As Integer, ByVal DeviceInterfaceData As IntPtr) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function SetupDiGetDeviceInstanceId(ByVal DeviceInfoSet As Integer, ByVal DeviceInfoData As IntPtr, ByVal DeviceInstanceID As IntPtr, ByVal DeviceInstanceIdSize As Integer, ByRef RequiredSize As IntPtr) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function SetupDiGetDeviceInterfaceDetail(ByVal DeviceInfoSet As Integer, ByVal DeviceInterfaceData As IntPtr, ByVal DeviceInterfaceDetailData As IntPtr, ByVal DeviceInterfaceDetailDataSize As Integer, ByVal RequiredSize As Integer, ByVal DeviceInfoData As IntPtr) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function CM_Get_Parent(ByRef pdnDevInst As Integer, ByVal dnDevInst As Integer, ByVal ulFlags As Integer) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function CM_Get_Device_ID_Size(ByRef pulLen As Integer, ByVal dnDevInst As Integer, ByVal ulFlags As Integer) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function CM_Get_Device_ID(ByVal dnDevInst As Integer, ByVal BufferPtr As IntPtr, ByVal BufferLen As String, ByVal ulFlags As Integer) As Integer
    End Function
    <DllImport("setupapi.dll", CallingConvention:=CallingConvention.Winapi)> _
    Public Function CM_Request_Device_Eject(ByVal dnDevInst As Integer, ByVal pVetoType As IntPtr, ByVal pszVetoName As IntPtr, ByVal ulNameLength As Integer, ByVal ulFlags As Integer) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)> _
    Public Class SP_DEVINFO_DATA
        Public cbSize As Integer
        Public InterfaceClassGUID As Guid
        Public DevInst As Integer
        Public Reserved As Integer
    End Class
    <StructLayout(LayoutKind.Sequential)> _
    Public Class SP_DEVICE_INTERFACE_DATA
        Public cbSize As Integer
        Public InterfaceClassGUID As Guid
        Public Flags As Integer
        Public Reserved As Integer
    End Class
    Public Const DIGCF_PRESENT As Integer = &H2
    Public Const DIGCF_ALLCLASSES As Integer = &H4
    Public Const DIGCF_DEVICEINTERFACE As Integer = &H10
    Public Const INVALID_HANDLE_VALUE As Integer = -1
End Module

Public Class USB
    Private _ErrorMsg As String

    Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function EjectUsbDrive(ByVal DriveLetter As String) As Boolean
        Dim GUID_DEVINTERFACE_VOLUME As Guid = New Guid(&H53F5630D, &HB6BF, &H11D0, &H94, &HF2, &H0, &HA0, &HC9, &H1E, &HFB, &H8B)
        Dim hDevInfo As Integer = SetupDiGetClassDevs(GUID_DEVINTERFACE_VOLUME, IntPtr.Zero, 0, DIGCF_PRESENT + DIGCF_DEVICEINTERFACE)
        If hDevInfo <> INVALID_HANDLE_VALUE Then
            Dim MemberIndex As Integer = 0
            Dim did As New SP_DEVINFO_DATA
            did.cbSize = Marshal.SizeOf(did)
            Dim didPtr As IntPtr = Marshal.AllocHGlobal(did.cbSize)
            Marshal.StructureToPtr(did, didPtr, True)
            While SetupDiEnumDeviceInfo(hDevInfo, MemberIndex, didPtr) <> 0
                Marshal.PtrToStructure(didPtr, did)
                'Console.WriteLine(did.DevInst & ": " & GetDeviceID(hDevInfo, didPtr) & "\")
                'Console.WriteLine("Mounted volume: " & GetVolumeName(hDevInfo, did))
                Dim pDevInst As Integer
                CM_Get_Parent(pDevInst, did.DevInst, 0)
                'Console.WriteLine(vbTab & pDevInst & ": " & GetDeviceID(pDevInst))
                CM_Get_Parent(pDevInst, pDevInst, 0)
                Dim BusDeviceID As String = GetDeviceID(pDevInst)
                'Console.WriteLine(vbTab & vbTab & pDevInst & ": " & BusDeviceID)
                Form1.AppendText(vbTab & vbTab & pDevInst & ": " & BusDeviceID)
                If BusDeviceID.StartsWith("USB\") Then
                    Form1.AppendText("Requesting USB drive removal...")
                    'Console.WriteLine("Requesting USB drive removal...")
                    If CM_Request_Device_Eject(pDevInst, IntPtr.Zero, IntPtr.Zero, 0, 0) = 0 Then
                        Form1.AppendText("*** It's now safe to remove your USB drive")
                        'Console.WriteLine("*** It's now safe to remove your USB drive")
                    Else
                        _ErrorMsg = "Problem removing USB drive" + vbCr
                        _ErrorMsg += did.DevInst & ": " & GetDeviceID(hDevInfo, didPtr) & "\" + vbCr
                        _ErrorMsg += "Mounted volume: " & GetVolumeName(hDevInfo, did) + vbCr
                        _ErrorMsg += vbTab & pDevInst & ": " & GetDeviceID(pDevInst) + vbCr
                        _ErrorMsg += vbTab & vbTab & pDevInst & ": " & BusDeviceID + vbCr
                        Return False
                    End If
                End If
                MemberIndex += 1
                Console.WriteLine()
            End While
            Marshal.FreeHGlobal(didPtr)
            SetupDiDestroyDeviceInfoList(hDevInfo)
        End If
        Return True
    End Function
    Private Function GetDeviceID(ByVal hDevInfo As Integer, ByVal didPtr As IntPtr) As String
        Dim DeviceID As String = ""
        Dim idPtr As IntPtr = Marshal.AllocHGlobal(1024)
        If SetupDiGetDeviceInstanceId(hDevInfo, didPtr, idPtr, 1024, IntPtr.Zero) <> 0 Then
            Try
                DeviceID = Marshal.PtrToStringAnsi(idPtr, 1024)
                DeviceID = Left(DeviceID, InStr(DeviceID, Chr(0)) - 1)
            Catch ex As Exception
            End Try
        End If
        Marshal.FreeHGlobal(idPtr)
        Return DeviceID
    End Function
    Private Function GetDeviceID(ByVal DevInst As Integer) As String
        Dim ReqLen As Integer
        Dim DeviceID As String = "???"
        If CM_Get_Device_ID_Size(ReqLen, DevInst, 0) = 0 Then
            Dim idPtr As IntPtr = Marshal.AllocHGlobal(ReqLen + 1)
            If CM_Get_Device_ID(DevInst, idPtr, ReqLen + 1, 0) = 0 Then
                Try
                    DeviceID = Marshal.PtrToStringAnsi(idPtr, 1024)
                    DeviceID = Left(DeviceID, InStr(DeviceID, Chr(0)) - 1)
                Catch ex As Exception
                End Try
            End If
            Marshal.FreeHGlobal(idPtr)
        End If
        Return DeviceID
    End Function
    Private Function GetVolumeName(ByVal hDevInfo As Integer, ByVal did As SP_DEVINFO_DATA) As String
        Return "???"
    End Function

    Public Function FindThumbDrive(ByRef DriveLetter) As Boolean
        Dim allDrives() As DriveInfo
        Dim d As DriveInfo
        Dim ThumbDriveCnt As Integer = 0
        Dim startTime As DateTime = Now

        allDrives = DriveInfo.GetDrives
        For Each d In allDrives
            If (d.DriveType = DriveType.Removable) Then
                While Not d.IsReady And Now.Subtract(startTime).TotalSeconds < 20

                End While
                If d.IsReady Then
                    ThumbDriveCnt += 1
                    DriveLetter = d.Name
                End If
            End If
        Next

        If Not ThumbDriveCnt = 1 Then
            _ErrorMsg = "Found " + ThumbDriveCnt.ToString + " thumb drives, expected 1"
            Return False
        End If

        Return True
    End Function

    Private AutoitOutputFile As FileStream
    Private AutoitOutputFileWriter As StreamWriter
    Private AutoitStderrFile As FileStream
    Private AutoitStderrFileWriter As StreamWriter

    Public Function EjectUsbDrive_AutoIt(ByVal DriveLetter As String) As Boolean
        Dim p As New Process
        Dim startTime As DateTime
        Dim AutoitResultFile As FileStream
        Dim AutoitResultFileReader As StreamReader
        Dim Line As String
        Dim Buffer As String
        Dim ExpectedString As String
        Dim USB_UNMOUNTED As Boolean = False

        DriveLetter = DriveLetter.Replace(":", "")
        DriveLetter = DriveLetter.Replace("\", "")

        p.StartInfo.UseShellExecute = False
        p.StartInfo.FileName = UtilitiesDirectory + "eject_usb.exe"
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.RedirectStandardError = True
        p.StartInfo.CreateNoWindow = True
        p.StartInfo.Arguments = DriveLetter

        AutoitOutputFile = New FileStream("\temp\autoit_output.txt", FileMode.Create, FileAccess.Write)
        AutoitOutputFileWriter = New StreamWriter(AutoitOutputFile)
        AutoitStderrFile = New FileStream("\temp\autoit_stderr.txt", FileMode.Create, FileAccess.Write)
        AutoitStderrFileWriter = New StreamWriter(AutoitStderrFile)

        AddHandler p.OutputDataReceived, AddressOf AutoitOutputHandler
        AddHandler p.ErrorDataReceived, AddressOf AutoitStderrHandler
        p.Start()
        p.BeginOutputReadLine()
        p.BeginErrorReadLine()

        startTime = Now
        While (Not p.HasExited And (Now.Subtract(startTime).TotalSeconds) < 30)
            System.Threading.Thread.Sleep(100)
            Application.DoEvents()
        End While
        If (Not p.HasExited) Then
            p.Kill()
            _ErrorMsg = "AUTOIT SCRIPT TIMEOUT"
            Return False
        End If

        System.Threading.Thread.Sleep(500)
        AutoitOutputFileWriter.Close()
        AutoitOutputFile.Close()
        AutoitResultFile = New FileStream("\temp\autoit_output.txt", FileMode.Open, FileAccess.Read)
        AutoitResultFileReader = New StreamReader(AutoitResultFile)

        Line = AutoitResultFileReader.ReadLine()
        Buffer = ""
        ExpectedString = "Ejecting drive <" + DriveLetter + ":> - True"
        While (Not Line Is Nothing)
            Buffer = Buffer + Line + Chr(13)
            Form1.AppendText(Line)
            If Line.Contains(ExpectedString) Then
                USB_UNMOUNTED = True
            End If
            Line = AutoitResultFileReader.ReadLine()
        End While
        AutoitResultFileReader.Close()
        AutoitResultFile.Close()

        AutoitStderrFileWriter.Close()
        AutoitStderrFile.Close()
        AutoitResultFile = New FileStream("\temp\autoit_stderr.txt", FileMode.Open, FileAccess.Read)
        AutoitResultFileReader = New StreamReader(AutoitResultFile)

        Line = AutoitResultFileReader.ReadLine()
        While (Not Line Is Nothing)
            Buffer = Buffer + Line + Chr(13)
            Form1.AppendText(Line)
            Line = AutoitResultFileReader.ReadLine()
        End While
        AutoitResultFileReader.Close()
        AutoitResultFile.Close()

        p.Close()

        If Not USB_UNMOUNTED Then
            _ErrorMsg = "Could not unmount drive"
            Form1.AppendText("Expected:  '" + ExpectedString + "'")
            Return False
        End If

        Return True
    End Function

    Private Sub AutoitOutputHandler(sendingProcess As Object, outLine As DataReceivedEventArgs)
        AutoitOutputFileWriter.Write(Environment.NewLine + outLine.Data)
    End Sub

    Private Sub AutoitStderrHandler(sendingProcess As Object, outLine As DataReceivedEventArgs)
        AutoitStderrFileWriter.Write(Environment.NewLine + outLine.Data)
    End Sub
End Class
