Imports System.IO

Public Enum RelayState
    RELAY_OFF = 0
    RELAY_ON = 1
End Enum

Public Enum AssemblyType
    CONTROLLER_BOARD = 0
    H2SCAN = 1
    TM1 = 2
    DISPLAY_KIT
End Enum

'Public Enum State
'    _OFF = 0
'    _ON = 1
'End Enum

'Public Enum LED
'    ALARM = 0
'    POWER = 1
'    SERVICE = 2
'End Enum

Public Enum DaqAnalogConfiguration
    SE = 0
    DIFF = 1
End Enum

Module Module1
    Public RetriesAllowed As Boolean = True
    Public ReportDir As String
    ' Public ReportDir As String = "\Operations\Production\TM1\Controller\Records\"
    ' Public ReportDir As String = "\Operations\Production\TM1\"
    'Public ReportDir As String = "M:\Operations\Production\Operations Data\TM1\"
    Public BaseFWDirectory As String = "C:\Operations\Production\TM1\Controller\firmware\"
    Public UtilitiesDirectory As String = "C:\Operations\Production\TM1\Controller\utilities\"
    Public Products As New Dictionary(Of String, Hashtable)
    Public Product As String = "STANDARD"
    Public SerialPort = Nothing
    Public Daq As MccDaq.MccBoard
    Public Stopped As Boolean = True
    Public LoggingData As Boolean = False
    Public Tmcom1SerialPort = Nothing
    Public TM1_SN As String
    Public ConnectString As String

    Public Enum State
        _OFF = 0
        _ON = 1
    End Enum
    Public State_Strings() As String = {"OFF", "ON"}

    Public Enum LED
        ALARM = 0
        POWER = 1
        SERVICE = 2
    End Enum
    Public LED_Strings() As String = {"ALARM", "POWER", "SERVICE"}

    Public AssemblyLevel As String = "UNDEFINED"

    Public Specs As Hashtable

    Public TestDataTable As DataTable
    Public TestDataString As String

    Public HeaterFields() As String = {"Timestamp|DT", "H2_OIL|I", "H2_DGA.PPM|I", "H2.PPM|I", "SnsrTemp|D", "PcbTemp|D", "OilTemp|D", "Status|S",
"HCurrent|I", "ResAdc|I", "AdjRes|I", "H2Res.PPM|I", "OilBlockTemp|D", "OilBlockPWM|D", "OilBlockDir|S", "AnalogBdTemp|D", "AnalogBdPWM|D",
"AnalogBdDir|S", "AmbientTemp|D", "OilblockRevPWM|D", "AnalogBdRevPWM|D", "Tach|I"}

    Public DaqLockFile As FileStream
    Public DaqLockFilename As String = "c:\temp\MCCDAQ_LOCK.lck"
    Public USB1208LS_Devices As Hashtable

    Public FtdiLockFile As FileStream
    Public FtdiLockFilename As String = "c:\temp\FTDI_LOCK.lck"
    Public FtdiMM_File As Object
    Public FtdiMM_FileName As String = "FTDI_MemoryFile"
    Public FtdiHash As New Hashtable

    Public RemoteDaqConnected As Boolean = False
    Public DaqServerProcess As Process
    Public DaqServerPath As String = "C:\Serveron\Visual Basic\Code\RemoteDaqServer\RemoteDaqServer.exe"

    Public AppHexColorStr() As String = {"#07F6FA", "#EBA2FA", "#AD7272"}
    'Public a As Color = Color.FromArgb("0xe6e6e6")
    'Public b As Color = Color.FromArgb(



End Module