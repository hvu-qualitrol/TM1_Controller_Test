Imports System.IO
Imports System.IO.File
Imports System.Text.RegularExpressions
Imports TM1_Controller_Test.Tests

Public Class Form1
    Private ReportFile As FileStream
    Private ReportFileWriter As StreamWriter
    Private ReportFileOpen As Boolean = False
    Private Test_Sequence As New ArrayList
    Public Serial_Number As String
    Public LogToDatabase As Boolean = False
    Public TestDetails As String
    Public TestButton_toolTip(30) As ToolTip

    Sub New()
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'For Each Str As String In System.IO.Ports.SerialPort.GetPortNames()
        '    Tmcom1ComboBox.Items.Add(Str)
        '    BK1856D_com_ComboBox.Items.Add(Str)
        'Next

        ReportDir = configurationAppSettings.GetValue("DrivePath", GetType(String))

        If Not CreateLocalDirs() Then
            MessageBox.Show("Problem creating local directory structure")
            Exit Sub
        End If

        If Not ExtractFirmware() Then
            MessageBox.Show("Problem exracting firwmare to local directory")
            Exit Sub
        End If

        If Not ExtractUtilities() Then
            MessageBox.Show("Problem extracting utilities to local directory")
            Exit Sub
        End If

        ' Me.Serial_Number_Entry.Text = "TM1010113030021"
        ' Me.Controller_Board_SN_Entry.Text = "2Z03U094"
        ' Me.PowerSupply_SN_Entry.Text = "xxxxxxxxxxxxxx"
        ' Me.Pump_SN_Entry.Text = "xxxxxxxxx"

    End Sub

    Private Sub Form1_Load( _
        ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles MyBase.Load

        Dim SerialConfigFile As FileStream
        Dim SerialConfigFileReader As StreamReader
        Dim Line As String
        Dim Name As String
        Dim SerialPort As String
        Dim Fields() As String

        ListBox_TestMode.SelectedIndex = 0
        'InitializeTestSequence()
        LoadProductInfo()

        If File.Exists("\Temp\TM1_Controller_Test_SerialConfig.txt") Then
            SerialConfigFile = New FileStream("\Temp\TM1_Controller_Test_SerialConfig.txt", FileMode.Open, FileAccess.Read)
            SerialConfigFileReader = New StreamReader(SerialConfigFile)
            Line = SerialConfigFileReader.ReadLine()
            While (Not Line Is Nothing)
                If Line.Contains(":") Then
                    Fields = Regex.Split(Line, ":")
                    If Not Fields.Length = 2 Then
                        Continue While
                    End If
                    Name = Fields(0)
                    SerialPort = Fields(1)
                    AppendText(Name + " " + SerialPort)
                    If Name = "TMCOM1" Then
                        If ComboBoxContains(Tmcom1ComboBox, SerialPort) Then
                            Tmcom1ComboBox.Text = SerialPort
                        End If
                    End If
                    If Name = "BK1856D" Then
                        If ComboBoxContains(BK1856D_com_ComboBox, SerialPort) Then
                            BK1856D_com_ComboBox.Text = SerialPort
                        End If
                    End If
                End If
                Line = SerialConfigFileReader.ReadLine()
            End While
        End If
        For i = 0 To 29
            TestButton_toolTip(i) = New ToolTip
            TestButton_toolTip(i).ShowAlways = True
        Next

    End Sub

    Sub SetSpecs()
        Dim S As Spec
        Specs = New Hashtable

        S = New Spec(4.9, 5.1)
        Specs.Add("5VREF", S)

        S = New Spec(23.3, 24.9)
        Specs.Add("24VREF", S)

        S = New Spec(19.8, 20.2)
        Specs.Add("AUX1_20mA", S)

        S = New Spec(9.8, 10.2)
        Specs.Add("AUX1_10mA", S)

        S = New Spec(5.0, 5.4)
        Specs.Add("AUX1_5mA", S)

        S = New Spec(0, 0.5)
        Specs.Add("AUX1_0mA", S)

        S = New Spec(19.8, 20.2)
        Specs.Add("AUX2_20mA", S)

        S = New Spec(9.8, 10.2)
        Specs.Add("AUX2_10mA", S)

        S = New Spec(5.0, 5.4)
        Specs.Add("AUX2_5mA", S)

        S = New Spec(0, 0.5)
        Specs.Add("AUX2_0mA", S)

        S = New Spec(4.75, 5.25)
        Specs.Add("420_20mA_250ohm", S)

        S = New Spec(2.25, 2.75)
        Specs.Add("420_10mA_250ohm", S)

        S = New Spec(1.1, 1.4)
        Specs.Add("420_5mA_250ohm", S)

        S = New Spec(-0.2, 0.2)
        Specs.Add("420_0mA_250ohm", S)

        S = New Spec(-0.0003, 0.0002)
        Specs.Add("Leak_Rate", S)

        S = New Spec(0.3, 0.7)
        Specs.Add("PWM_HEAT_analogbd_pwm", S)

        S = New Spec(0.08, 0.5)
        Specs.Add("PWM_HEAT_oilblock_pwm", S)

        S = New Spec(45.9, 46.1)
        Specs.Add("PWM_HEAT_oilblock_temp", S)

        S = New Spec(54.5, 55.5)
        Specs.Add("PWM_HEAT_analogbd_temp", S)

        S = New Spec(0, 0.4)
        Specs.Add("PWM_HEAT_oilblock_stdev", S)

        S = New Spec(0, 0.7)
        Specs.Add("PWM_HEAT_analogbd_stdev", S)

        S = New Spec(0.3, 0.65)
        Specs.Add("TEC_COOL_oilblock_pwm", S)

        S = New Spec(0.35, 0.71)
        Specs.Add("TEC_COOL_analogbd_pwm", S)

        S = New Spec(20, 22)
        Specs.Add("TEC_COOL_oilblock_temp", S)

        S = New Spec(20, 24.0)
        Specs.Add("TEC_COOL_analogbd_temp", S)

        S = New Spec(0, 0.5)
        Specs.Add("TEC_COOL_oilblock_stdev", S)

        S = New Spec(0, 0.6)
        Specs.Add("TEC_COOL_analogbd_stdev", S)

        S = New Spec(50, 80)
        Specs.Add("TEC_HEAT_oilblock_time", S)

        S = New Spec(50, 90)
        Specs.Add("TEC_HEAT_analogbd_time", S)

        S = New Spec(80, 240)
        Specs.Add("TEC_COOL_oilblock_time", S)

        S = New Spec(80, 300)
        Specs.Add("TEC_COOL_analogbd_time", S)
    End Sub

    Sub InitializeTestSequence()
        Dim TS As Test_Item
        Dim EntriesNeeded As Hashtable
        Dim testSequence As Integer = 0
        Dim buttons = {Test_1, Test_2, Test_3, Test_4, Test_5, Test_6, Test_7, Test_8, Test_9, Test_10, Test_11, Test_12, Test_13, Test_14, Test_15, Test_16, Test_17, Test_18, Test_19, Test_20}

        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Visible = False
        Next
        Test_Sequence.Clear()

        If AssemblyLevel = "CONTROLLER_BOARD" Then
            testSequence = 0
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "FLASH"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_FLASH)
            EntriesNeeded = New Hashtable
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_1, "Programs and verifies Freescale flash using the J-Link")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "USB_SERIAL"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_USB_SERIAL)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_2, "Programs the controller board SN into FTDI USB to RS232 device and finds the mapped COM port")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "BOOT"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_BOOT)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_3, "Verifies the controller board is booted by logging into the usb serial port." + vbCr +
                                                     "300s is allowed for the boot to accomodate the time required for the records.dat file" + vbCr +
                                                     "to be built on the SD card.  While the records.dat file is being built the power," + vbCr +
                                                     "alarm and service leds are flashed")
            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "VERSION"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_VERSION)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_4, "Verifies the flash version.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "RTC"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_RTC)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("BK1856D COM Port", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_5, "Adjusts the RTC crystal capactiance to center the RTC frequency." + vbCr +
                                                     "The RTC outputs a 1Hz signal from on the processor which is measured " + vbCr +
                                                     "by the frequency counter.  The board temp sensor is expected to be between 23C & 30C")
            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "SET_DATE"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_SETDATE)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_6, "Sets the controller board date to the workstation NTP controlled date.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "LED"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_LED)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_7, "Individually flashes the alarm, power and service LED's.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "RESET"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_RESET)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_8, "Test that the Reset button will reset the unit.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "RELAY"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_RELAY)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("USB1208LS", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_9, "Relay continuity is verified with the USB daq in both NC and NO positions")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "PUMP"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_PUMP)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_10, "Pump RPM and direction verified at 2 forward and reverse speeds")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "H2SCAN"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_H2SCAN)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_11, "Verifies the ability to communicate with the test fixture H2SCAN")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "OILBLOCK_HEATER"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_OILBLOCK_HEATER)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_12, "Verifies the ability to control the test fixture oilblock & analog board heater's and TEC's" + vbCr +
                                                       "The TEC's are replaced with 1.1 Ohm resistors, the voltage drop is measured by the USB-DAQ" + vbCr +
                                                       "The TEC control is tested in both heating and cooling mode, by verifying the polarity of the voltage drop." + vbCr +
                                                       "The PWM heating controls/sensing is verified by detecting a 3C temp rise")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "TMCOM1"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_TMCOM1)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TMCOM1 COM Port", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_13, "Tests functionality of TMCOM1 serial port")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "VER_DATE"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_VERDATE)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_14, "Verifies the date/time")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "SN_CONFIG"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_SNCONFIG)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_15, "Sets the config SERIAL_NUMBER$ field to a pseudo TM1 SN." + vbCr +
                                                       "This is needed to verify the USB copy")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "USB-A"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_USBA)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_16, "Verifies the functionality of the USB-A port." + vbCr +
                                                       "SD card files are synced to the thumb drive and then verified on the PC")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(16).Button = buttons(testSequence)
            Test_Sequence(16).Button.Text = "POWER_CHECK"
            Test_Sequence(16).Handler = New MyDelFun(AddressOf Test_PWRCHECK)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(16).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(16).SetToolTip(Test_17, "Verifies 24VREF and 5VREF are in spec")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "ADC"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_420)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("USB1208LS", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_18, "Test of the 4-20 IO" + vbCr +
                                                        "Outputs 1 & 2 are looped back to Input's 1 & 2." + vbCr +
                                                        "The dac command is used to set drive the outputs and the adc command to read the inputs" + vbCr +
                                                        "Output 3 is converted to a voltage with a 250 ohm resistor and read with the USB DAQ")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "VERIFY_CONFIG"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_VERIFY_CONFIG)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_19, "Verifies the config values")
        End If

        If AssemblyLevel = "SUB_ASSEMBLY" Then
            testSequence = 0
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "USB_SERIAL"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_USB_SERIAL)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("Controller Board Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_1, "Programs the TM1 SN into FTDI USB to RS232 device and finds the mapped COM port")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "BOOT"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_BOOT)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("Controller Board Serial Number", True)
            EntriesNeeded.Add("Power Supply Serial Number", True)
            EntriesNeeded.Add("Pump Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_2, "Verifies the controller board is booted by logging into the usb serial port.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "SN_CONFIG"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_SNCONFIG)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_3, "Sets the config SERIAL_NUMBER$ field to the TM1 SN.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "LED"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_LED)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_4, "Individually flashes the alarm, power and service LED's.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "RESET"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_RESET)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_5, "Test that the Reset button will reset the unit.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "H2SCAN"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_H2SCAN_SUB_ASSEMBLY)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_6, "Verifies the ability to communicate with the H2SCAN" + vbCr +
                                                     "Verifies the H2SCAN serial mode is set to MODBUS mode, if not the MODBUS ID is" + vbCr +
                                                     "is verified to be set to 1 an the serial mode is set to MODBUS." + vbCr +
                                                     "The H2SCAN model number is verified." + vbCr +
                                                     "The H2SCAN SN is linked in the Mfg Database to the TM1")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "H2SCAN_FWU"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_H2SCAN_UPDATE)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_7, "H2SCAN FW updated via the thumb drive.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "VERSION"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_VERSION)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_8, "Verifies controller board and H2SCAN FW versions.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "PWM_HEAT"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_THERMALS)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TMCOM1 COM Port", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_9, "The analog board and oilblock is manually driven to 49C & 40C." + vbCr +
                                                     "The analog board and oilblock PWM set points are set to 55C & 46C and PWM controls are set to auto" + vbCr +
                                                     "The TMCOM1 serial port is used to continuosuly monitor the heater output." + vbCr +
                                                     "The analog board and oilblock are expected to settle close to the set points" + vbCr +
                                                     "within a specified period of time and the PWM values are expected to be within limits")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "H2SCAN_READY"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_H2SCAN_READY)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TMCOM1 COM Port", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_10, "Verifies H2SCAN status reaches 8001." + vbCr +
                                                      "Verifies H2SCAN analog board temp matches analog board heater temp")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "TEC_COOL"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_THERMALS_LOW)
            'Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_TEC_COOL)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TMCOM1 COM Port", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_11, "Manually drives analogbd & oilblock zones to 25C." + vbCr +
                                                       "Set to auto with thermal zones setpoint at 20C" + vbCr +
                                                       "Turns H2SCAN off to prevent heating from H2SCAN" + vbCr +
                                                       "Expects thermal zones to stablize within expected time." + vbCr +
                                                       "Expects final zone temps & pwms at expected values")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "TEC_HEAT"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_TEC_HEAT)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TMCOM1 COM Port", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_12, "Manually heats with TEC's only and expects a 5C temp rise within expected times")


            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "4-20"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_420)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("USB1208LS", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_13, "Test of the 4-20 IO" + vbCr +
                                            "Outputs 1 & 2 are looped back to Input's 1 & 2." + vbCr +
                                            "The dac command is used to set drive the outputs and the adc command to read the inputs" + vbCr +
                                            "Output 3 is converted to a voltage with a 250 ohm resistor and read with the USB DAQ")


            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "RELAY"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_RELAY)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("USB1208LS", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_14, "Relay continuity is verified with the USB daq in both NC and NO positions")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "POWER_CHECK"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_PWRCHECK)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_15, "Verifies 24VREF and 5VREF are in spec")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "PUMP"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_PUMP)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_16, "Pump RPM and direction verified at 2 forward and reverse speeds")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(16).Button = buttons(testSequence)
            Test_Sequence(16).Button.Text = "SET_VER_DATE"
            Test_Sequence(16).Handler = New MyDelFun(AddressOf Test_SETDATE)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(16).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(16).SetToolTip(Test_17, "Sets and verifies the date/time")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "VERIFY_CONFIG"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_VERIFY_CONFIG)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_18, "Verifies the config values")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "LAB_MODE"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_H2SCAN_LAB_MODE)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_19, "Sets the H2SCAN serial mode to LAB MODE")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "EMB_MOISTURE"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf TestEmbeddedMoisture)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_20, "Verifies Embedded Moisture")
        End If

        If AssemblyLevel = "W223_REWORK" Then
            testSequence = 0
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "BOOT"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_BOOT)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("Controller Board Serial Number", True)
            EntriesNeeded.Add("Power Supply Serial Number", True)
            EntriesNeeded.Add("Pump Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_1, "Verifies the controller board is booted by logging into the usb serial port.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "LED"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_LED)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_2, "Individually flashes the alarm, power and service LED's.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "H2SCAN"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_H2SCAN_SUB_ASSEMBLY)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_3, "Verifies the ability to communicate with the H2SCAN" + vbCr +
                                                     "Verifies the H2SCAN serial mode is set to MODBUS mode, if not the MODBUS ID is" + vbCr +
                                                     "is verified to be set to 1 an the serial mode is set to MODBUS." + vbCr +
                                                     "The H2SCAN model number is verified." + vbCr +
                                                     "The H2SCAN SN is linked in the Mfg Database to the TM1")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "VERSION"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_VERSION)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_4, "Verifies controller board and H2SCAN FW versions.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "SET_VER_DATE"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_SETDATE)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_5, "Set and Verifies the date/time")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "THERMALS"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf TestReworkThermals)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TMCOM1 COM Port", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_6, "The analog board and oilblock is manually driven to 49C & 40C." + vbCr +
                                                     "The analog board and oilblock PWM set points are set to 55C & 46C and PWM controls are set to auto" + vbCr +
                                                     "The TMCOM1 serial port is used to continuosuly monitor the heater output." + vbCr +
                                                     "The analog board and oilblock are expected to settle close to the set points" + vbCr +
                                                     "within a specified period of time and the PWM values are expected to be within limits")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "H2SCAN_READY"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf TestReworkH2ScanReady)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TMCOM1 COM Port", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_7, "Verifies H2SCAN status reaches 8001." + vbCr +
                                                      "Verifies H2SCAN analog board temp matches analog board heater temp")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "4-20"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_420)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("USB1208LS", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_8, "Test of the 4-20 IO" + vbCr +
                                            "Outputs 1 & 2 are looped back to Input's 1 & 2." + vbCr +
                                            "The dac command is used to set drive the outputs and the adc command to read the inputs" + vbCr +
                                            "Output 3 is converted to a voltage with a 250 ohm resistor and read with the USB DAQ")


            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "RELAY"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_RELAY)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("USB1208LS", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_9, "Relay continuity is verified with the USB daq in both NC and NO positions")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "POWER_CHECK"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_PWRCHECK)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_10, "Verifies 24VREF and 5VREF are in spec")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "PUMP"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_PUMP)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_11, "Pump RPM and direction verified at 2 forward and reverse speeds")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "VERIFY_CONFIG"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf TestReworkConfig)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_12, "Verifies the config values")
        End If

        If AssemblyLevel = "LEAK" Then
            testSequence = 0
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "LEAK"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_LEAK)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TMCOM1 COM Port", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_1, "Leak rate test")
        End If

        If AssemblyLevel = "DISPLAY_KIT" Then
            testSequence = 0
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "DISPLAY_KIT"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf TestDisplay)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TM1 COM Port", True)
            EntriesNeeded.Add("Controller Board Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_1, "Verify functionalities of retrofit Display kits")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "USB-A"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_USBA)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("TM1 COM Port", True)
            EntriesNeeded.Add("Controller Board Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_2, "Verifies the functionality of the USB-A port." + vbCr +
                                                       "SD card files are synced to the thumb drive and then verified on the PC")
        End If
        If AssemblyLevel = "TM101_TM102" Then
            testSequence = 0
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "FTDI_SN"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf ConvertFtdiSerialNumber)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_1, "Programs the TM102 SN to FTDI memory.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "BOOT"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf Test_BOOT)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            EntriesNeeded.Add("Controller Board Serial Number", True)
            'EntriesNeeded.Add("Power Supply Serial Number", True)
            'EntriesNeeded.Add("Pump Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_2, "Verifies the unit is booted by logging into the usb serial port.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "CONFIG_SN"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf ConvertConfigSerialNumber)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_3, "Sets the config SERIAL_NUMBER$ field to the TM102 SN.")

            testSequence += 1
            TS = New Test_Item
            Test_Sequence.Add(TS)
            Test_Sequence(testSequence).Button = buttons(testSequence)
            Test_Sequence(testSequence).Button.Text = "INFO_CONVERT"
            Test_Sequence(testSequence).Handler = New MyDelFun(AddressOf ConvertTm101ToTm102)
            EntriesNeeded = New Hashtable
            EntriesNeeded.Add("Serial Number", True)
            Test_Sequence(testSequence).EntriesNeeded() = EntriesNeeded
            TestButton_toolTip(testSequence).SetToolTip(Test_4, "Performs info conversion from TM101 to TM102")

        End If

        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Visible = True
        Next
    End Sub

    Private Sub StartTest_Click(sender As System.Object, e As System.EventArgs) Handles StartTest.Click, Test_1.Click, Test_2.Click,
        Test_3.Click, Test_4.Click, Test_5.Click, Test_6.Click, Test_7.Click, Test_8.Click, Test_9.Click,
        Test_10.Click, Test_11.Click, Test_12.Click, Test_13.Click, Test_14.Click, Test_15.Click, Test_16.Click,
        Test_17.Click, Test_18.Click, Test_19.Click, Test_20.Click
        Dim RunSingleTest As Boolean = False
        Dim DebugMode As Boolean = False
        Dim Test_Index As Integer
        Dim TimeStamp As String
        Dim TestStartTime As DateTime = Now
        Dim i As Integer
        Dim Test_Status As Boolean
        Dim Retry As Boolean
        Dim RetryCnt As Integer
        Dim ReportFilePath As String
        Dim TestTime As Integer = 0
        Dim TestName As String
        Dim DB As New DB

        Stopped = False
        If ListBox_TestMode.Text = "DEBUG" Then
            DebugMode = True
        End If

        If Not sender.Equals(StartTest) Then
            RunSingleTest = True
            If Not DebugMode Then
                Stopped = True
                Exit Sub
            End If
        End If

        Results.Clear()
        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black

        If Not CheckEntries() Then
            Fail()
            Exit Sub
        End If
        If Not ListBox_TestMode.Text = "DEBUG" Then
            Serial_Number_Entry.BackColor = Color.Empty
        End If

        If AssemblyLevel_ComboBox.Text = "TM101_TM102" Then
            Serial_Number = Controller_Board_SN_Entry.Text
        End If

        ReportFilePath = ReportDir + AssemblyLevel_ComboBox.Text + "\" + Serial_Number
        Try
            Directory.CreateDirectory(ReportFilePath)
        Catch ex As Exception
            AppendText("Problem creating directory " + ReportFilePath)
            Fail()
        End Try
        'DisableEntries()

        'If Not DebugMode Then
        TimeStamp = Format(Date.UtcNow, "yyyyMMddHHmmss")
        Try
            'ReportFile = New FileStream(ReportDir + Serial_Number + "." + TimeStamp + ".txt", FileMode.Create, FileAccess.Write)
            ReportFile = New FileStream(ReportFilePath + "\" + TimeStamp + ".txt", FileMode.Create, FileAccess.Write)
            ReportFileWriter = New StreamWriter(ReportFile)
        Catch ex As Exception
            AppendText("Problem opening log File")
            Fail()
            Exit Sub
        End Try
        ReportFileOpen = True
        LoggingData = True
        'End If

        If Not DebugMode Then
            If Not DB.CheckPrevTests(Serial_Number, AssemblyLevel) Then
                AppendText("Previous tests have not passed")
                Fail()
                Exit Sub
            End If
        End If

        If RunSingleTest Then
            For i = 0 To Test_Sequence.Count - 1
                If sender.Equals(Test_Sequence(i).Button) Then
                    Test_Index = i
                End If
            Next
            If Test_Index < 0 Then
                AppendText("Could not find test for " + sender.Text)
                For i = 0 To Test_Sequence.Count - 1
                    Test_Sequence(i).Button.Enabled = True
                Next
                Fail()
                Exit Sub
            End If
        End If

        For i = 0 To Test_Sequence.Count - 1
            If RunSingleTest And Not i = Test_Index Then
                Continue For
            End If
            If Not Test_Sequence(i).Button.Enabled Then
                Continue For
            End If
            If Not CheckEntriesForTest(Test_Sequence(i)) Then
                Fail()
                Exit Sub
            End If
        Next
        DisableEntries()

        For i = 0 To Test_Sequence.Count - 1
            If ListBox_TestMode.Text = "DEBUG" Then
                If Test_Sequence(i).Button.Enabled Then
                    Test_Sequence(i).Button.BackColor = Color.LightGray
                    Test_Sequence(i).Button.ForeColor = Color.Black
                End If
            Else
                Test_Sequence(i).Button.Enabled = True
                Test_Sequence(i).Button.ForeColor = DefaultForeColor
                Test_Sequence(i).Button.BackColor = DefaultBackColor
            End If
        Next

        TestStatus.Text = "RUNNING"
        TestStatus.BackColor = Color.Yellow
        TestStatus.ForeColor = Color.Black

        If Stopped Then
            Fail()
            Exit Sub
        End If
        If Not DebugMode Then
            LogToDatabase = True
        End If

        '' Temporarily don't log to the data base if it is display kit test
        'If AssemblyLevel = "DISPLAY_KIT" Then
        '    LogToDatabase = False
        'End If

        'SerialPort = Nothing
        For i = 0 To Test_Sequence.Count - 1
            If RunSingleTest And Not i = Test_Index Then
                Continue For
            End If
            If Not Test_Sequence(i).Button.Enabled = True Then
                Continue For
            End If

            Test_Sequence(i).Button.BackColor = Color.Yellow
            Test_Sequence(i).Button.ForeColor = Color.Black
            RetryCnt = 0
            Retry = True

            ' The PWM_HEAT test is not allowed to retry (BugID # 5497). 
            ' All the other thermal tests should comply to that as well.
            TestName = Test_Sequence(i).Button.Text
            If TestName = "PWM_HEAT" Or TestName = "TEC_COOL" Or TestName = "TEC_HEAT" Then
                RetriesAllowed = False
            Else
                RetriesAllowed = True
            End If

            While Retry
                AppendText("##########################################################")
                If (RetryCnt = 0) Then
                    AppendText("TEST_START:  " + Test_Sequence(i).Button.Text)
                Else
                    AppendText("TEST_RETRY:  " + Test_Sequence(i).Button.Text)
                End If
                TestDetails = ""
                RetryCnt += 1
                Test_Status = Test_Sequence(i).Handler.Invoke()
                TestTime = Now.Subtract(TestStartTime).TotalSeconds
                TimeoutLabel.Text = "NA"
                Retry = False
                If Not Test_Status Then
                    AppendText("TEST_STATUS:  FAILED")
                    If RetriesAllowed And Not Stopped Then
                        If YesNo("Test failed, click yes to retry", "RETRY?") Then
                            Retry = True
                        End If
                        'If (MessageBox.Show("Test failed, click yes to retry", Serial_Number_Entry.Text + ":  RETRY?", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes) Then
                        '    Retry = True
                        'End If
                    End If
                    If Not Retry Then
                        If Stopped Then
                            Test_Sequence(i).Button.BackColor = Color.Blue
                            Test_Sequence(i).Button.ForeColor = Color.White
                        Else
                            Test_Sequence(i).Button.BackColor = Color.Red
                            Test_Sequence(i).Button.ForeColor = Color.White
                        End If
                        Fail(TestName, TestTime)
                        Exit Sub
                    End If
                    'ReportFileWriter.Close()
                    'ReportFile.Close()
                End If
            End While
            AppendText("TEST_STATUS:  PASSED")
            'Test_Sequence(i).Button.ForeColor = Color.White
            Test_Sequence(i).Button.ForeColor = Color.Black
            Test_Sequence(i).Button.BackColor = Color.LightGreen
        Next

        'ReportFileWriter.Close()
        'ReportFile.Close()
        'ReportFileOpen = False
        AppendText("FINAL_STATUS:  PASSED")
        Pass(TestTime)
        If Not DebugMode Then
            SaveSerialPortConfig()
        End If
        LogToDatabase = False
    End Sub

    Sub EnableEntries()
        Dim AssemblyLevel As String = AssemblyLevel_ComboBox.Text

        Serial_Number_Entry.Enabled = True
        Tmcom1ComboBox.Enabled = True
        StartTest.Enabled = True
        EnableAllTest.Enabled = True
        DisableAllTests.Enabled = True
        USB1208LS_ComboBox.Enabled = True

        'Select Case AssemblyLevel
        '    Case "CONTROLLER_BOARD"
        '        Tmcom1ComboBox.Enabled = True
        '        BK1856D_com_ComboBox.Enabled = True
        'End Select
    End Sub

    Sub DisableEntries()
        Serial_Number_Entry.Enabled = False
        Tmcom1ComboBox.Enabled = False
        StartTest.Enabled = False
        StopButton.Enabled = True
        EnableAllTest.Enabled = False
        DisableAllTests.Enabled = False
        USB1208LS_ComboBox.Enabled = False
    End Sub

    Sub Fail(Optional ByVal TestName As String = "", Optional ByVal TestTime As Integer = 0)
        Dim SF As New SerialFunctions
        Dim database As New DB
        Dim DB_success As Boolean = False

        If LogToDatabase Then
            DB_success = database.Fail(Serial_Number, AssemblyLevel, TestTime, TestName, TestDetails)
        End If

        If Stopped Then
            TestStatus.Text = "STOPPED"
            TestStatus.BackColor = Color.Blue
            TestStatus.ForeColor = Color.White
        Else
            TestStatus.Text = "FAILED"
            TestStatus.BackColor = Color.Red
            TestStatus.ForeColor = Color.Black
        End If
        If EnableAllTest.Visible Then
            EnableAllTest.Enabled = True
            DisableAllTests.Enabled = True
        End If
        If LoggingData Then
            ReportFileWriter.Close()
            ReportFile.Close()
            ReportFileOpen = False
            LoggingData = False
        End If
        EnableEntries()
        Stopped = True

        SF.Close()
    End Sub

    Sub Pass(ByVal TestTime As Integer)
        Dim SF As New SerialFunctions
        Dim database As New DB
        Dim DB_success As Boolean = False

        If LogToDatabase Then
            DB_success = database.Pass(Serial_Number, AssemblyLevel, TestTime)
        End If
        If Stopped Then
            TestStatus.BackColor = Color.Red
            TestStatus.Text = "FAILED"
        ElseIf ListBox_TestMode.Text = "DEBUG" Then
            TestStatus.Text = "DEBUG PASSED"
            TestStatus.BackColor = Color.Goldenrod
            TestStatus.ForeColor = Color.Black
        Else
            TestStatus.Text = "PASSED"
            TestStatus.BackColor = Color.Green
            TestStatus.ForeColor = Color.White
        End If
        If EnableAllTest.Visible Then
            EnableAllTest.Enabled = True
            DisableAllTests.Enabled = True
        End If
        'TestStatus.Text = "PASSED"
        'TestStatus.BackColor = Color.Green
        'TestStatus.ForeColor = Color.White
        If LoggingData Then
            ReportFileWriter.Close()
            ReportFile.Close()
            ReportFileOpen = False
            LoggingData = False
        End If
        EnableEntries()
        Stopped = True
        ' SaveSerialPortConfig()

        SF.Close()
    End Sub

    Sub DebugPass()
        TestStatus.Text = "DEBUG PASSED"
        TestStatus.BackColor = Color.Goldenrod
        TestStatus.ForeColor = Color.Black
        EnableEntries()
    End Sub

    Private Sub AddTextLine(ByVal line As String)
        If Results.InvokeRequired Then
            Results.BeginInvoke(Sub() AddTextLine(line))
        Else
            Results.AppendText(line)
        End If
    End Sub

    Public Sub AppendText(ByVal Line As String, Optional ByVal LogToFile As Boolean = True, Optional ByVal CR As Boolean = True)
        TestDetails += Line
        If Not ReportFileOpen Then
            LogToFile = False
        End If
        If LogToFile Then
            ReportFileWriter.Write(Line)
        End If
        'Results.AppendText(Line)
        AddTextLine(Line)
        If CR Then
            If Not Line.Contains(System.Environment.NewLine) Then
                If LogToFile Then
                    ReportFileWriter.Write(System.Environment.NewLine)
                End If
                'Results.AppendText(System.Environment.NewLine)
                AddTextLine(System.Environment.NewLine)
                TestDetails += System.Environment.NewLine
            End If
        End If
    End Sub

    Private Sub SN_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles Serial_Number_Entry.KeyDown
        'Dim format_ok As Boolean = True

        If e.KeyCode = Keys.Enter Then
            If Not Validate_SN() Then
                Serial_Number_Entry.BackColor = Color.Red
                Exit Sub
            Else
                Serial_Number_Entry.BackColor = Color.LightGreen
                Controller_Board_SN_Entry.Focus()
            End If
        End If
    End Sub

    Private Sub Controller_Board_SN_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles Controller_Board_SN_Entry.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim minLength As Integer = 8
            Dim label As String = "Controller Board"
            If AssemblyLevel_ComboBox.Text = "TM101_TM102" Then
                minLength = 15
                label = "TM102 SN"
            ElseIf AssemblyLevel_ComboBox.Text = "DISPLAY_KIT" Then
                minLength = 6
                label = "Display"
            End If

            If Controller_Board_SN_Entry.Text.Length >= minLength Then
                Controller_Board_SN_Entry.BackColor = Color.LightGreen
                PowerSupply_SN_Entry.Focus()
            Else
                AppendText(label + " SN length = " + Controller_Board_SN_Entry.Text.Length.ToString + ", expected " + minLength.ToString)
                Controller_Board_SN_Entry.BackColor = Color.Red
            End If
        End If
    End Sub

    Private Sub PowerSupply_SN_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles PowerSupply_SN_Entry.KeyDown
        If e.KeyCode = Keys.Enter Then
            If PowerSupply_SN_Entry.Text.Length = 14 Then
                PowerSupply_SN_Entry.BackColor = Color.LightGreen
                Pump_SN_Entry.Focus()
            Else
                AppendText("Power Supply SN length = " + PowerSupply_SN_Entry.Text.Length.ToString + ", expected 14")
                PowerSupply_SN_Entry.BackColor = Color.Red
            End If
        End If
    End Sub

    Private Sub Pump_SN_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles Pump_SN_Entry.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Pump_SN_Entry.Text.Length = 9 Then
                Pump_SN_Entry.BackColor = Color.LightGreen
            Else
                AppendText("Pump SN length = " + Pump_SN_Entry.Text.Length.ToString + ", expected 9")
                Pump_SN_Entry.BackColor = Color.Red
            End If
        End If
    End Sub

    Private Function Validate_SN() As Boolean
        Dim Expected_SN_Length As Integer = 15
        Dim Fields() As String
        Dim YY As Integer
        Dim WW As Integer
        Dim NNNN As Integer
        Dim TM1_SN_OKAY As Boolean

        If Serial_Number_Entry.Text = Nothing Then
            AppendText("SN not entered", False)
            Return False
        End If

        ' Note: Due to the SN variation, skip checking for controller board serial number 9/28/2106
        Serial_Number = Serial_Number_Entry.Text
        If AssemblyLevel_ComboBox.Text = "CONTROLLER_BOARD" Then
            TM1_SN = BoardSN_to_TM1SN(Serial_Number)
            Return True
        End If


        'If AssemblyLevel_ComboBox.Text = "CONTROLLER_BOARD" Then
        '    ' This is a temporary fix for the first article batch of TM1 controller boards
        '    ' whose serial numbers are non-conform to the standard format
        '    If Serial_Number_Entry.Text.Length = 10 Then
        '        Serial_Number_Entry.Text = Serial_Number_Entry.Text.Substring(2)
        '    End If
        '    Expected_SN_Length = 8
        '    'ElseIf AssemblyLevel_ComboBox.Text = "DISPLAY_KIT" Then
        '    '    Expected_SN_Length = 3
        'Else
        '    Expected_SN_Length = 15
        'End If
        'If Not Serial_Number_Entry.Text.Length = Expected_SN_Length Then
        '    AppendText("Serial Number length = " + Serial_Number_Entry.Text.Length.ToString + ", expecting " + Expected_SN_Length.ToString, False)
        '    Return False
        'End If

        'Serial_Number = Serial_Number_Entry.Text

        'If AssemblyLevel_ComboBox.Text = "CONTROLLER_BOARD" Then
        'If Not Regex.IsMatch(Serial_Number, "^[A-Z|\d]*$") Then
        '    AppendText("Serial Number should only contain upper case letters and digits")
        '    Return False
        'End If
        'TM1_SN = BoardSN_to_TM1SN(Serial_Number)
        'ElseIf AssemblyLevel_ComboBox.Text = "DISPLAY_KIT" Then
        '    TM1_SN = "TM1010113160094"
        '    Dim Price As Integer
        '    If Not Int32.TryParse(Serial_Number, Price) Or Price < 0 Or Price > 999 Then
        '        AppendText("Display Kit Serial Number should be between 0-999")
        '        Return False
        '    End If
        'Else
        TM1_SN_OKAY = True
        TM1_SN = Serial_Number
        If Not Regex.IsMatch(TM1_SN, "TM10\d01\d\d\d\d\d\d\d\d") Then
            AppendText("TM1 Serial Number does not match format 'TM10x01YYWWNNNN'")
            Return False
        End If
        Fields = Regex.Split(TM1_SN, "TM10\d01(\d\d)(\d\d)(\d\d\d\d)")
        YY = Fields(1)
        WW = Fields(2)
        NNNN = Fields(3)
        If YY < 12 Then
            TM1_SN_OKAY = False
        End If
        If WW > 52 Then
            TM1_SN_OKAY = False
        End If
        If Not TM1_SN_OKAY Then
            AppendText("TM1 Serial Number does not match format 'TM10x01YYWWNNNN'")
            Return False
        End If
        'End If

        Return True
    End Function

    ' Convert a board SN to TM1 SN
    Private Function BoardSN_to_TM1SN(ByVal SN As String) As String
        Dim Serial_Number_Array As Array
        Dim Serial_Number_TM1 As String = ""

        If SN.Length = 8 Then
            Serial_Number_Array = SN.ToCharArray()
            Serial_Number_TM1 = "TM1" +
                Serial_Number_Array(0) +
                Asc(Serial_Number_Array(1)).ToString +
                Asc(Serial_Number_Array(2)).ToString +
                Asc(Serial_Number_Array(3)).ToString +
                Asc(Serial_Number_Array(4)).ToString +
                Serial_Number_Array(5) +
                Serial_Number_Array(6) +
                Serial_Number_Array(7)
        ElseIf SN.Length = 10 And SN.StartsWith("TM1") Then
            Serial_Number_TM1 = SN + "00000"
        ElseIf SN.Length = 10 And Not SN.StartsWith("TM1") Then
            Serial_Number_TM1 = "TM1" + SN + "00"
        End If

        Return Serial_Number_TM1
    End Function

    Private Function TM1SN_to_BoardSN(ByVal SN As String) As String
        Dim Serial_Number_Board As String
        Dim Serial_Number_Array As Array

        Serial_Number_Array = SN.ToCharArray()
        Serial_Number_Board = Serial_Number_Array(3) +
            Chr(CInt(Serial_Number_Array(4) + Serial_Number_Array(5))) +
            Chr(CInt(Serial_Number_Array(6) + Serial_Number_Array(7))) +
            Chr(CInt(Serial_Number_Array(8) + Serial_Number_Array(9))) +
            Chr(CInt(Serial_Number_Array(10) + Serial_Number_Array(11))) +
            Serial_Number_Array(12) +
            Serial_Number_Array(13) +
            Serial_Number_Array(14)
        Return Serial_Number_Board
    End Function


    Private Function CheckEntries() As Boolean
        Dim Entries_OK As Boolean = True

        If Not Serial_Number_Entry.BackColor = Color.LightGreen Then
            MessageBox.Show("Please correct SN entry")
            Return False
        End If

        'If Not Tmcom1ComboBox.BackColor = Color.LightGreen Then
        '    MessageBox.Show("Please select a valid COM port for TMCOM1")
        '    Return False
        'End If

        Return True
    End Function

    Private Function CreateLocalDirs() As Boolean
        Dim Dir As String
        Dim Directores() As String = {"C:\Operations",
                              "C:\Operations\Production",
                              "C:\Operations\Production\TM1",
                              "C:\Operations\Production\TM1\Controller",
                              "C:\Operations\Production\TM1\Controller\Firmware",
                              "C:\Operations\Production\TM1\Controller\Records",
                              "C:\Operations\Production\TM1\Controller\Utilities",
                              "C:\Operations\Production\TM1\CONTROLLER_BOARD",
                              "C:\Operations\Production\TM1\SUB_ASSEMBLY",
                              "C:\Operations\Production\TM1\DISPLAY_KIT",
                              "C:\Temp"}

        'Dim Directores() As String = {"C:\Operations",
        '                      "C:\Operations\Production",
        '                      "C:\Operations\Production\TM1",
        '                      "C:\Operations\Production\TM1\Controller",
        '                      "C:\Operations\Production\TM1\Controller\Firmware",
        '                      "C:\Operations\Production\TM1\Controller\Records",
        '                      "C:\Operations\Production\TM1\Controller\Utilities",
        '                      "C:\Operations\Production\TM1\CONTROLLER_BOARD",
        '                      "C:\Operations\Production\TM1\SUB_ASSEMBLY",
        '                      "C:\Operations\Production\TM1\W223_REWORK",
        '                      "C:\Temp"}
        For Each Dir In Directores
            Try
                Directory.CreateDirectory(Dir)
            Catch ex As Exception
                Return False
            End Try
        Next

        Return True
    End Function

    Private Function ExtractFirmware() As Boolean
        Dim res() As String
        Dim fw_file As String
        Dim filename As String = ""
        Dim fFile As FileStream
        Dim swriter As Stream

        res = Me.GetType.Assembly.GetManifestResourceNames()
        For i = 0 To UBound(res)
            If res(i).EndsWith(".bin") Or res(i).EndsWith(".hex") Then
                ' fw_file = Regex.Split(res(i), "\.")(1) + ".bin"
                fw_file = Regex.Split(res(i), "\w+\.(.*)")(1)
                Try
                    filename = BaseFWDirectory + fw_file
                    fFile = New FileStream(filename, FileMode.Create)
                Catch ex As Exception
                    AppendText("Problem creating file " + filename)
                    AppendText(ex.ToString)
                    Return False
                End Try

                Try
                    swriter = Me.GetType.Assembly.GetManifestResourceStream(res(i))
                    For x = 1 To swriter.Length
                        fFile.WriteByte(swriter.ReadByte)
                    Next
                    fFile.Close()
                Catch ex As Exception
                    AppendText("Problem writing file " + filename)
                    Return False
                End Try
            End If
        Next

        Return True
    End Function

    Private Function ExtractUtilities() As Boolean
        Dim res() As String
        Dim fw_file As String
        Dim filename As String = ""
        Dim fFile As FileStream
        Dim swriter As Stream

        res = Me.GetType.Assembly.GetManifestResourceNames()
        For i = 0 To UBound(res)
            If res(i).EndsWith(".exe") Or res(i).EndsWith(".dll") Then
                fw_file = Regex.Split(res(i), "\w+\.(.*)")(1)
                Try
                    filename = UtilitiesDirectory + fw_file
                    fFile = New FileStream(filename, FileMode.Create)
                Catch ex As Exception
                    AppendText("Problem creating file " + filename)
                    AppendText(ex.ToString)
                    Return False
                End Try

                Try
                    swriter = Me.GetType.Assembly.GetManifestResourceStream(res(i))
                    For x = 1 To swriter.Length
                        fFile.WriteByte(swriter.ReadByte)
                    Next
                    fFile.Close()
                Catch ex As Exception
                    AppendText("Problem writing file " + filename)
                    Return False
                End Try
            End If
        Next

        Return True
    End Function

    Private Sub EnableTest(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles Test_1.MouseUp,
        Test_2.MouseUp, Test_3.MouseUp, Test_4.MouseUp, Test_5.MouseUp, Test_6.MouseUp, Test_7.MouseUp,
        Test_8.MouseUp, Test_9.MouseUp, Test_10.MouseUp, Test_11.MouseUp, Test_12.MouseUp, Test_13.MouseUp, Test_14.MouseUp, Test_15.MouseUp,
        Test_16.MouseUp, Test_17.MouseUp, Test_18.MouseUp, Test_19.MouseUp, Test_20.MouseUp
        If Not ListBox_TestMode.Text = "DEBUG" Then
            Exit Sub
        End If
        If e.Button = MouseButtons.Right Then
            For i = 0 To Test_Sequence.Count - 1
                If sender.Equals(Test_Sequence(i).Button) Then
                    sender.BackColor = DefaultBackColor
                    sender.Enabled = False
                End If
            Next
        End If
    End Sub

    Private Sub DisableTest(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TestButtonPanel.Click
        Dim button As Button

        If Not ListBox_TestMode.Text = "DEBUG" Then
            Exit Sub
        End If
        For i = 0 To Test_Sequence.Count - 1
            button = Test_Sequence(i).Button
            If (e.X > button.Location.X) And (e.X < button.Location.X + button.Width) And
               (e.Y > button.Location.Y) And (e.Y < button.Location.Y + button.Height) Then
                button.BackColor = Color.LightGray
                button.ForeColor = Color.Black
                button.Enabled = True
                Exit For
            End If
        Next
    End Sub

    Private Sub LoadProductInfo()
        Dim ProductInfo As Hashtable

        ProductInfo = New Hashtable
        'ProductInfo.Add("FW_VERSION", "0.5.4722")
        'ProductInfo.Add("FW_VERSION", "0.6.4749")
        'ProductInfo.Add("FW_VERSION", "0.7.4773")
        'ProductInfo.Add("FW_VERSION", "0.8.4811")
        'ProductInfo.Add("FW_VERSION", "0.9.4823")
        'ProductInfo.Add("FW_VERSION", "0.9.4862")
        'ProductInfo.Add("FW_VERSION", "0.9.4876")
        'ProductInfo.Add("FW_VERSION", "0.9.4884")
        'ProductInfo.Add("FW_VERSION", "0.9.4890")
        'ProductInfo.Add("FW_VERSION", "1.0.4898")
        'ProductInfo.Add("FW_VERSION", "1.1.5321")

        ' Current production release version
        'ProductInfo.Add("FW_VERSION", "1.2.5541")

        ' Future Version
        'ProductInfo.Add("FW_VERSION 0", "1.0.4898")
        'ProductInfo.Add("FW_VERSION 1", "1.0.5244")
        'ProductInfo.Add("sensor firmware version", "3.31")
        'ProductInfo.Add("sensor firmware checksum", "751c")
        'ProductInfo.Add("sensor firmware version", "3.35B")
        'ProductInfo.Add("sensor firmware checksum", "ab53")
        ' Future Version
        'ProductInfo.Add("sensor firmware version 0", "3.36B")
        'ProductInfo.Add("sensor firmware version 1", "3.82B")
        'ProductInfo.Add("sensor firmware checksum 0", "C2B9")
        'ProductInfo.Add("sensor firmware checksum 1", "9651")
        ' Get version info from the app.config
        Dim configValvue As String
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Try
            configValvue = Convert.ToString(configurationAppSettings.GetValue("TM1FirmwareVersion", GetType(String)))
            ProductInfo.Add("FW_VERSION", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("H2ScanFirmwareVersion", GetType(String)))
            ProductInfo.Add("sensor firmware version", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("H2ScanMatchedFirmwareVersion", GetType(String)))
            ProductInfo.Add("sensor firmware matched version", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("H2ScanFirmwareCheckSum", GetType(String)))
            ProductInfo.Add("sensor firmware checksum", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("HardwareVersion0", GetType(String)))
            ProductInfo.Add("hardware version 0", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("AssemblyVersion0", GetType(String)))
            ProductInfo.Add("assembly version 0", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("HardwareVersion1", GetType(String)))
            ProductInfo.Add("hardware version 1", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("AssemblyVersion1", GetType(String)))
            ProductInfo.Add("assembly version 1", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("HardwareVersion2", GetType(String)))
            ProductInfo.Add("hardware version 2", configValvue)
            configValvue = Convert.ToString(configurationAppSettings.GetValue("AssemblyVersion2", GetType(String)))
            ProductInfo.Add("assembly version 2", configValvue)
        Catch ex As Exception
            MessageBox.Show("LoadProductInfo(): Caught exception " + ex.Message)
        End Try
        Products.Add("STANDARD", ProductInfo)

        ' Current version
        'ProductInfo.Add("sensor firmware version", "3.955B")
        'ProductInfo.Add("sensor firmware checksum", "5E37")
        ''ProductInfo.Add("sensor firmware version", "3.36B")
        ''ProductInfo.Add("sensor firmware checksum", "C2B9")

        'ProductInfo.Add("hardware version 0", "0x0")
        'ProductInfo.Add("hardware version 1", "0x1")
        'ProductInfo.Add("hardware version 2", "0x2")
        'ProductInfo.Add("assembly version 0", "0x0")
        'ProductInfo.Add("assembly version 1", "0x1")
        'ProductInfo.Add("assembly version 2", "0x2")
        'Products.Add("STANDARD", ProductInfo)
        TargetFirmwareVersion_Label.Text = Products("STANDARD")("FW_VERSION")
    End Sub

    Private Sub LedFlasher_Tick(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub StopButton_Click(sender As System.Object, e As System.EventArgs) Handles StopButton.Click
        StopButton.Enabled = False
        Stopped = True
        TestStatus.Text = "STOPPING"
        TestStatus.BackColor = Color.OrangeRed
        TestStatus.ForeColor = Color.Black
        If LoggingData Then
            ReportFileOpen = False
            ReportFileWriter.Close()
            ReportFile.Close()
            LoggingData = False
        End If
    End Sub

    Private Sub AssemblyLevel_ComboBox_SelectedValueChanged(sender As System.Object, e As System.EventArgs) Handles AssemblyLevel_ComboBox.SelectedValueChanged

        AssemblyLevel = AssemblyLevel_ComboBox.Text
        InitializeTestSequence()
        EnableEntries()
        SetSpecs()

        Select Case AssemblyLevel
            Case "CONTROLLER_BOARD"
                ComPortLabel.Text = "TMCOM1 COM Port"
                Label1.Text = "Serial Number"
                Label_BK1856D.Visible = True
                BK1856D_com_ComboBox.Visible = True
                Controller_Board_SN_Entry.Visible = False
                Controller_Board_SN_Label.Visible = False
                Pump_SN_Entry.Visible = False
                Pump_SN_Label.Visible = False
                PowerSupply_SN_Entry.Visible = False
                Power_Supply_SN_Label.Visible = False
                USB1208LS_ComboBox.Visible = True
                Label_USB1208LS.Visible = True
            Case "LEAK"
                ComPortLabel.Text = "LEAK FIX COM Port"
                Label_BK1856D.Visible = False
                BK1856D_com_ComboBox.Visible = False
                Controller_Board_SN_Entry.Visible = False
                Controller_Board_SN_Label.Visible = False
                Pump_SN_Entry.Visible = False
                Pump_SN_Label.Visible = False
                PowerSupply_SN_Entry.Visible = False
                Power_Supply_SN_Label.Visible = False
                USB1208LS_ComboBox.Visible = False
                Label_USB1208LS.Visible = False
            Case "DISPLAY_KIT"
                ComPortLabel.Text = "Main COM Port"
                Label_BK1856D.Visible = False
                BK1856D_com_ComboBox.Visible = False
                Controller_Board_SN_Entry.Visible = True
                Controller_Board_SN_Label.Text = "Display Kit SN"
                Controller_Board_SN_Label.Visible = True
                Pump_SN_Entry.Visible = False
                Pump_SN_Label.Visible = False
                PowerSupply_SN_Entry.Visible = False
                Power_Supply_SN_Label.Visible = False
                USB1208LS_ComboBox.Visible = False
                Label_USB1208LS.Visible = False
            Case "SUB_ASSEMBLY", "W223_REWORK"
                Label1.Text = "Serial Number"
                Controller_Board_SN_Label.Text = "Controller Board SN"
                ComPortLabel.Text = "TMCOM1 COM Port"
                Label_BK1856D.Visible = False
                BK1856D_com_ComboBox.Visible = False
                Controller_Board_SN_Entry.Visible = True
                Controller_Board_SN_Label.Visible = True
                Pump_SN_Entry.Visible = True
                Pump_SN_Label.Visible = True
                PowerSupply_SN_Entry.Visible = True
                Power_Supply_SN_Label.Visible = True
                USB1208LS_ComboBox.Visible = True
                Label_USB1208LS.Visible = True
            Case "TM101_TM102"
                Label_BK1856D.Visible = False
                BK1856D_com_ComboBox.Visible = False
                Label1.Text = "TM101 SN"
                ComPortLabel.Visible = False
                Tmcom1ComboBox.Visible = False
                Controller_Board_SN_Entry.Visible = True
                Controller_Board_SN_Label.Visible = True
                Controller_Board_SN_Label.Text = "TM102 SN"
                Pump_SN_Entry.Visible = False
                Pump_SN_Label.Visible = False
                PowerSupply_SN_Entry.Visible = False
                Power_Supply_SN_Label.Visible = False
                USB1208LS_ComboBox.Visible = False
                Label_USB1208LS.Visible = False
        End Select

        If Not AssemblyLevel = "TM101_TM102" Then
            ComPortLabel.Visible = True
            Tmcom1ComboBox.Visible = True
        End If

        If USB1208LS_ComboBox.Visible Then
            FindDaqs()
        End If
    End Sub

    Private Function CheckEntriesForTest(ByVal Test As Test_Item)
        Dim EntriesNeeded As New Hashtable
        Dim FailureMessage As String = ""
        Dim EntriesOkay As Boolean = True

        EntriesNeeded = Test.EntriesNeeded()
        If EntriesNeeded("Serial Number") And Serial_Number_Entry.Text = Nothing Then
            FailureMessage += "'Serial Number' not defined" + System.Environment.NewLine
            EntriesOkay = False
        End If
        If EntriesNeeded("Controller Board Serial Number") And Controller_Board_SN_Entry.Text = Nothing Then
            FailureMessage += "'Controller Board SN' not defined" + System.Environment.NewLine
            EntriesOkay = False
        End If
        If EntriesNeeded("Power Supply Serial Number") And PowerSupply_SN_Entry.Text = Nothing Then
            FailureMessage += "'Power Supply SN' not defined" + System.Environment.NewLine
            EntriesOkay = False
        End If
        If EntriesNeeded("Pump Serial Number") And Pump_SN_Entry.Text = Nothing Then
            FailureMessage += "'Pump SN' not defined" + System.Environment.NewLine
            EntriesOkay = False
        End If
        If EntriesNeeded("TMCOM1 COM Port") And Tmcom1ComboBox.Text = Nothing Then
            FailureMessage += "TMCOM1 COM Port not defined" + System.Environment.NewLine
            EntriesOkay = False
        End If
        If EntriesNeeded("BK1856D COM Port") And BK1856D_com_ComboBox.Text = Nothing Then
            FailureMessage += "BK1856D COM Port not defined" + System.Environment.NewLine
            EntriesOkay = False
        End If
        If EntriesNeeded("USB1208LS") And USB1208LS_ComboBox.Text = Nothing Then
            FailureMessage += "USB1208-LS device not selected" + System.Environment.NewLine
            EntriesOkay = False
        End If

        If Not EntriesOkay Then
            MessageBox.Show(FailureMessage, "MISSING ENTRIES")
        End If

        Return EntriesOkay
    End Function

    Private Sub EnableAllTest_Click(sender As System.Object, e As System.EventArgs) Handles EnableAllTest.Click
        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Enabled = True
            Test_Sequence(i).Button.BackColor = Color.LightGray
            Test_Sequence(i).Button.ForeColor = Color.Black
        Next
    End Sub

    Private Sub DisableAllTests_Click(sender As System.Object, e As System.EventArgs) Handles DisableAllTests.Click
        For i = 0 To Test_Sequence.Count - 1
            Test_Sequence(i).Button.Enabled = False
            Test_Sequence(i).Button.BackColor = DefaultBackColor
        Next
    End Sub

    Private Sub ListBox_TestMode_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox_TestMode.SelectedIndexChanged
        If ListBox_TestMode.SelectedIndex = 1 Then
            DisableAllTests.Visible = True
            EnableAllTest.Visible = True
        Else
            DisableAllTests.Visible = False
            EnableAllTest.Visible = False
        End If
    End Sub

    Public Function GetControl() As Control
        Return Me
    End Function

    Private Sub Tmcom1ComboBox_DropDown(sender As System.Object, e As System.EventArgs) Handles Tmcom1ComboBox.DropDown, BK1856D_com_ComboBox.DropDown

        Tmcom1ComboBox.Items.Clear()
        BK1856D_com_ComboBox.Items.Clear()
        For Each Str As String In System.IO.Ports.SerialPort.GetPortNames()
            If Not BK1856D_com_ComboBox.Text = Str Then
                Tmcom1ComboBox.Items.Add(Str)
            End If
            If Not Tmcom1ComboBox.Text = Str Then
                BK1856D_com_ComboBox.Items.Add(Str)
            End If
        Next
    End Sub

    Private Sub SaveSerialPortConfig()
        Dim SerialConfigFile As FileStream
        Dim SerialConfigFileWriter As StreamWriter

        If (Not Tmcom1ComboBox.Text Is Nothing) And
            (Not BK1856D_com_ComboBox.Text Is Nothing) Then
            Try
                SerialConfigFile = New FileStream("\Temp\TM1_Controller_Test_SerialConfig.txt", FileMode.Create, FileAccess.Write)
                SerialConfigFileWriter = New StreamWriter(SerialConfigFile)
                SerialConfigFileWriter.WriteLine("TMCOM1:" + Tmcom1ComboBox.Text + System.Environment.NewLine +
                                                 "BK1856D:" + BK1856D_com_ComboBox.Text + System.Environment.NewLine)
                SerialConfigFileWriter.Close()
                SerialConfigFile.Close()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Function ComboBoxContains(ByVal CB As ComboBox, ByVal Name As String)

        For Each Str As String In System.IO.Ports.SerialPort.GetPortNames()
            If Str = Name Then
                Return True
            End If
        Next
        'For Each item As Object In CB.Items
        '    If item.ToString = Name Then
        '        Return True
        '    End If
        'Next

        Return False
    End Function

    Private Sub AboutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim About1 As New About
        About1.Show()
    End Sub

    Private Sub FindDaqs()
        'Dim D As New DAQ_functions
        Dim USB1208LS_SN As String
        Dim Found_USB1208 As Boolean

        USB1208LS_ComboBox.Items.Clear()
        USB1208LS_ComboBox.Text = ""
        AppendText("Identifying installed USB1208-LS devices")
        'If Not D.FindDAQs(USB1208LS_Devices) Then
        '    AppendText(D.ErrorMsg)
        '    MsgBox("Problem detecting USB1208-LS DAQ device" + vbCr + vbCr + D.ErrorMsg)
        '    Exit Sub
        'End If

        Dim RDC As New RemoteDaqClient
        If Not RDC.LockDaq Then
            AppendText(RDC.ErrorMsg)
            Exit Sub
        End If
        If RDC.FindDAQs(USB1208LS_Devices) Then
            Found_USB1208 = True
        Else
            AppendText("Problem detecting USB1208-LS DAQ device" + vbCr + vbCr)
            AppendText(RDC.ErrorMsg)
        End If
        If Not RDC.UnlockDaq Then
            AppendText(RDC.ErrorMsg)
            Exit Sub
        End If
        If Not Found_USB1208 Then
            Exit Sub
        End If


        'Dim DB As New DaqBackground
        'DB.Start()
        'If Not DB.Command("FIND") Then
        '    AppendText(DB.ErrorMsg)
        '    MsgBox("Problem detecting USB1208-LS DAQ device" + vbCr + vbCr)
        'End If
        'DB.StopDaq = True

        AppendText("DAQ count = " + USB1208LS_Devices.Count.ToString)
        For Each USB1208LS_SN In USB1208LS_Devices.Keys
            USB1208LS_ComboBox.Items.Add(USB1208LS_SN)
        Next
        USB1208LS_ComboBox.Enabled = True

    End Sub

    Private Sub USB1208LS_ComboBox_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles USB1208LS_ComboBox.SelectedIndexChanged
        AppendText(USB1208LS_ComboBox.Text)
        If USB1208LS_ComboBox.SelectedIndex < 0 Then
            MenuStrip1.BackColor = Control.DefaultBackColor
        Else
            MenuStrip1.BackColor = ColorTranslator.FromHtml(Module1.AppHexColorStr(USB1208LS_ComboBox.SelectedIndex))
        End If
    End Sub

    Private Sub ButtonClearSNs_Click(sender As System.Object, e As System.EventArgs) Handles ButtonClearSNs.Click
        ClearSerialNumbers()
    End Sub

    Private Sub ClearSerialNumbers()
        Serial_Number_Entry.Text = ""
        Serial_Number_Entry.BackColor = Color.Empty
        Serial_Number_Entry.Focus()

        Controller_Board_SN_Entry.Text = ""
        Controller_Board_SN_Entry.BackColor = Color.Empty

        PowerSupply_SN_Entry.Text = ""
        PowerSupply_SN_Entry.BackColor = Color.Empty

        Pump_SN_Entry.Text = ""
        Pump_SN_Entry.BackColor = Color.Empty

    End Sub
End Class

Friend Class MyComboBox
    Inherits ComboBox

    Protected Overrides Sub OnMouseWheel(ByVal e As MouseEventArgs)
        Dim mwe As HandledMouseEventArgs = DirectCast(e, HandledMouseEventArgs)
        mwe.Handled = True
    End Sub
End Class
'                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 