VERSION 5.00
Object = "{02E0654E-AAC5-4BBF-A1DE-45576B24DFC1}#2.1#0"; "ProXR.ocx"
Begin VB.Form frmSelectDevice 
   Caption         =   "Select Device Type"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Height          =   1335
      Left            =   2040
      TabIndex        =   28
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label Label1 
         Caption         =   "Detecting Devices, Please wait....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.CommandButton AutoDetect 
      Caption         =   "Automatically Detect Wireless Modem Type"
      Height          =   975
      Left            =   4800
      TabIndex        =   27
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "PWM8x"
      Height          =   375
      Index           =   13
      Left            =   480
      TabIndex        =   26
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "868 ProXR"
      Enabled         =   0   'False
      Height          =   375
      Index           =   12
      Left            =   7200
      TabIndex        =   25
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "868R1x/868R2x"
      Enabled         =   0   'False
      Height          =   375
      Index           =   11
      Left            =   7200
      TabIndex        =   24
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Frame Frame7 
      Caption         =   "XBee-Pro 868 Long Range"
      Height          =   1095
      Left            =   6960
      TabIndex        =   23
      Top             =   2160
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMR ProXR"
      Height          =   375
      Index           =   10
      Left            =   5040
      TabIndex        =   22
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DMR1x/DMR2x"
      Height          =   375
      Index           =   9
      Left            =   5040
      TabIndex        =   21
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      Caption         =   "DigiMesh 900"
      Height          =   1095
      Left            =   4800
      TabIndex        =   20
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "XSCR1x/XSCR2x"
      Height          =   375
      Index           =   7
      Left            =   7200
      TabIndex        =   19
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "XSC ProXR"
      Height          =   375
      Index           =   8
      Left            =   7200
      TabIndex        =   18
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "XSC Long Range"
      Height          =   1095
      Left            =   6960
      TabIndex        =   17
      Top             =   960
      Width           =   2175
   End
   Begin NCDProXR.ProXR ProXR1 
      Left            =   8640
      Top             =   120
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin VB.TextBox HelpText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   3360
      Width           =   6615
   End
   Begin VB.Frame Frame4 
      Caption         =   "Com Port Setting"
      Height          =   800
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cmbPorts 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   300
         Width           =   1335
      End
      Begin VB.ComboBox cmbBaudrate 
         Height          =   315
         ItemData        =   "frmSelectDeviceType.frx":0000
         Left            =   4440
         List            =   "frmSelectDeviceType.frx":0002
         TabIndex        =   12
         Text            =   "9600"
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lbBaudrate 
         Caption         =   "Baudrate"
         Height          =   240
         Left            =   3480
         TabIndex        =   15
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Com Port"
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ZB ProXR"
      Height          =   375
      Index           =   6
      Left            =   2760
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ZBR1x/ZBR2x"
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "XBee ProXR"
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ZR1x/ZR2x"
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ProXR"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "R4xPro/R8xPro"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   6615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "R1x/R2x"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Most NCD Devices"
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "XBee/XBee Pro Series 1"
      Height          =   1095
      Left            =   2520
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "ZB Series 2 Mesh Devices"
      Height          =   1095
      Left            =   2520
      TabIndex        =   10
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Image ImageDigiMesh 
      Height          =   2940
      Left            =   6840
      Picture         =   "frmSelectDeviceType.frx":0004
      Top             =   3360
      Width           =   2235
   End
   Begin VB.Image ImageXSC 
      BorderStyle     =   1  'Fixed Single
      Height          =   3030
      Left            =   6840
      Picture         =   "frmSelectDeviceType.frx":A60E
      Top             =   3360
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Image ImageS2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   6840
      Picture         =   "frmSelectDeviceType.frx":14FC5
      Top             =   3360
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Image ImageS1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3030
      Left            =   6840
      Picture         =   "frmSelectDeviceType.frx":1DB49
      Top             =   3360
      Visible         =   0   'False
      Width           =   2310
   End
End
Attribute VB_Name = "frmSelectDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' define line data info for read line function
Private Type LineDataInfo
    content As String       ' content from serial port
    timeout As Boolean      ' flag for timeout, if ture, the content will be invalid
End Type

Public SelectOKButton As Boolean

Private Function GetSelect() As Integer
    For i = 0 To 13
        If (Option1(i).Value = True) Then
            SaveSetting "ProXR_V4", "Settings", "DeviceType", i
            SaveSetting "ProXR_V4", "Settings", "ComPort", cmbPorts.Text
            SaveSetting "ProXR_V4", "Settings", "Baudrate", cmbBaudrate.Text
        End If
    Next
End Function

Private Sub AutoSelect()
    ProXR1.PortName = cmbPorts.Text
    ProXR1.BaudRate = br_9600
    ProXR1.OpenPort
    ProXR1.ClearBuffer
    ProXR1.Sleep 3000
    SendString "+++", False
    Dim info As LineDataInfo
    ProXR1.Sleep 3000
    info = ReadLine(3000)
    
    If info.timeout Then
        ProXR1.BaudRate = br_115200
        ProXR1.ClosePort
        ProXR1.OpenPort
        ProXR1.ClearBuffer
        SendString "+++", False
        ProXR1.Sleep 3000
        info = ReadLine(3000)
    End If
    If info.timeout Then
        ProXR1.ClosePort
        MsgBox "Fail to detect wireless device type"
        Exit Sub
    End If
    
    If info.content <> "OK" Then
        ProXR1.ClosePort
        MsgBox "Fail to detect wireless device type"
        Exit Sub
    End If
    
    SendString "ATVR", True
    ProXR1.Sleep 2000
    info = ReadLine(3000)
    If info.timeout Then
        ProXR1.ClosePort
        MsgBox "Fail to detect wireless device type"
        Exit Sub
    End If
    
    SelectDevice info.content
    SendString "ATCN", True
End Sub

Private Sub AutoDetect_Click()
    
    frmSelectDevice.Enabled = False
    curWirelessDevice = -1
    Frame8.Visible = True
    EnableAllTypes True
    AutoSelect
    Frame8.Visible = False

    frmSelectDevice.Enabled = True
End Sub

Private Sub SelectDevice(ByVal ver As String)
    Select Case ver
        Case "1081", "1083", "1084", "10A1", "10A2", "10A4", "10A5", "10CD"
            SelectXBee
        Case "2021", "2041", "2821", "2841", "2221", "2241"
            SelectZBSeries2Mesh
        Case "1820"
            SelectDigiMesh900
        Case "1020"
            SelectXSCLongRange
    End Select
End Sub

Private Sub EnableAllTypes(ByVal enable As Boolean)
    Frame2.Enabled = enable
    Frame3.Enabled = enable
    Frame5.Enabled = enable
    Frame6.Enabled = enable
    For i = 3 To 10
        Option1(i).Enabled = enable
    Next
End Sub

Private Sub SelectXBee()
    EnableAllTypes False
    Frame2.Enabled = True
    Option1(3).Enabled = True
    Option1(4).Enabled = True
    Option1(3).Value = True
    RefreshPicture 1
End Sub

Private Sub SelectZBSeries2Mesh()
    EnableAllTypes False
    Frame3.Enabled = True
    Option1(5).Enabled = True
    Option1(6).Enabled = True
    Option1(5).Value = True
    RefreshPicture 2
End Sub

Private Sub SelectDigiMesh900()
    EnableAllTypes False
    Frame6.Enabled = True
    Option1(9).Enabled = True
    Option1(10).Enabled = True
    Option1(9).Value = True
    RefreshPicture 4
End Sub

Private Sub SelectXSCLongRange()
    EnableAllTypes False
    Frame5.Enabled = True
    Option1(7).Enabled = True
    Option1(8).Enabled = True
    Option1(7).Value = True
    RefreshPicture 3
End Sub

Private Sub AutoDetect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HelpText.Text = "If you are not sure about which wireless modem you are using, you can click this button to help narrow down the possible choices."
End Sub

Private Sub cmdSelect_Click()
    GetSelect

    s = GetSetting("ProXR_V4", "Settings", "DeviceType", 0)
    Dim i As Integer
    i = Val(s)
    Select Case i
        Case MostNcdDevicesR2X
            ShowMostNcdDevicesR2X
        Case MostNcdDevicesR8X
            ShowMostNcdDevicesR8X
        Case MostNcdDevicesProXr
            ShowMostNcdDevicesProXr
        Case ZigBeeS1ZR2X
            ShowZigBeeS1ZR2X
        Case ZigBeeS1ProXr
            ShowZigBeeS1ProXr
        Case ZigBeeMeshR2X
            ShowZigBeeMeshR2X
        Case ZigBeeMeshProXr
            ShowZigBeeMeshProXr
        Case XscLongRangeR2X
            ShowXscLongRangeR2X
        Case XscLongRangeProXr
            ShowXscLongRangeProXr
        Case DigiMesh900R2X
            ShowDigiMesh900R2X
        Case DigiMesh900ProXr
            ShowDigiMesh900ProXr
        Case MostNcdPWM8x
            ShowMostNcdPWM8x
    End Select

End Sub

Private Sub Form_Load()
    For i = 0 To 13
         Option1(i).Value = False
    Next
    SelectOKButton = False
    Dim s As String
    s = GetSetting("ProXR_V4", "Settings", "DeviceType", 0)
    i = Val(s)
    Option1(i).Value = True
    Debug.Print s
    Dim ports As Collection
    Set ports = ProXR1.EnumComPorts
    cmbPorts.Clear
    For Each obj In ports
        cmbPorts.AddItem obj
    Next
    
    Dim strDefaultPort As String
    strDefaultPort = "COM1"
    If (ports.Count > 0) Then
        strDefaultPort = ports(1)
    End If
    s = GetSetting("ProXR_V4", "Settings", "ComPort", strDefaultPort)
    cmbPorts.Text = s
    
    Dim strDefaultBaudrate As String
    strDefaultBaudrate = "9600"
    s = GetSetting("ProXR_V4", "Settings", "Baudrate", strDefaultBaudrate)
    cmbBaudrate.Text = s
    cmbBaudrate.AddItem "9600"
    cmbBaudrate.AddItem "19200"
    cmbBaudrate.AddItem "38400"
    cmbBaudrate.AddItem "57600"
    cmbBaudrate.AddItem "115200"
    
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HelpText.Text = "Choose between the controllers in this category if you are NOT using any form of ZigBee Wireless communications.  Choose R1x/R2x if you are using single or dual channel RS-232 relay controllers.  Choose R4xPro/R8xPro if you are using RS-232 4-Channel/8-Channel Relay Controllers.  Choose ProXR if you are using any ProXR series controller (other than ZigBee Wireless/Mesh/XSC series)."
    RefreshPicture 0
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HelpText.Text = "Choose between the controllers in this category if you are using XBee or XBee Pro Series Point-to-Point 802.15.4 (Series 1) Wireless Relay Devices."
    RefreshPicture 1
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HelpText.Text = "Choose between the controllers in this category if you are using XBee or XBee Pro ZB (Series 2) Mesh Networking Devices (NCD is NOT currently supporting DigiMesh 2.4 or ZNet 2.5 Devices, ONLY ZigBee 2007 specification, and is NOT over the air compatible with EmberZNet 2.5 devices)."
    RefreshPicture 2
End Sub
Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HelpText.Text = "Choose between the controllers in this category if you are using XSC Series Long Range Devices.  These devices look VERY similar to the DigiMesh 900 version, but the communication protocols are different, and will not work with each other."
    RefreshPicture 3
End Sub
Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HelpText.Text = "Choose between the controllers in this category if you are using DigiMesh 900 Series Long Range Mesh Networking Devices.  These devices look VERY similar to the XSC version, but the communication protocols are different, and will not work with each other."
    RefreshPicture 4
End Sub
Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HelpText.Text = "Coming Soon! These Devices will be supported in the next release of our ProXR software."
End Sub

Private Sub Option1_Click(Index As Integer)
    If MostNcdDevicesR8X = Index Then
        lbBaudrate.Visible = True
        cmbBaudrate.Visible = True
    Else
        lbBaudrate.Visible = False
        cmbBaudrate.Visible = False
    End If
End Sub
Private Sub Option1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 Then HelpText.Text = "Choose R1x/R2x if you are using a RS-232 Single or Dual Channel Relay Controller."
    If Index = 1 Then HelpText.Text = "Choose R4xPro/R8xPro if you are using a RS-232 4/8 Channel Relay Pro Series Relay Controller WITHOUT Inputs or UXP Expansion Ports.  Do NOT use this option if you are using a 4/8 Channel ProXR Series Relay Controller."
    If Index = 2 Then HelpText.Text = "Choose ProXR if you are using a USB, RS-232, Ethernet, WiFi, Fiber Optic, or Bluetooth ProXR Series Relay Controllers."
    If Index = 3 Then HelpText.Text = "Choose ZR1x/ZR2x if you are using a XBee or XBee Pro (also known as Series 1) Point-to-Point ZigBee Wireless Relay Controller WITHOUT Mesh Network Support."
    If Index = 4 Then HelpText.Text = "Choose ProXR if you are using a ProXR Series Relay Controller Equipped with ZigBee XBee or XBee Pro (also known as Series 1) Point-to-Point Controller WITHOUT Mesh Network Support."
    If Index = 5 Then HelpText.Text = "Choose ZR1x/ZR2x if you are using a Series 2 Mesh Network ZB Wireless Relay Controller.  This Option should not be used with Series 2 XBee Devices running ZNET 2.5 Firmware."
    If Index = 6 Then HelpText.Text = "Choose ProXR if you are using a ProXR Series 2 Mesh Network ZB Wireless Relay Controller.  This Option should not be used with Series 2 XBee Devices running ZNET 2.5 Firmware."
    If Index = 7 Then HelpText.Text = "Choose XSCR1x/XSCR2x if you are using Long Range 1-Channel or 2-Channel Relay Controllers.  These modules look similar to the DigiMesh 900 Series, but they do not work together."
    If Index = 8 Then HelpText.Text = "Choose XSCProXR if you are using Long Range ProXR Series Controllers.  These modules look similar to the DigiMesh 900 Series, but they do not work together."
    If Index = 9 Then HelpText.Text = "Choose DMR1x/DMR2x if you are using Long Range DigiMesh Networking Module with our 1-Channel/2-Channel Relay Controllers.  These modules look similar to the XSC, but they do not work together."
    If Index = 10 Then HelpText.Text = "Choose DMR ProXR if you are using Long Range DigiMesh Networking Modules with our ProXR Series Controllers.  These modules look similar to the XSC, but they do not work together."
    If Index = 13 Then HelpText.Text = "Choose PWM8x if you are using the PWM81, PWM83, or PWM85 Series Controllers."
End Sub

Private Sub RefreshPicture(ByVal id As Integer)
    If id = 1 Then
        ImageS1.Visible = True
    Else
        ImageS1.Visible = False
    End If
    
    If id = 2 Then
        ImageS2.Visible = True
    Else
        ImageS2.Visible = False
    End If
    
    If id = 3 Then
        ImageXSC.Visible = True
    Else
        ImageXSC.Visible = False
    End If
    
    If id = 4 Then
        ImageDigiMesh.Visible = True
    Else
        ImageDigiMesh.Visible = False
    End If
End Sub



Sub ShowMostNcdDevicesR2X()
    Set frmR2X.parentForm = Me
    frmR2X.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmR2X.ProXR1.BaudRate = 9600
    frmR2X.ProXR1.OpenPort
    Me.Hide
    frmR2X.Show
End Sub

Sub ShowMostNcdDevicesR8X()
    Set frmR8X.parentForm = Me
    frmR8X.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmR8X.ProXR1.BaudRate = Val(frmSelectDevice.cmbBaudrate.Text)
    frmR8X.ProXR1.OpenPort
    Me.Hide
    frmR8X.Show
End Sub

Sub ShowMostNcdDevicesProXr()
    Set Form1.parentForm = Me
    Form1.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    Form1.ProXR1.BaudRate = br_115200
    Form1.ProXR1.OpenPort
    Me.Hide
    Form1.Show
    Form1.LoadForm
End Sub

Sub ShowZigBeeS1ZR2X()
    
    frmSelectZigbeeS1Device.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeS1Device.ProXR1.BaudRate = br_115200
    frmSelectZigbeeS1Device.ProXR1.OpenPort
    Me.Hide
    frmSelectZigbeeS1Device.IsProXRForm = False
    frmSelectZigbeeS1Device.Show
'    frmSelectZigbeeS1Device.ProXR1.ClosePort
'    Dim selectOK As Boolean
'    selectOK = frmSelectZigbeeS1Device.SelectOKButton
'    Unload frmSelectZigbeeS1Device
'    If Not selectOK Then Exit Sub
'
'
'
'    frmR2X.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
'    frmR2X.ProXR1.BaudRate = 115200
'    frmR2X.ProXR1.OpenPort
'    frmR2X.Show
'    Unload frmR2X
End Sub

Sub ShowZigBeeS1ProXr()
    frmSelectZigbeeS1Device.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeS1Device.ProXR1.BaudRate = br_115200
    frmSelectZigbeeS1Device.ProXR1.OpenPort
    Me.Hide
    frmSelectZigbeeS1Device.IsProXRForm = True
    frmSelectZigbeeS1Device.Show
'    frmSelectZigbeeS1Device.ProXR1.ClosePort
'    Dim selectOK As Boolean
'    selectOK = frmSelectZigbeeS1Device.SelectOKButton
'    Unload frmSelectZigbeeS1Device
'    If Not selectOK Then Exit Sub
'
'
'    ShowMostNcdDevicesProXr
End Sub

Sub ShowZigBeeMeshR2X()
    frmSelectZigbeeDevice.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeDevice.ProXR1.BaudRate = br_115200
    frmSelectZigbeeDevice.ProXR1.OpenPort
    Me.Hide
    frmSelectZigbeeDevice.IsProXRForm = False
    frmSelectZigbeeDevice.Show
'    frmSelectZigbeeDevice.ProXR1.ClosePort
'    Dim selectOK As Boolean
'    selectOK = frmSelectZigbeeDevice.SelectOKButton
'    Unload frmSelectZigbeeDevice
'    If Not selectOK Then Exit Sub
'
'    frmR2X.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
'    frmR2X.ProXR1.BaudRate = 115200
'    frmR2X.ProXR1.OpenPort
'    frmR2X.Show
'    Unload frmR2X
End Sub

Sub ShowZigBeeMeshProXr()
    frmSelectZigbeeDevice.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeDevice.ProXR1.BaudRate = br_115200
    frmSelectZigbeeDevice.ProXR1.OpenPort
    Me.Hide
    frmSelectZigbeeDevice.IsProXRForm = True
    frmSelectZigbeeDevice.Show
'    frmSelectZigbeeDevice.ProXR1.ClosePort
'    Dim selectOK As Boolean
'    selectOK = frmSelectZigbeeDevice.SelectOKButton
'    Unload frmSelectZigbeeDevice
'    If Not selectOK Then Exit Sub
'
'    Form1.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
'    Form1.ProXR1.BaudRate = br_115200
'    Form1.ProXR1.OpenPort
'    Form1.LoadForm
'    Form1.Show
End Sub

Sub ShowXscLongRangeR2X()
    frmSelectXSC.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectXSC.ProXR1.BaudRate = 9600
    frmSelectXSC.ProXR1.OpenPort
    Me.Hide
    frmSelectXSC.IsProXRForm = False
    frmSelectXSC.Show
'    Dim selectOK As Boolean
'    selectOK = frmSelectXSC.SelectOKButton
'    Unload frmSelectXSC
'    If Not selectOK Then Exit Sub
'    ShowZigBeeS1ZR2X

End Sub
        
Sub ShowXscLongRangeProXr()
    frmSelectXSC.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectXSC.ProXR1.BaudRate = 9600
    frmSelectXSC.ProXR1.OpenPort
    Me.Hide
    frmSelectXSC.IsProXRForm = True
    frmSelectXSC.Show
'    frmSelectXSC.ProXR1.ClosePort
'    Dim selectOK As Boolean
'    selectOK = frmSelectXSC.SelectOKButton
'    Unload frmSelectXSC
'    If Not selectOK Then Exit Sub
'    ShowZigBeeS1ZR2X
End Sub

Sub ShowDigiMesh900R2X()
    frmSelectZigbeeDevice.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeDevice.ProXR1.BaudRate = 9600
    frmSelectZigbeeDevice.ProXR1.OpenPort
    Me.Hide
    frmSelectZigbeeDevice.IsProXRForm = False
    frmSelectZigbeeDevice.Caption = "Select Digi Mesh 900 Devices"
    frmSelectZigbeeDevice.ImageDigiMesh.Visible = True
    frmSelectZigbeeDevice.Show

End Sub

Sub ShowDigiMesh900ProXr()
    frmSelectZigbeeDevice.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeDevice.ProXR1.BaudRate = 9600
    frmSelectZigbeeDevice.ProXR1.OpenPort
    Me.Hide
    frmSelectZigbeeDevice.IsProXRForm = True
    frmSelectZigbeeDevice.Caption = "Select Digi Mesh 900 Devices"
    frmSelectZigbeeDevice.ImageDigiMesh.Visible = True
    frmSelectZigbeeDevice.Show

End Sub

Sub ShowMostNcdPWM8x()
    Set frmPwm8x.parentForm = Me
    frmPwm8x.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmPwm8x.ProXR1.BaudRate = 9600
    frmPwm8x.ProXR1.OpenPort
    Me.Hide
    frmPwm8x.Show
End Sub


Private Sub SendString(str As String, sendCarriage As Boolean)
    Dim l As Integer
    l = Len(str)
    Dim i As Integer
    For i = 1 To l
        PutData Asc(Mid(str, i, 1))
    Next
    If sendCarriage Then
        PutData 13
    End If
End Sub

' read a line from serila port
' will read by reading a carriage return character or timeout
' the string return won't contain carriage return
Private Function ReadLine(ByVal timeout As Integer) As LineDataInfo
    ReadLine.timeout = False
    Dim ln As String
    ln = ""
    Dim i As Integer
    i = GetData
    Do Until i = -1 Or i = 13
        ln = ln + Chr(i)
        'i = ProXR1.GetData2(timeout)
        i = GetData
    Loop
    If i = -1 Then ReadLine.timeout = True
    ReadLine.content = ln
    Debug.Print ln
End Function


Private Sub PutData(ByVal bt As Integer)
    If Not ProXR1.IsPortOpen Then ProXR1.OpenPort
    ProXR1.SendData bt
End Sub

Private Function GetData() As Integer
    If Not ProXR1.IsPortOpen Then ProXR1.OpenPort
    GetData = ProXR1.GetData
End Function


