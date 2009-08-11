VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{02E0654E-AAC5-4BBF-A1DE-45576B24DFC1}#2.1#0"; "ProXR.ocx"
Begin VB.Form frmSelectZigbeeDevice 
   Caption         =   "Select Zigbee Device"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowAll 
      Caption         =   "Show All"
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Device"
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   6975
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Text            =   "Name"
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtSerialLow 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Text            =   "A300000"
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtSerialHigh 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Text            =   "A300000"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Image ImageDigiMesh 
         Height          =   2955
         Left            =   4680
         Picture         =   "frmSelectZigbeeDevice.frx":0000
         Top             =   120
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Image Image1 
         Height          =   2505
         Left            =   4680
         Picture         =   "frmSelectZigbeeDevice.frx":82CD
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Serial Number Low"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Serial Number High"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
   End
   Begin NCDProXR.ProXR ProXR1 
      Left            =   6360
      Top             =   2040
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdEditName 
      Caption         =   "Edit Name"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelectZigbeeDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const AllXBeeListFile = "AllXBeeS2.txt"
Const ShowXBeeListFile = "ShowXBeeS2.txt"
Const BoardCaseItemStr = "00000000:0000FFFF Broadcast to all devices"
Const KeyValue = "ZBeeS2Settings"

Public IsProXRForm  As Boolean

Private Type ZigBeeInfo
    MY As String
    SH As String
    SL As String
    NI As String
    PARENT_NETWORK_ADDRESS As String
    DEVICE_TYPE As String
    STATUS As String
    PROFILe_ID As String
    MANUFACTURER_ID As String
End Type




Public SelectOKButton As Boolean


Dim ShowDeviceList() As ZigBeeInfo
Dim AllDeviceList() As ZigBeeInfo

' define line data info for read line function
Private Type LineDataInfo
    content As String
    timeout As Boolean
End Type

Private Sub EnterCommandMode()
    UpdateProgressBar
    ProXR1.ClearBuffer
    ProXR1.SendData Asc("+")
    ProXR1.SendData Asc("+")
    ProXR1.SendData Asc("+")
    
End Sub

Private Sub ExitCommandMode()
    ProXR1.ClearBuffer
    SendString "ATCN", True
    UpdateProgressBar
    
End Sub

Private Sub ATND_Command()
    ProXR1.ClearBuffer
    SendString "ATND", True
    Erase ShowDeviceList
    ReDim ShowDeviceList(0) As ZigBeeInfo
    Dim info As ZigBeeInfo
        ProXR1.Sleep (3000)
    info = GetZigBeeInfo
    Do Until Not IsValidZigBeeInfo(info)
        dumpZigbeeInfo info
        AddToList info
        ReDim Preserve ShowDeviceList(0 To UBound(ShowDeviceList) + 1) As ZigBeeInfo
        ShowDeviceList(UBound(ShowDeviceList)) = info
        info = GetZigBeeInfo
    Loop
    WriteShowDeviceList
    WriteAllDeviceList
End Sub

' dump the ZigbeeInfo in text format
Private Sub dumpZigbeeInfo(ByRef info As ZigBeeInfo)
    Debug.Print "MY:", info.MY
    Debug.Print "SH:", info.SH
    Debug.Print "SL:", info.SL
    Debug.Print "NI:", info.NI
    Debug.Print "PARENT_NETWORK_ADDRESS:", info.PARENT_NETWORK_ADDRESS
    Debug.Print "DEVICE_TYPE:", info.DEVICE_TYPE
    Debug.Print "STATUS:", info.STATUS
    Debug.Print "PROFILe_ID:", info.PROFILe_ID
    Debug.Print "MANUFACTURER_ID:", info.MANUFACTURER_ID
End Sub

' Get one zigbeeinfo
' the zigbeeinfo will be invalid if failded to read
Private Function GetZigBeeInfo() As ZigBeeInfo
    Dim info As ZigBeeInfo
    GetZigBeeInfo = InitZigBeeInfo
    Dim ln As LineDataInfo
    
    ln = ReadLine(1000)
    GetZigBeeInfo.MY = ln.content
        
    If Not ln.timeout Then
        ln = ReadLine(1000)
        GetZigBeeInfo.SH = ln.content
    End If
    
    If Not ln.timeout Then
        ln = ReadLine(1000)
        GetZigBeeInfo.SL = ln.content
    End If
    
    If Not ln.timeout Then
        ln = ReadLine(1000)
        GetZigBeeInfo.NI = ln.content
    End If
    
    If Not ln.timeout Then
        ln = ReadLine(1000)
        GetZigBeeInfo.PARENT_NETWORK_ADDRESS = ln.content
    End If
    
    If Not ln.timeout Then
        ln = ReadLine(1000)
        GetZigBeeInfo.DEVICE_TYPE = ln.content
    End If
    
    If Not ln.timeout Then
        ln = ReadLine(1000)
        GetZigBeeInfo.STATUS = ln.content
    End If
    
    If Not ln.timeout Then
        ln = ReadLine(1000)
        GetZigBeeInfo.PROFILe_ID = ln.content
    End If
    
    If Not ln.timeout Then
        ln = ReadLine(1000)
        GetZigBeeInfo.MANUFACTURER_ID = ln.content
    End If
    ln = ReadLine(1000)
    dumpZigbeeInfo GetZigBeeInfo
End Function

' Initailize the zigbeeinfo data
' It will be an empty data, and be invalid
Private Function InitZigBeeInfo() As ZigBeeInfo
    Dim info As ZigBeeInfo
    info.MY = ""
    info.SH = ""
    info.SL = ""
    info.NI = ""
    info.PARENT_NETWORK_ADDRESS = ""
    info.DEVICE_TYPE = ""
    info.STATUS = ""
    info.PROFILe_ID = ""
    info.MANUFACTURER_ID = ""
    InitZigBeeInfo = info
End Function

' test if the zigbeeinfo is valid
Private Function IsValidZigBeeInfo(ByRef info As ZigBeeInfo) As Boolean
    IsValidZigBeeInfo = True
    'If info.MY = "" Then IsValidZigBeeInfo = False
    If info.SH = "" Then IsValidZigBeeInfo = False
    If info.SL = "" Then IsValidZigBeeInfo = False
    Debug.Print IsValidZigBeeInfo
End Function

' read a line from serila port
' will read by reading a carriage return character or timeout
' the string return won't contain carriage return
Private Function ReadLine(ByVal timeout As Integer) As LineDataInfo
    ReadLine.timeout = False
    Dim ln As String
    ln = ""
    Dim i As Integer
    i = GetData
    UpdateProgressBar
    Do Until i = -1 Or i = 13
        ln = ln + Chr(i)
        'i = ProXR1.GetData2(timeout)
        i = GetData
    Loop
    If i = -1 Then ReadLine.timeout = True
    ReadLine.content = ln
    Debug.Print ln
End Function


Private Sub GetShowDeviceList()
    ProXR1.Sleep (2000)
    EnterCommandMode
    Dim ln As LineDataInfo
    ln = ReadLine(1000)
    If ln.timeout Then Exit Sub
    If ln.content <> "OK" Then Exit Sub
    ' get devcie list
    ATND_Command
    
    'Debug.Print GetString(1000)
    
    ExitCommandMode
    
End Sub

Private Sub cmdGetList_Click()
    GetShowDeviceList
End Sub

Private Sub cmdAdd_Click()
    Dim info As ZigBeeInfo
    info = InitZigBeeInfo
    info.SH = txtSerialHigh.Text
    info.SL = txtSerialLow.Text
    WriteNameToSetting info, txtName.Text
    AddDevice info
    FillList
    WriteShowDeviceList
    WriteAllDeviceList
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    i = List1.ListIndex
    If i > 0 Then
        Dim info As ZigBeeInfo
        info = ShowDeviceList(i)
        DeleteFromAllDeviceList info
        RemoveItemFromShowDeviceList i
        List1.RemoveItem i
    End If
    WriteShowDeviceList
    WriteAllDeviceList
End Sub

Private Sub DeleteFromAllDeviceList(ByRef info As ZigBeeInfo)
    ' find the postion
    For i = 1 To UBound(AllDeviceList)
        If IsSame(info, AllDeviceList(i)) Then
             RemoveItemFromAllDeviceList i
             Exit For
        End If
    Next
    'ArrayRemoveItem AllDeviceList,
End Sub

Private Sub cmdEditName_Click()
    Dim i As Integer
    i = List1.ListIndex
    If i > 0 Then
        Dim info As ZigBeeInfo
        info = ShowDeviceList(i)
        
        Dim strName As String
        strName = ReadNameFromSetting(info)
        strName = InputBox("Enter Name", "Enter Name", strName)
        WriteNameToSetting info, strName
        Debug.Print strName
        List1.List(i) = GetShowString(info)
    End If

End Sub

Private Sub cmdRefresh_Click()
    List1.Clear
    List1.Refresh
    List1.AddItem BoardCaseItemStr
    ProgressBar1.Visible = True
    ProgressBar1.Value = 0
    ProgressBar1.Refresh
    GetShowDeviceList
    ProgressBar1.Visible = False
End Sub

Private Sub cmdSelect_Click()
    Dim success As Boolean
    success = False
    Dim i As Integer
    i = List1.ListIndex
    If i > 0 Then
        Dim info As ZigBeeInfo
        info = ShowDeviceList(i)
        success = TalkToDevice(info)
    ElseIf i = 0 Then
        ' talk to all device
        success = TalkToAllDevices
    Else
        MsgBox "Please select device"
    End If
    
    If success Then
        Me.Hide
        
        Me.ProXR1.ClosePort
        If IsProXRForm Then
            Set Form1.parentform = Me
            Form1.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
            Form1.ProXR1.BaudRate = ProXR1.BaudRate
            Form1.ProXR1.OpenPort
            Me.Hide
            Form1.Show
            Form1.LoadForm
        Else
            Set frmR2X.parentform = Me
            frmR2X.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
            frmR2X.ProXR1.BaudRate = ProXR1.BaudRate
            frmR2X.ProXR1.OpenPort
            frmR2X.Frame5.Visible = False
            
            frmR2X.Show
        End If
    End If
    
    
End Sub

Private Sub UpdateProgressBar()
    If ProgressBar1.Value < 100 Then
        ProgressBar1.Value = ProgressBar1.Value + 10
    Else
        ProgressBar1.Value = 0
    End If
    ProgressBar1.Refresh
    Debug.Print ProgressBar1.Value
End Sub

Private Sub PutData(ByVal bt As Integer)
    If Not ProXR1.IsPortOpen Then ProXR1.OpenPort
    ProXR1.SendData bt
End Sub

Private Function GetData() As Integer
    If Not ProXR1.IsPortOpen Then ProXR1.OpenPort
    GetData = ProXR1.GetData
End Function


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

' set the target device
Private Function TalkToDevice(ByRef info As ZigBeeInfo) As Boolean
    TalkToDevice = False
    Dim strFailMsg As String
    strFailMsg = "Failed to select device"
    ProXR1.Sleep (2000)
    EnterCommandMode
    Dim ln As LineDataInfo
    ProXR1.Sleep 1000
    
    ln = ReadLine(1000)
    If ln.timeout Then
        MsgBox strFailMsg
        Exit Function
    End If
    If ln.content <> "OK" Then
        MsgBox strFailMsg
        Exit Function
    End If
    SendString "atdh" + info.SH + "wr", True
    ProXR1.Sleep 1000
    ln = ReadLine(1000)
    If ln.content <> "OK" Then
        MsgBox strFailMsg
        Exit Function
    End If
    SendString "atdl" + info.SL + "wr", True
    ProXR1.Sleep 1000
    ln = ReadLine(1000)
    If ln.content <> "OK" Then
        MsgBox strFailMsg
        Exit Function
    End If
    ExitCommandMode
    MsgBox "Device " + info.SH + ":" + info.SL + " is selected"
    TalkToDevice = True
End Function

Private Function TalkToAllDevices() As Boolean
    TalkToAllDevices = False
    Dim strFailMsg As String
    strFailMsg = "Failed to set broadcast mode"
    ProXR1.Sleep (2000)
    EnterCommandMode
    Dim ln As LineDataInfo
    ProXR1.Sleep 1000
    ln = ReadLine(1000)
    If ln.timeout Then
        MsgBox strFailMsg
        Exit Function
    End If
    If ln.content <> "OK" Then
        MsgBox strFailMsg
        Exit Function
    End If
    SendString "atdh0", True
    ProXR1.Sleep 1000
    ln = ReadLine(1000)
    If ln.content <> "OK" Then
        MsgBox strFailMsg
        Exit Function
    End If
    SendString "atdlFFFF", True
    ProXR1.Sleep 1000
    ln = ReadLine(1000)
    If ln.content <> "OK" Then
        MsgBox strFailMsg
        Exit Function
    End If
    ExitCommandMode
    MsgBox "Broadcast to all devices"
    TalkToAllDevices = True
End Function


' write the zig info array to file
Private Sub WriteShowDeviceList()
    Dim file As Long
    file = FreeFile
    Open ShowXBeeListFile For Output As file
    For i = 1 To UBound(ShowDeviceList)
        WriteZigBeeInfoToFile ShowDeviceList(i), file
    Next
    Close file
End Sub

Private Sub WriteZigBeeInfoToFile(ByRef info As ZigBeeInfo, file As Long)
    Print #file, info.MY
    Print #file, info.SH
    Print #file, info.SL
    Print #file, info.NI
    Print #file, info.PARENT_NETWORK_ADDRESS
    Print #file, info.DEVICE_TYPE
    Print #file, info.STATUS
    Print #file, info.PROFILe_ID
    Print #file, info.MANUFACTURER_ID
End Sub

Private Function ReadZigbeeInfoFromFile(file As Long) As ZigBeeInfo
    Dim str As String
    Line Input #file, str
    ReadZigbeeInfoFromFile.MY = str
    
    Line Input #file, str
    ReadZigbeeInfoFromFile.SH = str
    
    Line Input #file, str
    ReadZigbeeInfoFromFile.SL = str
    
    Line Input #file, str
    ReadZigbeeInfoFromFile.NI = str
    
    Line Input #file, str
    ReadZigbeeInfoFromFile.PARENT_NETWORK_ADDRESS = str
    
    Line Input #file, str
    ReadZigbeeInfoFromFile.DEVICE_TYPE = str
    
    Line Input #file, str
    ReadZigbeeInfoFromFile.STATUS = str
    
    Line Input #file, str
    ReadZigbeeInfoFromFile.PROFILe_ID = str
    
    Line Input #file, str
    ReadZigbeeInfoFromFile.MANUFACTURER_ID = str
End Function
' write zig bee info to file
Private Sub ReadShowDeviceList()
    Dim file As Long
    file = FreeFile
    Erase ShowDeviceList
    ReDim ShowDeviceList(0) As ZigBeeInfo
On Error GoTo ErrorHandler
    Open ShowXBeeListFile For Input As file
    Dim info As ZigBeeInfo
    Do While Not EOF(file)
        info = ReadZigbeeInfoFromFile(file)
        ReDim Preserve ShowDeviceList(0 To UBound(ShowDeviceList) + 1) As ZigBeeInfo
        ShowDeviceList(UBound(ShowDeviceList)) = info
        dumpZigbeeInfo info
    Loop
    Close file
ErrorHandler:

End Sub

' read all devices information from list
Private Sub ReadAllDeviceList()
    Erase AllDeviceList
    ReDim AllDeviceList(0) As ZigBeeInfo
    Dim file As Long
    file = FreeFile
On Error GoTo ErrorHandler
    Open AllXBeeListFile For Input As file
    Dim info As ZigBeeInfo
    Do While Not EOF(file)
        info = ReadZigbeeInfoFromFile(file)
        ReDim Preserve AllDeviceList(0 To UBound(AllDeviceList) + 1) As ZigBeeInfo
        AllDeviceList(UBound(AllDeviceList)) = info
        dumpZigbeeInfo info
    Loop
    Close file
ErrorHandler:

End Sub

' write all devices information to list
Private Sub WriteAllDeviceList()
    Dim file As Long
    file = FreeFile
    Open AllXBeeListFile For Output As file
    For i = 1 To UBound(ShowDeviceList)
        WriteZigBeeInfoToFile ShowDeviceList(i), file
    Next
    Close file
End Sub

Private Sub AddToList(ByRef info As ZigBeeInfo)
    List1.AddItem GetShowString(info)
End Sub

Private Function GetShowString(ByRef info As ZigBeeInfo) As String
    Dim strName As String
    strName = ReadNameFromSetting(info)
    Dim strType As String
    'strType = "unknown    "
    strType = "           "
    If info.DEVICE_TYPE = "01" Then strType = "Router     "
    If info.DEVICE_TYPE = "00" Then strType = "Coord      "
    If info.DEVICE_TYPE = "02" Then strType = "End Device "
    GetShowString = strType + " " + info.SH + ":" + info.SL + "  " + strName
End Function

Private Function ReadNameFromSetting(ByRef info As ZigBeeInfo) As String
    ReadNameFromSetting = GetSetting("ProXR_V4", KeyValue, info.SH + info.SL, "")
End Function

Private Sub WriteNameToSetting(ByRef info As ZigBeeInfo, strName As String)
    SaveSetting "ProXR_V4", KeyValue, info.SH + info.SL, strName
End Sub

Private Sub cmdShowAll_Click()
    cmdRefresh_Click
    ReDim ShowDeviceList(0 To UBound(AllDeviceList)) As ZigBeeInfo
    For i = 1 To UBound(AllDeviceList)
        ShowDeviceList(i) = AllDeviceList(i)
    Next
    FillList
End Sub

Private Sub Form_Load()
    SelectOKButton = False
    ReadAllDeviceList
    ReadShowDeviceList
    FillList
End Sub

Private Sub AddDevice(ByRef info As ZigBeeInfo)
    Dim isExist  As Boolean
    isExist = False
    ' check if it is already in showdevicelist
    For i = 1 To UBound(ShowDeviceList)
        If IsSame(info, ShowDeviceList(i)) Then
            ShowDeviceList(i) = info    ' replace in list
            isExist = True
        End If
    Next
    If Not isExist Then
        ' add to list
        ReDim Preserve ShowDeviceList(0 To UBound(ShowDeviceList) + 1) As ZigBeeInfo
        ShowDeviceList(UBound(ShowDeviceList)) = info
    End If
    ' merge to all device list
    AddDeviceToAllDeviceList info
End Sub

Private Sub AddDeviceToAllDeviceList(ByRef info As ZigBeeInfo)
    Dim isExist  As Boolean
    isExist = False
    ' check if it is already in showdevicelist
    For i = 1 To UBound(AllDeviceList)
        If IsSame(info, AllDeviceList(i)) Then
            ShowDeviceList(i) = info    ' replace in list
            isExist = True
        End If
    Next
    If Not isExist Then
        ' add to list
        ReDim Preserve AllDeviceList(0 To UBound(AllDeviceList) + 1) As ZigBeeInfo
        AllDeviceList(UBound(AllDeviceList)) = info
    End If
End Sub

' test if the zigbee information are same. Only test SH and SL
Private Function IsSame(ByRef info1 As ZigBeeInfo, ByRef info2 As ZigBeeInfo) As Boolean
    IsSame = False
    If info1.SH = info2.SH And info1.SL = info2.SL Then
        IsSame = True
    End If
End Function

' Fill the device list in dialog with items in ShowDeviceList
Private Sub FillList()
    ' Add broadcast list
    List1.Clear
    List1.AddItem BoardCaseItemStr
    For i = 1 To UBound(ShowDeviceList)
        AddToList ShowDeviceList(i)
    Next
End Sub



Public Sub RemoveItemFromShowDeviceList(ByVal ItemElement As Long)

'PURPOSE:       Remove an item from an array, then
'               resize the array

'PARAMETERS:    ItemArray: Array, passed by reference, with
'               item to be removed.  Array must not be fixed

'               ItemElement: Element to Remove
                
'EXAMPLE:
'           dim iCtr as integer
'           Dim sTest() As String
'           ReDim sTest(2) As String
'           sTest(0) = "Hello"
'           sTest(1) = "World"
'           sTest(2) = "!"
'           ArrayRemoveItem sTest, 1
'           for iCtr = 0 to ubound(sTest)
'               Debug.print sTest(ictr)
'           next
'
'           Prints
'
'           "Hello"
'           "!"
'           To the Debug Window

Dim lCtr As Long
Dim lTop As Long
Dim lBottom As Long



lTop = UBound(ShowDeviceList)
lBottom = LBound(ShowDeviceList)

If ItemElement < lBottom Or ItemElement > lTop Then
    Err.Raise 9, , "Subscript out of Range"
    Exit Sub
End If

For lCtr = ItemElement To lTop - 1
    ShowDeviceList(lCtr) = ShowDeviceList(lCtr + 1)
Next
On Error GoTo ErrorHandler:

ReDim Preserve ShowDeviceList(lBottom To lTop - 1)

Exit Sub
ErrorHandler:
  'An error will occur if array is fixed
    Err.Raise Err.Number, , _
       "You must pass a resizable array to this function"
End Sub

Public Sub RemoveItemFromAllDeviceList(ByVal ItemElement As Long)

'PURPOSE:       Remove an item from an array, then
'               resize the array

'PARAMETERS:    ItemArray: Array, passed by reference, with
'               item to be removed.  Array must not be fixed

'               ItemElement: Element to Remove
                
'EXAMPLE:
'           dim iCtr as integer
'           Dim sTest() As String
'           ReDim sTest(2) As String
'           sTest(0) = "Hello"
'           sTest(1) = "World"
'           sTest(2) = "!"
'           ArrayRemoveItem sTest, 1
'           for iCtr = 0 to ubound(sTest)
'               Debug.print sTest(ictr)
'           next
'
'           Prints
'
'           "Hello"
'           "!"
'           To the Debug Window

Dim lCtr As Long
Dim lTop As Long
Dim lBottom As Long



lTop = UBound(AllDeviceList)
lBottom = LBound(AllDeviceList)

If ItemElement < lBottom Or ItemElement > lTop Then
    Err.Raise 9, , "Subscript out of Range"
    Exit Sub
End If

For lCtr = ItemElement To lTop - 1
    AllDeviceList(lCtr) = AllDeviceList(lCtr + 1)
Next
On Error GoTo ErrorHandler:

ReDim Preserve AllDeviceList(lBottom To lTop - 1)

Exit Sub
ErrorHandler:
  'An error will occur if array is fixed
    Err.Raise Err.Number, , _
       "You must pass a resizable array to this function"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProXR1.ClosePort
    frmSelectDevice.Show
End Sub
