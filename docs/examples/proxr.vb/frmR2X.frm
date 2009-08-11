VERSION 5.00
Object = "{02E0654E-AAC5-4BBF-A1DE-45576B24DFC1}#2.1#0"; "ProXR.ocx"
Begin VB.Form frmR2X 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin NCDProXR.ProXR ProXR1 
      Left            =   5640
      Top             =   120
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin VB.Frame Frame5 
      Caption         =   "E3C Networking: Control 256 NCD Devices from a Single Serial Port"
      Height          =   2175
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   6135
      Begin VB.Frame Frame6 
         Caption         =   "Select a Device to Control"
         Height          =   975
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   5895
         Begin VB.HScrollBar HScroll3 
            Height          =   495
            Left            =   120
            Max             =   255
            TabIndex        =   24
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "ALL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5040
            TabIndex        =   25
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   615
         Left            =   240
         Max             =   255
         TabIndex        =   22
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Program Device Number"
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Retrieve Device Number"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Default Power-Up Status of Relays"
      Height          =   1455
      Left            =   3240
      TabIndex        =   13
      Top             =   1440
      Width           =   3015
      Begin VB.CommandButton Command9 
         Caption         =   "Get Power-Up Default Status"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Set Current Relay Status as Power-Up Defaul Status"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   1680
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
      Begin VB.CommandButton Command6 
         Caption         =   "Get Relay 2 Status"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
      Begin VB.CommandButton Command5 
         Caption         =   "Get Relay 1 Status"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control Both Relays at Once"
      Height          =   1335
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command7 
         Caption         =   "Read Status of Both Relays"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2775
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   375
         Left            =   120
         Max             =   3
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Relay 2 On"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Relay 2 Off"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Relay 1 On"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Relay 1 Off"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmR2X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public parentform As Form

Private Sub Form_Load()
    ProXR1.OpenPort
    Caption = "R25/R210 Relay Test   " + ProXR1.PortName + "    www.controlanything.com"
    ProXR1.ClearBuffer
    frmLog.Show
    Set frmLog.proxrObj = ProXR1
End Sub
Private Sub Command1_Click()
    PutData 254  'Turn Off Relay 1
    PutData 0
End Sub
Private Sub Command2_Click()
    PutData 254  'Turn On Relay 1
    PutData 1
End Sub
Private Sub Command3_Click()
    PutData 254  'Turn Off Relay 2
    PutData 2
End Sub
Private Sub Command4_Click()
    PutData 254  'Turn On Relay 2
    PutData 3
End Sub
Private Sub Command5_Click()
    PutData 254  'Get Status of Relay 1
    PutData 4
    Label1.Caption = GetData
End Sub
Private Sub Command6_Click()
    PutData 254  'Get Status of Relay 2
    PutData 5
    Label2.Caption = GetData
End Sub
Private Sub Command7_Click()
    PutData 254  'Get Status of Both Relays
    PutData 7
    Label3.Caption = GetData
End Sub
Private Sub Command8_Click()
    PutData 254  'Store Relay Status as Powerup Default
    PutData 8
End Sub
Private Sub Command9_Click()
    PutData 254  'Get Powerup Default Relay Status
    PutData 9
    Label4.Caption = GetData
End Sub
Private Sub Command10_Click()
    PutData 254  'Program E3C Device Number
    PutData 255
    PutData HScroll2.Value
End Sub
Private Sub Command11_Click()
    PutData 254  'Retrieve Stored E3C Device Number
    PutData 247
    Label5.Caption = GetData
    HScroll2.Value = Label5.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ProXR1.ClosePort
    Unload frmLog
    ProXR1.ClosePort
    If Not parentform Is Nothing Then
        parentform.Show
    End If
End Sub

Private Sub HScroll1_Change()
    PutData 254  'Set Status of Both Relays at Once
    PutData 6
    PutData HScroll1.Value
    Label3.Caption = HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
    PutData 254  'Set Status of Both Relays at Once
    PutData 6
    PutData HScroll1.Value
    Label3.Caption = HScroll1.Value
End Sub
Private Sub HScroll2_Change()
    Label5.Caption = HScroll2.Value
End Sub
Private Sub HScroll2_Scroll()
    Label5.Caption = HScroll2.Value
End Sub
Private Sub HScroll3_Change()
    PutData 254  'E3C: Select a Device to Control
    PutData 252
    PutData HScroll3.Value
    Label6.Caption = HScroll3.Value
End Sub
Private Sub HScroll3_Scroll()
    PutData 254  'E3C: Select a Device to Control
    PutData 252
    PutData HScroll3.Value
    Label6.Caption = HScroll3.Value
End Sub

' send data to serial port
Private Sub PutData(ByVal data As Integer)
    ProXR1.SendData data
End Sub

' get data from serial port
Private Function GetData() As Integer
    GetData = ProXR1.GetData
End Function


Private Sub ProXR1_OnDataReceived(ByVal data As Integer)
    'data received
    frmLog.OnDataReceived data
End Sub

Private Sub ProXR1_OnDataSent(ByVal data As Integer)
    'data sent
    frmLog.OnDataSent data
End Sub
