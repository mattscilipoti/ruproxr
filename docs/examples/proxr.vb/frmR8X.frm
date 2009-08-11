VERSION 5.00
Object = "{02E0654E-AAC5-4BBF-A1DE-45576B24DFC1}#2.1#0"; "ProXR.ocx"
Begin VB.Form frmR8X 
   Caption         =   "R4x/R8x Pro Example Software for Visual Basic 6"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin NCDProXR.ProXR ProXR1 
      Left            =   6240
      Top             =   120
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin VB.Frame Frame21 
      Caption         =   "Relay ON/OFF Status on Powerup"
      Height          =   1215
      Left            =   3120
      TabIndex        =   69
      Top             =   8160
      Width           =   4095
      Begin VB.CommandButton GetPattern 
         Caption         =   "Get Powerup Default Relay Pattern"
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton StorePattern 
         Caption         =   "Store Current Pattern as Powerup Default"
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label DefPat 
         Alignment       =   2  'Center
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   72
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton StoreBank 
      Caption         =   "Store Current Relay Battern in Above Memory Bank"
      Height          =   255
      Left            =   3240
      TabIndex        =   68
      Top             =   7440
      Width           =   3855
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   3360
      Max             =   15
      TabIndex        =   67
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Frame Frame19 
      Caption         =   "Store and Recall Relay Pattern in Memory Bank"
      Height          =   1695
      Left            =   3120
      TabIndex        =   64
      Top             =   6480
      Width           =   4095
      Begin VB.CommandButton GetBank 
         Caption         =   "Get Current Relay Battern from Above Memory Bank"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Frame Frame20 
         Caption         =   "Memory Bank"
         Height          =   615
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   3855
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   66
            Top             =   240
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "Extended Relay Control Commands"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   48
      Top             =   4560
      Width           =   7215
      Begin VB.Frame Frame18 
         Caption         =   "Program Emulation Device Number"
         Height          =   1095
         Left            =   120
         TabIndex        =   60
         Top             =   1920
         Width           =   2775
         Begin VB.CommandButton ProgramEmu 
            Caption         =   "Program"
            Height          =   255
            Left            =   1440
            TabIndex        =   63
            Top             =   720
            Width           =   1215
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   120
            Max             =   15
            TabIndex        =   61
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Set Status of All Relays at Once"
         Height          =   975
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   2535
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   58
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.CommandButton Test2Way 
         Caption         =   "Test 2-Way Communication"
         Height          =   495
         Left            =   5880
         TabIndex        =   53
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Reverse 
         Caption         =   "Reverse On/Off Relay Pattern"
         Height          =   495
         Left            =   4200
         TabIndex        =   52
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Invert 
         Caption         =   "Invert Status of All Relays"
         Height          =   495
         Left            =   2760
         TabIndex        =   51
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton AllOn 
         Caption         =   "All Relays On"
         Height          =   495
         Left            =   1440
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton AllOff 
         Caption         =   "All Relays Off"
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label T2Way 
         Alignment       =   2  'Center
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   56
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Relay On/Off Pattern is Reversed, Relay 12345678 status is coppied to Relay 87654321"
         Height          =   975
         Left            =   4200
         TabIndex        =   55
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Relays that are on turn off.  Relays that are off turn on."
         Height          =   855
         Left            =   2760
         TabIndex        =   54
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Reporting 
      Caption         =   "OFF"
      Height          =   495
      Left            =   6360
      TabIndex        =   47
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Read_All 
      Caption         =   "Read Status of All Relays"
      Height          =   255
      Left            =   5160
      TabIndex        =   44
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Read_Relay 
      Caption         =   "Read"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   27
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "R8x Pro ONLY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   3615
      Begin VB.Frame Frame10 
         Caption         =   "Relay 8"
         Height          =   855
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Relay 5"
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Relay 6"
         Height          =   855
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Relay 7"
         Height          =   855
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Relay 1"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   855
      Begin VB.CommandButton RELAY 
         Caption         =   "OFF"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Turn Individual Relay On/Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame6 
         Caption         =   "Relay 4"
         Height          =   855
         Left            =   2640
         TabIndex        =   5
         Top             =   480
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Relay 3"
         Height          =   855
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Relay 2"
         Height          =   855
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Read Status of Relays"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   7215
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   7
         Left            =   6120
         TabIndex        =   43
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   6
         Left            =   5280
         TabIndex        =   42
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   41
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   40
         Top             =   1200
         Width           =   855
      End
      Begin VB.Frame Frame14 
         Caption         =   "R8x Pro ONLY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3480
         TabIndex        =   31
         Top             =   240
         Width           =   3615
         Begin VB.Frame Frame12 
            Caption         =   "Relay 5"
            Height          =   855
            Index           =   3
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   855
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Caption         =   "???"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   39
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Relay 6"
            Height          =   855
            Index           =   4
            Left            =   960
            TabIndex        =   36
            Top             =   240
            Width           =   855
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Caption         =   "???"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   37
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Relay 7"
            Height          =   855
            Index           =   5
            Left            =   1800
            TabIndex        =   34
            Top             =   240
            Width           =   855
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Caption         =   "???"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   35
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Relay 8"
            Height          =   855
            Index           =   6
            Left            =   2640
            TabIndex        =   32
            Top             =   240
            Width           =   855
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Caption         =   "???"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   33
               Top             =   360
               Width           =   615
            End
         End
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   30
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   29
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   28
         Top             =   1200
         Width           =   855
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 4"
         Height          =   855
         Index           =   2
         Left            =   2640
         TabIndex        =   25
         Top             =   480
         Width           =   855
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "???"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 3"
         Height          =   855
         Index           =   1
         Left            =   1800
         TabIndex        =   23
         Top             =   480
         Width           =   855
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "???"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Relay 2"
         Height          =   855
         Left            =   960
         TabIndex        =   21
         Top             =   480
         Width           =   855
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "???"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 1"
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   855
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "???"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "Reporting Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   45
      Top             =   3480
      Width           =   7215
      Begin VB.Label Label2 
         Caption         =   $"frmR8X.frx":0000
         Height          =   615
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmR8X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public parentform As Form

Private Sub AllOff_Click()
    PutData 254             'Enter Command Mode
    PutData 29              'Send Command Turn All Relays Off
End Sub
Private Sub AllOn_Click()
    PutData 254             'Enter Command Mode
    PutData 30             'Send Command Turn All Relays On
End Sub
Private Sub Form_Load()
    Set frmLog.proxrObj = ProXR1
    frmLog.Show
'C44: RELAY = 0: GoSub All8                            'Relay Select:
'    Temp = (Param & 7)                                'Safe Make Before Break: 1 of 8 Activates One Relay at a Time Only, All Others OFF
'    Branch Temp, [C8,C9,C10,C11,C12,C13,C14,C15]      'Parameter Selects the Relay to Activate
'C45:    RELAY = 255: GoSub All8                       'Relay DeSelect:
'    Temp = (Param & 7)                                'Safe Break Before Make:    1 of 8 Deactivates One Relay at a Time Only, All Others ON
'    Branch Temp, [C0,C1,C2,C3,C4,C5,C6,C7]            'Parameter Selects the Relay to Deactivate
'C46:    Temp = (Param & 7)                            'Relay Invert: Whatever Status Is, Reverse It
'    RELAY = RELAY ^ Temp
'    GoTo Update
'    ProXR1.OpenPort
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Unload frmLog
    ProXR1.ClosePort
    If Not parentform Is Nothing Then
        parentform.Show
    End If
End Sub

Private Sub GetBank_Click()
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 43               'Get Status of All Relays in a Memory Bank
    PutData HScroll3.Value   'Select a Memory Bank to Recall Pattern (0-15)
End Sub
Private Sub GetPattern_Click()
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 26               'Send Command to Read the Status of All Relays on Powerup
    DefPat.Caption = GetData                'Read Status from Relay Board
End Sub
Private Sub HScroll1_Change()
    Label6.Caption = HScroll1.Value
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 40               'Set Status of All Relays Command
    PutData HScroll1.Value   'Set Status of All Relays Command
End Sub
Private Sub HScroll1_Scroll()
    Label6.Caption = HScroll1.Value
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 40               'Set Status of All Relays Command
    PutData HScroll1.Value   'Set Status of All Relays Command
End Sub
Private Sub HScroll2_Change()
    Label7.Caption = HScroll2.Value
End Sub
Private Sub HScroll2_Scroll()
    Label7.Caption = HScroll2.Value
End Sub
Private Sub HScroll3_Change()
    Label8.Caption = HScroll3.Value
End Sub
Private Sub HScroll3_Scroll()
    Label8.Caption = HScroll3.Value
End Sub
Private Sub Invert_Click()
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 31               'Send Command Invert Relay Status
End Sub
Private Sub ProgramEmu_Click()
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 41               'Program Emulation Device Number
    PutData HScroll2.Value   'Device Number 0-15
End Sub

Private Sub ProXR1_OnDataReceived(ByVal data As Integer)
    frmLog.OnDataReceived data
End Sub

Private Sub ProXR1_OnDataSent(ByVal data As Integer)
    frmLog.OnDataSent data
End Sub

Private Sub Read_All_Click()
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 24               'Send Command to Read the Status of All Relays
    Dim temp As Integer
    temp = GetData                          'Read Status from Relay Board
        
    'Set All to OFF
    For N = 0 To 7
        Label1(Index).Caption = "OFF"
    Next N
    
    'Determine Which Relays are ON
    If (temp And 1) = 1 Then Label1(0).Caption = "ON"
    If (temp And 2) = 2 Then Label1(1).Caption = "ON"
    If (temp And 4) = 4 Then Label1(2).Caption = "ON"
    If (temp And 8) = 8 Then Label1(3).Caption = "ON"
    If (temp And 16) = 16 Then Label1(4).Caption = "ON"
    If (temp And 32) = 32 Then Label1(5).Caption = "ON"
    If (temp And 64) = 64 Then Label1(6).Caption = "ON"
    If (temp And 128) = 128 Then Label1(7).Caption = "ON"
    
End Sub
Private Sub Read_Relay_Click(Index As Integer)
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 16 + Index       'Send Command to Read the Status of a Relay (1-8)
    temp = GetData                          'Read Status from Relay Board
    If temp = 0 Then
        Label1(Index).Caption = "OFF"
    Else
        Label1(Index).Caption = "ON"
    End If
End Sub
Private Sub RELAY_Click(Index As Integer)
    If RELAY(Index).Caption = "ON" Then
        RELAY(Index).Caption = "OFF"
        ProXR1.ClearBuffer
        PutData 254          'Enter Command Mode
        PutData Index        'Turn Relay Off
    Else
        RELAY(Index).Caption = "ON"
        ProXR1.ClearBuffer
        PutData 254          'Enter Command Mode
        PutData Index + 8    'Turn Relay Off
    End If
End Sub
Private Sub Reporting_Click()
    If Reporting.Caption = "OFF" Then
        Reporting.Caption = "ON"
        ProXR1.ClearBuffer
        PutData 254          'Enter Command Mode
        PutData 27           'Turn Reporting Mode ON
    Else
        Reporting.Caption = "OFF"
        ProXR1.ClearBuffer
        PutData 254          'Enter Command Mode
        PutData 28           'Turn Reporting Mode OFF
    End If
End Sub
Private Sub Reverse_Click()
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 32               'Send Command to Reverse Relay Pattern
End Sub
Private Sub StoreBank_Click()
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 42               'Set Status of All Relays in a Memory Bank
    PutData HScroll3.Value   'Select a Memory Bank to Store Pattern (0-15)
End Sub
Private Sub StorePattern_Click()
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 25               'Store Relay Pattern as Powerup Default
End Sub
Private Sub Test2Way_Click()
    ProXR1.ClearBuffer
    PutData 254              'Enter Command Mode
    PutData 33               'Send Command to Test 2-Way Communication
    If GetData = 85 Then                    'Read Status from Relay Board
        T2Way.Caption = "PASS"              '2-Way Communication Test Passed
    Else
        T2Way.Caption = "FAIL"
    End If
End Sub

' send data to serial port
Private Sub PutData(ByVal data As Integer)
    ProXR1.SendData data
End Sub

' get data from serial port
Private Function GetData() As Integer
    GetData = ProXR1.GetData
End Function
