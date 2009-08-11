VERSION 5.00
Object = "{02E0654E-AAC5-4BBF-A1DE-45576B24DFC1}#2.1#0"; "ProXR.ocx"
Begin VB.Form frmPwm8x 
   Caption         =   "Dimmer/Speed Controller 9600 COM1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "E3C 252"
      Height          =   495
      Left            =   3720
      TabIndex        =   26
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get E3C Device Number"
      Height          =   495
      Left            =   2040
      TabIndex        =   25
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Brightness Level"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin NCDProXR.ProXR ProXR1 
         Left            =   3840
         Top             =   4200
         _ExtentX        =   926
         _ExtentY        =   926
         BaudRate        =   9600
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select an E3C Device Number"
         Height          =   1215
         Left            =   120
         TabIndex        =   21
         Top             =   4440
         Width           =   4335
         Begin VB.CommandButton Command2 
            Caption         =   "Store E3C Device Number"
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   22
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label2 
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
            Left            =   3600
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   1
         TabIndex        =   20
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Store as Default Powerup"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   3600
         Width           =   4095
      End
      Begin VB.HScrollBar Bright 
         Height          =   255
         Index           =   0
         Left            =   480
         Max             =   255
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
      Begin VB.HScrollBar Bright 
         Height          =   255
         Index           =   8
         Left            =   480
         Max             =   255
         TabIndex        =   8
         Top             =   3240
         Width           =   3135
      End
      Begin VB.HScrollBar Bright 
         Height          =   255
         Index           =   7
         Left            =   480
         Max             =   255
         TabIndex        =   7
         Top             =   2880
         Width           =   3135
      End
      Begin VB.HScrollBar Bright 
         Height          =   255
         Index           =   6
         Left            =   480
         Max             =   255
         TabIndex        =   6
         Top             =   2520
         Width           =   3135
      End
      Begin VB.HScrollBar Bright 
         Height          =   255
         Index           =   5
         Left            =   480
         Max             =   255
         TabIndex        =   5
         Top             =   2160
         Width           =   3135
      End
      Begin VB.HScrollBar Bright 
         Height          =   255
         Index           =   4
         Left            =   480
         Max             =   255
         TabIndex        =   4
         Top             =   1800
         Width           =   3135
      End
      Begin VB.HScrollBar Bright 
         Height          =   255
         Index           =   3
         Left            =   480
         Max             =   255
         TabIndex        =   3
         Top             =   1440
         Width           =   3135
      End
      Begin VB.HScrollBar Bright 
         Height          =   255
         Index           =   2
         Left            =   480
         Max             =   255
         TabIndex        =   2
         Top             =   1080
         Width           =   3135
      End
      Begin VB.HScrollBar Bright 
         Height          =   255
         Index           =   1
         Left            =   480
         Max             =   255
         TabIndex        =   1
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "8"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "7"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "5"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "4"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "ALL"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Slow On Mode"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label1 
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
         Index           =   8
         Left            =   3720
         TabIndex        =   18
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   3720
         TabIndex        =   17
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   6
         Left            =   3720
         TabIndex        =   16
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   3720
         TabIndex        =   15
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   4
         Left            =   3720
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   3720
         TabIndex        =   13
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   3720
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   3720
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   3720
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPwm8x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public parentForm As Form

Private Sub Bright_Change(Index As Integer)
    PutData 254              'Enter Command Mode
    PutData Index            'Send Command
    PutData (Bright(Index))   'Set Brightness
    Label1(Index).Caption = Bright(Index)
End Sub
Private Sub Bright_Scroll(Index As Integer)
    PutData (254)             'Enter Command Mode
    PutData (Index)           'Send Command
    PutData (Bright(Index))   'Set Brightness
    Label1(Index).Caption = Bright(Index)
End Sub
Private Sub Command1_Click()
    PutData (254)             'Enter Command Mode
    PutData (9)   'Send PWM Store Command
    PutData (0)   'Any Value Goes Here, Not Used
End Sub
Private Sub Command2_Click()
    PutData (254)             'Enter Command Mode
    PutData (255)             'Send PWM Store Command
    PutData (HScroll2.Value)   'Any Value Goes Here, Not Used
End Sub
Private Sub Command3_Click()
    junk = GetData
    Debug.Print "Junk"
    PutData (254)             'Enter Command Mode
    PutData (247)             'Send PWM Store Command
    HScroll2.Value = GetData
End Sub
Private Sub Command4_Click()
    PutData (254)             'Enter Command Mode
    PutData (252)   'Send PWM Store Command
    PutData (HScroll2.Value)   '0=Fast On, 1=Slow Soft On
End Sub
Private Sub Form_Load()
    ProXR1.OpenPort
    ProXR1.ClearBuffer
    frmLog.Show
    Set frmLog.proxrObj = ProXR1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProXR1.ClosePort
    Unload frmLog
    ProXR1.ClosePort
    If Not parentForm Is Nothing Then
        parentForm.Show
    End If
End Sub

Private Sub HScroll1_Change()
    PutData (254)             'Enter Command Mode
    PutData (10)   'Send PWM Store Command
    PutData (HScroll1.Value)   '0=Fast On, 1=Slow Soft On
    If HScroll1.Value = 1 Then
        Label3.Caption = "Fast On Mode"
    Else
        Label3.Caption = "Slow On Mode"
    End If
End Sub
Private Sub HScroll2_Change()
    Label2.Caption = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
    Label2.Caption = HScroll2.Value
End Sub


Private Sub PutData(ByVal data As Integer)
    ProXR1.SendData data
End Sub

Private Function GetData() As Integer
    GetData = ProXR1.GetData
    If GetData < 0 Then GetData = 0
End Function


Private Sub ProXR1_OnDataReceived(ByVal data As Integer)
    'data received
    frmLog.OnDataReceived data
End Sub

Private Sub ProXR1_OnDataSent(ByVal data As Integer)
    'data sent
    frmLog.OnDataSent data
End Sub
