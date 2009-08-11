VERSION 5.00
Begin VB.Form AdvancedFeatures 
   Caption         =   "Advanced Features Configuration                                       WWW.CONTROLANYTHING.COM"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form3"
   ScaleHeight     =   9615
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Relay Repetitions"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton WriteReps 
         Caption         =   "Write/Store Reps Value to Controller"
         Height          =   495
         Left            =   7440
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton ReadReps 
         Caption         =   "Read Reps Value from Controller"
         Height          =   495
         Left            =   5760
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.HScrollBar RepsSlider 
         Height          =   255
         Left            =   120
         Max             =   255
         Min             =   1
         TabIndex        =   2
         Top             =   360
         Value           =   1
         Width           =   4695
      End
      Begin VB.Label Label6 
         Caption         =   $"AdvancedFeatures.frx":0000
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   9495
      End
      Begin VB.Label Label10 
         Caption         =   $"AdvancedFeatures.frx":00F0
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   7440
         TabIndex        =   6
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label RepsTag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   $"AdvancedFeatures.frx":01BD
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   7215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Serial Timing"
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   9735
      Begin VB.HScrollBar TimingSlider 
         Height          =   255
         Left            =   120
         Max             =   255
         Min             =   3
         TabIndex        =   19
         Top             =   240
         Value           =   5
         Width           =   4215
      End
      Begin VB.CommandButton WriteTiming 
         Caption         =   "Store Serial Timing Value Into the Controller"
         Height          =   495
         Left            =   7320
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton ReadTiming 
         Caption         =   "Read Serial Timing Value from Controller"
         Height          =   495
         Left            =   5280
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Clicking This Button Will Set the Communication Baud Rate of the Program to 38.4K Baud"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   7320
         TabIndex        =   30
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "This Setting Can ONLY Be Changed while in Configuration Mode."
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   7320
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   $"AdvancedFeatures.frx":0525
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   6975
      End
      Begin VB.Label TimingTag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Attached Banks"
      Height          =   2055
      Left            =   120
      TabIndex        =   13
      Top             =   5280
      Width           =   9735
      Begin VB.HScrollBar BanksSlider 
         Height          =   255
         Left            =   120
         Max             =   32
         Min             =   1
         TabIndex        =   20
         Top             =   240
         Value           =   1
         Width           =   4095
      End
      Begin VB.CommandButton ReadBanks 
         Caption         =   "Read Attached Banks Value from Controller"
         Height          =   495
         Left            =   5160
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton WriteBanks 
         Caption         =   "Store Attached Banks Value Into the Controller"
         Height          =   495
         Left            =   7320
         TabIndex        =   14
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Clicking This Button Will Set the Communication Baud Rate of the Program to 38.4K Baud"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   7320
         TabIndex        =   29
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label BanksTag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   $"AdvancedFeatures.frx":0707
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   7095
      End
      Begin VB.Label Label4 
         Caption         =   "This Setting Can ONLY Be Changed while in Configuration Mode."
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   7320
         TabIndex        =   16
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Test Cycle Mode"
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   7320
      Width           =   9735
      Begin VB.CommandButton StoreTCycle 
         Caption         =   "Store Test Cycle Data (Works in Any Mode)"
         Height          =   375
         Left            =   4440
         TabIndex        =   31
         Top             =   240
         Width           =   5175
      End
      Begin VB.HScrollBar TCycle 
         Height          =   255
         Left            =   120
         Max             =   32
         TabIndex        =   26
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   $"AdvancedFeatures.frx":08E9
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   9495
      End
      Begin VB.Label TCycleTag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Defaults 
      Caption         =   "Configure Controller to Factory Default Settings - Configuration Mode ONLY"
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   9240
      Width           =   6135
   End
   Begin VB.CommandButton ChangeBaud 
      Caption         =   "Change Baud Rate to 38.4K Baud"
      Height          =   495
      Left            =   8160
      TabIndex        =   23
      Top             =   8640
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   $"AdvancedFeatures.frx":09F6
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   8640
      Width           =   7935
   End
End
Attribute VB_Name = "AdvancedFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChangeBaud_Click()
    Form1.BaudSet(5).Value = True
End Sub
Private Sub Defaults_Click()
    If Check_Store = 1 Then
        Form1.PutData (254)    'Enter Command Mode
        Form1.PutData (50)     'Timer/Setup Branch Commands
        Form1.PutData (144)    'Restore to Factory Default Settings
        Temp = Form1.GetData
    End If
End Sub
Private Sub Form_Load()
    
On Error GoTo EXT
    Form1.GetData
    
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)     'Timer/Setup Branch Commands
    Form1.PutData (136)    'Read Reps Command
    Temp = Form1.GetData
    If Temp <> "" Then
        RepsSlider.Value = Temp 'Form1.GetData    'Get Value from Controller
 '       RepsSlider.Value = Form1.GetData
    End If
    
    If Frame4.Visible = True Then
        Form1.PutData (254)    'Enter Command Mode
        Form1.PutData (50)     'Timer/Setup Branch Commands
        Form1.PutData (138)    'Read Serial Timing Command
        TimingSlider.Value = Form1.GetData  'Get Value from Controller
    End If
    
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)     'Timer/Setup Branch Commands
    Form1.PutData (140)    'Read Attached Banks Command
    BanksSlider.Value = Form1.GetData   'Get Value from Controller
    
    If Frame4.Visible = True Then
        Form1.PutData (254)    'Enter Command Mode
        Form1.PutData (50)     'Timer/Setup Branch Commands
        Form1.PutData (145)    'Get Test Cycle Data Command
        Temp = Form1.GetData
        If Temp <= TCycle.Max Then
            TCycle.Value = Temp            'Get Value from Controller
        End If
    Else
        Debug.Print "Skipping Test Cycle Configuration, this is an advanced controller"
    End If
        Exit Sub
EXT:
    MsgBox ("Communications error, exiting program.")
    'End
End Sub

Private Sub ReadBanks_Click()
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)     'Timer/Setup Branch Commands
    Form1.PutData (140)    'Read Attached Banks Command
    BanksSlider.Value = Form1.GetData   'Get Value from Controller
End Sub
Private Sub ReadReps_Click()
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)     'Timer/Setup Branch Commands
    Form1.PutData (136)    'Read Reps Command
    Temp = Form1.GetData
    'Debug.Print Temp
    RepsSlider.Value = Temp    'Get Value from Controller
End Sub
Private Sub ReadTiming_Click()
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)     'Timer/Setup Branch Commands
    Form1.PutData (138)    'Read Serial Timing Command
    Temp = Form1.GetData
    If Temp <= TimingSlider.Max Then
        If Temp >= TimingSlider.Min Then
            TimingSlider.Value = Temp 'Form1.GetData  'Get Value from Controller
        End If
    End If
End Sub
Private Sub RepsSlider_Change()
    RepsTag.Caption = RepsSlider
End Sub
Private Sub RepsSlider_Scroll()
    RepsTag.Caption = RepsSlider
End Sub
Private Sub StoreTCycle_Click()
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)     'Timer/Setup Branch Commands
    Form1.PutData (146)    'Set Test Cycle Data Command
    Form1.PutData (TCycle) 'Set Value into Controller
    Temp = Form1.GetData
End Sub
Private Sub TCycle_Change()
    TCycleTag.Caption = TCycle
End Sub
Private Sub TCycle_Scroll()
    TCycleTag.Caption = TCycle
End Sub
Private Sub TimingSlider_Change()
    TimingTag.Caption = TimingSlider
End Sub
Private Sub TimingSlider_Scroll()
    TimingTag.Caption = TimingSlider
End Sub
Private Sub BanksSlider_Change()
    BanksTag.Caption = BanksSlider
End Sub
Public Function Check_Store()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (33)               'Send Command to Test 2-Way Communication
    Temp = Form1.GetData
    If Temp = 85 Or Temp = 86 Then          'Read Status from Relay Board
        Form1.T2Way.Caption = "PASS"              '2-Way Communication Test Passed
        If Temp = 85 Then
            Form1.Label21.BackColor = &HC000&
            Form1.Label21.Caption = " RUN  MODE"
            Check_Store = -1
            If Form1.INTER(1).Caption <> "RS-232" Then
                MsgBox ("Your Settings Cannot be Stored, This Device MUST be in Configuration (or Program) Mode to Store this Setting.  Change the PGM/RUN Jumper to Program and Try Again.")
            Else
                Check_Store = 2 'RS-232 Mode
            '    MsgBox ("Your Settings Cannot be Stored, This Device MUST be in Configuration (or Program) Mode to Store this Setting.  Turn Off All DIP Switches, Power Cycle the Controller, and Try Again.")
            End If
        End If
        If Temp = 86 Then
            Form1.Label21.BackColor = &HFF&
            Form1.Label21.Caption = "CONFIG MODE"
            Check_Store = 1
        End If
    Else
        Form1.T2Way.Caption = "FAIL"
        MsgBox ("Unable to Communicate with the Controller.")
        Check_Store = -1
    End If
End Function

Private Sub BanksSlider_Scroll()
    BanksTag.Caption = BanksSlider
End Sub
Private Sub WriteBanks_Click()
    Temp = Check_Store
    If Temp = 1 Or Temp = 2 Then
        If Form1.INTER(1).Caption = "RS-232" And Form1.INTER(0).Caption = "ProXR" Then
            Form1.BaudSet(5).Value = True
        End If
        Form1.PutData (254)         'Enter Command Mode
        Form1.PutData (50)          'Timer/Setup Branch Commands
        Form1.PutData (141)         'Set Attached Banks Command
        Form1.PutData (BanksSlider) 'Send Value to Controller
        Temp = Form1.GetData
    End If
End Sub
Private Sub WriteReps_Click()
    Check_Store
    Form1.PutData (254)        'Enter Command Mode
    Form1.PutData (50)         'Timer/Setup Branch Commands
    Form1.PutData (137)        'Set Reps Command
    Form1.PutData (RepsSlider) 'Send Value to Controller
    Temp = Form1.GetData
End Sub
Private Sub WriteTiming_Click()
    If Check_Store = 1 Then
        If Form1.INTER(1).Caption = "RS-232" And Form1.INTER(0).Caption = "ProXR" Then
            Form1.BaudSet(5).Value = True
        End If
        Form1.PutData (254)            'Enter Command Mode
        Form1.PutData (50)             'Timer/Setup Branch Commands
        Form1.PutData (139)            'Set Serial Timing Command
        Form1.PutData (TimingSlider)   'Send Value to Controller
        Temp = Form1.GetData
    End If
End Sub
