VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ScanVolt 
   Caption         =   "ScanVolt Input Expansion Modules"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form3"
   ScaleHeight     =   7440
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ProgressBar Volt 
      Height          =   615
      Index           =   0
      Left            =   2880
      TabIndex        =   5
      Top             =   1920
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Volt 
      Height          =   615
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   2520
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Volt 
      Height          =   615
      Index           =   2
      Left            =   2880
      TabIndex        =   8
      Top             =   3120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Volt 
      Height          =   615
      Index           =   3
      Left            =   2880
      TabIndex        =   9
      Top             =   3720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Volt 
      Height          =   615
      Index           =   4
      Left            =   2880
      TabIndex        =   10
      Top             =   4320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Volt 
      Height          =   615
      Index           =   5
      Left            =   2880
      TabIndex        =   11
      Top             =   4920
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Volt 
      Height          =   615
      Index           =   6
      Left            =   2880
      TabIndex        =   12
      Top             =   5520
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Volt 
      Height          =   615
      Index           =   7
      Left            =   2880
      TabIndex        =   13
      Top             =   6120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please Choose a Input Bank to Read"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin VB.CheckBox LoopRead 
         Caption         =   "Loop (repeat read bank operation constantly)"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton ReadBank 
         Caption         =   "Read Bank"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.HScrollBar BankSelect 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   1
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label BankLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Notice: Only AC Voltages can be Read by this Controller.  DC Voltages can only be Detected (voltage present/not present)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   120
      TabIndex        =   39
      Top             =   6840
      Width           =   8415
   End
   Begin VB.Label Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   7
      Left            =   1320
      TabIndex        =   28
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   6
      Left            =   1320
      TabIndex        =   27
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   5
      Left            =   1320
      TabIndex        =   26
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   4
      Left            =   1320
      TabIndex        =   25
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   1320
      TabIndex        =   24
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   1320
      TabIndex        =   23
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   1320
      TabIndex        =   22
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Voltage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   7
      Left            =   6840
      TabIndex        =   20
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Voltage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   6
      Left            =   6840
      TabIndex        =   19
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Voltage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   5
      Left            =   6840
      TabIndex        =   18
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Voltage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   4
      Left            =   6840
      TabIndex        =   17
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Voltage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   6840
      TabIndex        =   16
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Voltage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   6840
      TabIndex        =   15
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Voltage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   6840
      TabIndex        =   14
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Voltage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   6840
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Program Computed Voltage Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   1
      Left            =   6840
      TabIndex        =   30
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Controller Returned this Value:"
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
      Index           =   0
      Left            =   1320
      TabIndex        =   29
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Input 8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   0
      TabIndex        =   38
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Input 7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   37
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Input 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   36
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Input 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   35
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Input 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   34
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Input 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   33
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Input 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   32
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Input 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   31
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "ScanVolt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BankSelect_Change()
    BankLabel.Caption = BankSelect.Value
End Sub
Private Sub BankSelect_Scroll()
    BankLabel.Caption = BankSelect.Value
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Form1.CCFeatures.Enabled = True
End Sub
Public Sub ReadBank_Click()
    Do
        ReadBank.Enabled = False
            For Channel = 0 To 7
                Form1.PutData (254)
                Form1.PutData (176)
                Form1.PutData (BankSelect.Value)
                Form1.PutData (Channel)  'Select Channel to Read
                Value(Channel).Caption = Form1.GetData
                If Val(Value(Channel).Caption) < 101 Then
                    If Value(Channel).Caption <> "" Then
                        Volt(Channel) = Value(Channel).Caption
                        Voltage(Channel) = LookupVolt(Volt(Channel))
                    End If
                End If
            Next Channel
        ReadBank.Enabled = True
    Loop Until LoopRead.Value = 0
End Sub
'Public Sub Value_Click()
'    For X = 0 To 7
'        ValBin(X).Caption = 0
'        BitBal(X).Caption = 0
'    Next X
'    If (Value.Caption And 128) = 128 Then ValBin(0).Caption = 1: BitBal(0).Caption = 128
'    If (Value.Caption And 64) = 64 Then ValBin(1).Caption = 1: BitBal(1).Caption = 64
'    If (Value.Caption And 32) = 32 Then ValBin(2).Caption = 1: BitBal(2).Caption = 32
'    If (Value.Caption And 16) = 16 Then ValBin(3).Caption = 1: BitBal(3).Caption = 16
'    If (Value.Caption And 8) = 8 Then ValBin(4).Caption = 1: BitBal(4).Caption = 8
'    If (Value.Caption And 4) = 4 Then ValBin(5).Caption = 1: BitBal(5).Caption = 4
'    If (Value.Caption And 2) = 2 Then ValBin(6).Caption = 1: BitBal(6).Caption = 2
'    If (Value.Caption And 1) = 1 Then ValBin(7).Caption = 1: BitBal(7).Caption = 1
'End Sub
Public Function LookupVolt(Volt)
    If Volt = 0 Then LookupVolt = 0: Exit Function
    If Volt <= 6 Then LookupVolt = 16: Exit Function
    If Volt <= 21 Then LookupVolt = 17: Exit Function
    If Volt <= 28 Then LookupVolt = 18: Exit Function
    If Volt <= 39 Then LookupVolt = 19: Exit Function
    If Volt <= 47 Then LookupVolt = 20: Exit Function
    If Volt <= 49 Then LookupVolt = 21: Exit Function
    If Volt <= 52 Then LookupVolt = 22: Exit Function
    If Volt <= 55 Then LookupVolt = 23: Exit Function
    If Volt <= 58 Then LookupVolt = 24: Exit Function
    If Volt <= 60 Then LookupVolt = 25: Exit Function
    If Volt <= 62 Then LookupVolt = 26: Exit Function
    If Volt <= 64 Then LookupVolt = 27: Exit Function
    If Volt <= 66 Then LookupVolt = 28: Exit Function
    If Volt <= 67 Then LookupVolt = 29: Exit Function
    If Volt <= 68 Then LookupVolt = 30: Exit Function
    If Volt <= 70 Then LookupVolt = 31: Exit Function
    If Volt <= 71 Then LookupVolt = 32: Exit Function
    If Volt <= 72 Then LookupVolt = 33: Exit Function
    If Volt <= 73 Then LookupVolt = 34: Exit Function
    If Volt <= 74 Then LookupVolt = 35: Exit Function
    If Volt <= 76 Then LookupVolt = 36: Exit Function
    If Volt <= 77 Then LookupVolt = 38: Exit Function
    If Volt <= 78 Then LookupVolt = 39: Exit Function
    If Volt <= 79 Then LookupVolt = 40: Exit Function
    If Volt <= 80 Then LookupVolt = 42: Exit Function
    If Volt <= 81 Then LookupVolt = 43: Exit Function
    If Volt <= 82 Then LookupVolt = 44: Exit Function
    If Volt <= 83 Then LookupVolt = 45: Exit Function
    If Volt <= 255 Then LookupVolt = "HIGH": Exit Function
End Function

