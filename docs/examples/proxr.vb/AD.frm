VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AD 
   Caption         =   "A/D Enhancements"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form3"
   ScaleHeight     =   4245
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Voltage"
      Height          =   4095
      Left            =   9360
      TabIndex        =   39
      Top             =   120
      Width           =   1695
      Begin VB.Label V 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   7
         Left            =   120
         TabIndex        =   47
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label V 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   6
         Left            =   120
         TabIndex        =   46
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label V 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   5
         Left            =   120
         TabIndex        =   45
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label V 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   4
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label V 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   3
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label V 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label V 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label V 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Value"
      Height          =   4095
      Left            =   7560
      TabIndex        =   26
      Top             =   120
      Width           =   1695
      Begin VB.Label ADC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   7
         Left            =   120
         TabIndex        =   34
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label ADC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label ADC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label ADC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label ADC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label ADC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label ADC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label ADC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
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
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "10-Bit"
      Height          =   4095
      Left            =   1680
      TabIndex        =   17
      Top             =   0
      Width           =   1455
      Begin VB.CheckBox Loop10 
         Caption         =   "Loop"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton ReadAll10 
         Caption         =   "Read All"
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 7"
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 6"
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 5"
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 4"
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 3"
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 2"
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 1"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 0"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "8-Bit"
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   1455
      Begin VB.CheckBox Loop8 
         Caption         =   "Loop"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton ReadAll8 
         Caption         =   "Read All"
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 0"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 1"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 2"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 3"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 4"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 5"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 6"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton ReadAD 
         Caption         =   "Read AD 7"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
   End
   Begin MSComctlLib.ProgressBar ADLevel 
      Height          =   405
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ADLevel 
      Height          =   405
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ADLevel 
      Height          =   405
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ADLevel 
      Height          =   405
      Index           =   3
      Left            =   3240
      TabIndex        =   3
      Top             =   1800
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ADLevel 
      Height          =   405
      Index           =   4
      Left            =   3240
      TabIndex        =   4
      Top             =   2280
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ADLevel 
      Height          =   405
      Index           =   5
      Left            =   3240
      TabIndex        =   5
      Top             =   2760
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ADLevel 
      Height          =   405
      Index           =   6
      Left            =   3240
      TabIndex        =   6
      Top             =   3240
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ADLevel 
      Height          =   405
      Index           =   7
      Left            =   3240
      TabIndex        =   7
      Top             =   3720
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
End
Attribute VB_Name = "AD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Form1.Visible = True
    Form1.ADFeatures.Enabled = True
End Sub
Private Sub Loop10_Click()
    Form1.Visible = True
    If Loop8.Value = 1 Then
        Form1.Visible = False
    End If
    If Loop10.Value = 1 Then
        Form1.Visible = False
    End If
End Sub
Private Sub Loop8_Click()
    Form1.Visible = True
    If Loop8.Value = 1 Then
        Form1.Visible = False
    End If
    If Loop10.Value = 1 Then
        Form1.Visible = False
    End If
End Sub

Private Sub ReadAD_Click(Index As Integer)
    Form1.ProXR1.ClearBuffer
    Form1.PutData (254)
    Form1.PutData (150 + Index)
    If Index < 8 Then
        ADLevel(Index).Max = 255
        ADLevel(Index).Value = Form1.GetData
        ADC(Index).Caption = Str$(ADLevel(Index).Value) + "/255"
        V(Index).Caption = ADLevel(Index) * 0.01953125
    Else
        MSB = Form1.GetData
        LSB = Form1.GetData
        ADLevel(Index - 8).Max = 1023
        'ADLevel(Index - 8).Value = LSB + (MSB * 256)
        'ADC(Index - 8).Caption = Str$(ADLevel(Index - 8).Value) + "/1023"
        ADCalc = LSB + (MSB * 256)
        If ADCalc <= ADLevel(Index - 8).Max Then
            ADLevel(Index - 8).Value = ADCalc
            ADC(Index - 8).Caption = Str$(ADLevel(Index - 8).Value) + "/1023"
            V(Index - 8).Caption = Mid$(ADLevel(Index - 8) * 0.0048828125, 1, 6)
        Else
            ADC(Index - 8).Caption = "Read Error"
        End If

    End If
End Sub
Private Sub ReadAll8_Click()
Form1.ProXR1.ClearBuffer
Do
    ReadAll8.Enabled = False
    Form1.PutData (254)
    Form1.PutData (166)
    For Index = 0 To 7
        ADLevel(Index).Max = 255
        ADLevel(Index).Value = Form1.GetDataQuiet
        ADC(Index).Caption = Str$(ADLevel(Index).Value) + "/255"
        V(Index).Caption = Mid$(ADLevel(Index) * 0.01953125, 1, 6)
    Next Index
    ReadAll8.Enabled = True
Loop Until Loop8.Value = False
End Sub
Private Sub ReadAll10_Click()
Form1.ProXR1.ClearBuffer
Do
    ReadAll10.Enabled = False
    Form1.PutData (254)
    Form1.PutData (167)
    For Index = 0 To 7
        MSB = Form1.GetDataQuiet
        LSB = Form1.GetDataQuiet
        ADLevel(Index).Max = 1023
        ADCalc = LSB + (MSB * 256)
        If ADCalc <= ADLevel(Index).Max Then
            ADLevel(Index).Value = ADCalc
            ADC(Index).Caption = Str$(ADLevel(Index).Value) + "/1023"
            V(Index).Caption = Mid$(ADLevel(Index) * 0.0048828125, 1, 6)
        Else
            ADC(Index).Caption = "Read Error"
        End If
    Next Index
    ReadAll10.Enabled = True
Loop Until Loop10.Value = False
End Sub

