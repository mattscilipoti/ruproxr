VERSION 5.00
Begin VB.Form ScanSwitch 
   Caption         =   "ScanSwitch Input Expansion Modules"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5325
   LinkTopic       =   "Form3"
   ScaleHeight     =   4230
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Device Returned the Following Data:"
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5055
      Begin VB.Line Line8 
         X1              =   3240
         X2              =   4680
         Y1              =   960
         Y2              =   1560
      End
      Begin VB.Line Line7 
         X1              =   3000
         X2              =   3960
         Y1              =   960
         Y2              =   1560
      End
      Begin VB.Line Line6 
         X1              =   2760
         X2              =   3360
         Y1              =   960
         Y2              =   1560
      End
      Begin VB.Line Line5 
         X1              =   2520
         X2              =   2760
         Y1              =   960
         Y2              =   1560
      End
      Begin VB.Line Line4 
         X1              =   2400
         X2              =   2160
         Y1              =   960
         Y2              =   1560
      End
      Begin VB.Line Line3 
         X1              =   2280
         X2              =   1680
         Y1              =   960
         Y2              =   1560
      End
      Begin VB.Line Line2 
         X1              =   2040
         X2              =   1080
         Y1              =   960
         Y2              =   1560
      End
      Begin VB.Line Line1 
         X1              =   1680
         X2              =   360
         Y1              =   960
         Y2              =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Add All Small Numbers Above to Get the Returned Value"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   4815
      End
      Begin VB.Label BitBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   22
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label BitBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   21
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label BitBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   20
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label BitBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   19
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label BitBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   18
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label BitBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   17
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label BitBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   16
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label BitBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label ValBin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   4320
         TabIndex        =   12
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label ValBin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   3720
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label ValBin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   3120
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label ValBin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   2520
         TabIndex        =   9
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label ValBin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   1920
         TabIndex        =   8
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label ValBin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label ValBin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label ValBin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Value 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please Choose a Input Bank to Read"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CheckBox LoopRead 
         Caption         =   "Loop (repeat read bank operation constantly)"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton ReadBank 
         Caption         =   "Read Bank"
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar BankSelect 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   1
         Top             =   240
         Width           =   2415
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
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "ScanSwitch"
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
            Form1.PutData (254)
            Form1.PutData (175)
            Form1.PutData (BankSelect.Value)
            Value.Caption = Form1.GetData
            Value_Click
        ReadBank.Enabled = True
    Loop Until LoopRead.Value = 0
End Sub
Public Sub Value_Click()
    For X = 0 To 7
        ValBin(X).Caption = 0
        BitBal(X).Caption = 0
    Next X
    If (Value.Caption And 128) = 128 Then ValBin(0).Caption = 1: BitBal(0).Caption = 128
    If (Value.Caption And 64) = 64 Then ValBin(1).Caption = 1: BitBal(1).Caption = 64
    If (Value.Caption And 32) = 32 Then ValBin(2).Caption = 1: BitBal(2).Caption = 32
    If (Value.Caption And 16) = 16 Then ValBin(3).Caption = 1: BitBal(3).Caption = 16
    If (Value.Caption And 8) = 8 Then ValBin(4).Caption = 1: BitBal(4).Caption = 8
    If (Value.Caption And 4) = 4 Then ValBin(5).Caption = 1: BitBal(5).Caption = 4
    If (Value.Caption And 2) = 2 Then ValBin(6).Caption = 1: BitBal(6).Caption = 2
    If (Value.Caption And 1) = 1 Then ValBin(7).Caption = 1: BitBal(7).Caption = 1
End Sub
