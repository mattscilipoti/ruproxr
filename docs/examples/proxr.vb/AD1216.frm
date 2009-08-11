VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AD1216
   Caption         =   "8-Bit /12-Bit Analog to Digital Conversion Features"
   ClientHeight    =   11535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form3"
   Picture         =   "AD1216.frx":0000
   ScaleHeight     =   11535
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3
      Caption         =   "Data Method"
      Height          =   735
      Left            =   8040
      TabIndex        =   203
      Top             =   1560
      Width           =   2055
      Begin VB.OptionButton Method
         Caption         =   "Framed"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   205
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Method
         Caption         =   "RAW"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   204
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame2
      Caption         =   "Resolution"
      Height          =   615
      Left            =   8040
      TabIndex        =   199
      Top             =   840
      Width           =   2055
      Begin VB.OptionButton RES
         Caption         =   "12-Bit"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   201
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton RES
         Caption         =   "8-Bit"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   200
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CheckBox LoopRead
      Caption         =   "Loop"
      Height          =   255
      Left            =   8040
      TabIndex        =   54
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Frame Frame1
      Caption         =   "Channels"
      Height          =   615
      Left            =   8040
      TabIndex        =   50
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Channels
         Caption         =   "48"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   53
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Channels
         Caption         =   "32"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   52
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Channels
         Caption         =   "16"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton All_At_A_Time
      Caption         =   "Read 16 Inputs at a Time"
      Height          =   375
      Left            =   8040
      TabIndex        =   49
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton One_At_A_Time
      Caption         =   "Read 1 Input at a Time"
      Height          =   375
      Left            =   8040
      TabIndex        =   48
      Top             =   2760
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   1680
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   8
      Top             =   1920
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Top             =   2400
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   11
      Top             =   2640
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   12
      Left            =   480
      TabIndex        =   12
      Top             =   2880
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   13
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   14
      Left            =   480
      TabIndex        =   14
      Top             =   3360
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   15
      Left            =   480
      TabIndex        =   15
      Top             =   3600
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   16
      Left            =   480
      TabIndex        =   16
      Top             =   3840
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   17
      Left            =   480
      TabIndex        =   17
      Top             =   4080
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   18
      Left            =   480
      TabIndex        =   18
      Top             =   4320
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   19
      Left            =   480
      TabIndex        =   19
      Top             =   4560
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   20
      Left            =   480
      TabIndex        =   20
      Top             =   4800
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   21
      Left            =   480
      TabIndex        =   21
      Top             =   5040
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   22
      Left            =   480
      TabIndex        =   22
      Top             =   5280
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   23
      Left            =   480
      TabIndex        =   23
      Top             =   5520
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   24
      Left            =   480
      TabIndex        =   24
      Top             =   5760
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   25
      Left            =   480
      TabIndex        =   25
      Top             =   6000
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   26
      Left            =   480
      TabIndex        =   26
      Top             =   6240
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   27
      Left            =   480
      TabIndex        =   27
      Top             =   6480
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   28
      Left            =   480
      TabIndex        =   28
      Top             =   6720
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   29
      Left            =   480
      TabIndex        =   29
      Top             =   6960
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   30
      Left            =   480
      TabIndex        =   30
      Top             =   7200
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   31
      Left            =   480
      TabIndex        =   31
      Top             =   7440
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   32
      Left            =   480
      TabIndex        =   32
      Top             =   7680
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   33
      Left            =   480
      TabIndex        =   33
      Top             =   7920
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   34
      Left            =   480
      TabIndex        =   34
      Top             =   8160
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   35
      Left            =   480
      TabIndex        =   35
      Top             =   8400
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   36
      Left            =   480
      TabIndex        =   36
      Top             =   8640
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   37
      Left            =   480
      TabIndex        =   37
      Top             =   8880
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   38
      Left            =   480
      TabIndex        =   38
      Top             =   9120
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   39
      Left            =   480
      TabIndex        =   39
      Top             =   9360
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   40
      Left            =   480
      TabIndex        =   40
      Top             =   9600
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   41
      Left            =   480
      TabIndex        =   41
      Top             =   9840
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   42
      Left            =   480
      TabIndex        =   42
      Top             =   10080
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   43
      Left            =   480
      TabIndex        =   43
      Top             =   10320
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   44
      Left            =   480
      TabIndex        =   44
      Top             =   10560
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   45
      Left            =   480
      TabIndex        =   45
      Top             =   10800
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   46
      Left            =   480
      TabIndex        =   46
      Top             =   11040
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar AD
      Height          =   255
      Index           =   47
      Left            =   480
      TabIndex        =   47
      Top             =   11280
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   255
      Scrolling       =   1
   End
   Begin VB.Label HELP
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   7920
      TabIndex        =   202
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   143
      Left            =   6600
      TabIndex        =   198
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   142
      Left            =   6600
      TabIndex        =   197
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   141
      Left            =   6600
      TabIndex        =   196
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   140
      Left            =   6600
      TabIndex        =   195
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   139
      Left            =   6600
      TabIndex        =   194
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   138
      Left            =   6600
      TabIndex        =   193
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   137
      Left            =   6600
      TabIndex        =   192
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   136
      Left            =   6600
      TabIndex        =   191
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   135
      Left            =   6600
      TabIndex        =   190
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   134
      Left            =   6600
      TabIndex        =   189
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   133
      Left            =   6600
      TabIndex        =   188
      Top             =   2400
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   132
      Left            =   6600
      TabIndex        =   187
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   131
      Left            =   6600
      TabIndex        =   186
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   130
      Left            =   6600
      TabIndex        =   185
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   129
      Left            =   6600
      TabIndex        =   184
      Top             =   3360
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   128
      Left            =   6600
      TabIndex        =   183
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   127
      Left            =   6600
      TabIndex        =   182
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   126
      Left            =   6600
      TabIndex        =   181
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   125
      Left            =   6600
      TabIndex        =   180
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   124
      Left            =   6600
      TabIndex        =   179
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   123
      Left            =   6600
      TabIndex        =   178
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   122
      Left            =   6600
      TabIndex        =   177
      Top             =   5040
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   121
      Left            =   6600
      TabIndex        =   176
      Top             =   5280
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   120
      Left            =   6600
      TabIndex        =   175
      Top             =   5520
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   119
      Left            =   6600
      TabIndex        =   174
      Top             =   5760
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   118
      Left            =   6600
      TabIndex        =   173
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   117
      Left            =   6600
      TabIndex        =   172
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   116
      Left            =   6600
      TabIndex        =   171
      Top             =   6480
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   115
      Left            =   6600
      TabIndex        =   170
      Top             =   6720
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   114
      Left            =   6600
      TabIndex        =   169
      Top             =   6960
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   113
      Left            =   6600
      TabIndex        =   168
      Top             =   7200
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   112
      Left            =   6600
      TabIndex        =   167
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   111
      Left            =   6600
      TabIndex        =   166
      Top             =   7680
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   110
      Left            =   6600
      TabIndex        =   165
      Top             =   7920
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   109
      Left            =   6600
      TabIndex        =   164
      Top             =   8160
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   108
      Left            =   6600
      TabIndex        =   163
      Top             =   8400
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   107
      Left            =   6600
      TabIndex        =   162
      Top             =   8640
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   106
      Left            =   6600
      TabIndex        =   161
      Top             =   8880
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   105
      Left            =   6600
      TabIndex        =   160
      Top             =   9120
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   104
      Left            =   6600
      TabIndex        =   159
      Top             =   9360
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   103
      Left            =   6600
      TabIndex        =   158
      Top             =   9600
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   102
      Left            =   6600
      TabIndex        =   157
      Top             =   9840
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   101
      Left            =   6600
      TabIndex        =   156
      Top             =   10080
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   100
      Left            =   6600
      TabIndex        =   155
      Top             =   10320
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   99
      Left            =   6600
      TabIndex        =   154
      Top             =   10560
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   98
      Left            =   6600
      TabIndex        =   153
      Top             =   10800
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   97
      Left            =   6600
      TabIndex        =   152
      Top             =   11040
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   96
      Left            =   6600
      TabIndex        =   151
      Top             =   11280
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   95
      Left            =   5400
      TabIndex        =   150
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   94
      Left            =   5400
      TabIndex        =   149
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   93
      Left            =   5400
      TabIndex        =   148
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   92
      Left            =   5400
      TabIndex        =   147
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   91
      Left            =   5400
      TabIndex        =   146
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   90
      Left            =   5400
      TabIndex        =   145
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   89
      Left            =   5400
      TabIndex        =   144
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   88
      Left            =   5400
      TabIndex        =   143
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   87
      Left            =   5400
      TabIndex        =   142
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   86
      Left            =   5400
      TabIndex        =   141
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   85
      Left            =   5400
      TabIndex        =   140
      Top             =   2400
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   84
      Left            =   5400
      TabIndex        =   139
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   83
      Left            =   5400
      TabIndex        =   138
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   82
      Left            =   5400
      TabIndex        =   137
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   81
      Left            =   5400
      TabIndex        =   136
      Top             =   3360
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   80
      Left            =   5400
      TabIndex        =   135
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   79
      Left            =   5400
      TabIndex        =   134
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   78
      Left            =   5400
      TabIndex        =   133
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   77
      Left            =   5400
      TabIndex        =   132
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   76
      Left            =   5400
      TabIndex        =   131
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   75
      Left            =   5400
      TabIndex        =   130
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   74
      Left            =   5400
      TabIndex        =   129
      Top             =   5040
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   73
      Left            =   5400
      TabIndex        =   128
      Top             =   5280
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   72
      Left            =   5400
      TabIndex        =   127
      Top             =   5520
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   71
      Left            =   5400
      TabIndex        =   126
      Top             =   5760
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   70
      Left            =   5400
      TabIndex        =   125
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   69
      Left            =   5400
      TabIndex        =   124
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   68
      Left            =   5400
      TabIndex        =   123
      Top             =   6480
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   67
      Left            =   5400
      TabIndex        =   122
      Top             =   6720
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   66
      Left            =   5400
      TabIndex        =   121
      Top             =   6960
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   65
      Left            =   5400
      TabIndex        =   120
      Top             =   7200
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   64
      Left            =   5400
      TabIndex        =   119
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   63
      Left            =   5400
      TabIndex        =   118
      Top             =   7680
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   62
      Left            =   5400
      TabIndex        =   117
      Top             =   7920
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   61
      Left            =   5400
      TabIndex        =   116
      Top             =   8160
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   60
      Left            =   5400
      TabIndex        =   115
      Top             =   8400
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   59
      Left            =   5400
      TabIndex        =   114
      Top             =   8640
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   58
      Left            =   5400
      TabIndex        =   113
      Top             =   8880
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   57
      Left            =   5400
      TabIndex        =   112
      Top             =   9120
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   56
      Left            =   5400
      TabIndex        =   111
      Top             =   9360
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   55
      Left            =   5400
      TabIndex        =   110
      Top             =   9600
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   54
      Left            =   5400
      TabIndex        =   109
      Top             =   9840
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   53
      Left            =   5400
      TabIndex        =   108
      Top             =   10080
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   52
      Left            =   5400
      TabIndex        =   107
      Top             =   10320
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   51
      Left            =   5400
      TabIndex        =   106
      Top             =   10560
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   50
      Left            =   5400
      TabIndex        =   105
      Top             =   10800
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   49
      Left            =   5400
      TabIndex        =   104
      Top             =   11040
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   48
      Left            =   5400
      TabIndex        =   103
      Top             =   11280
      Width           =   1200
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   47
      Left            =   0
      TabIndex        =   102
      Top             =   11280
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   46
      Left            =   0
      TabIndex        =   101
      Top             =   11040
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   45
      Left            =   0
      TabIndex        =   100
      Top             =   10800
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   44
      Left            =   0
      TabIndex        =   99
      Top             =   10560
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   43
      Left            =   0
      TabIndex        =   98
      Top             =   10320
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   0
      TabIndex        =   97
      Top             =   10080
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   0
      TabIndex        =   96
      Top             =   9840
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   0
      TabIndex        =   95
      Top             =   9600
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   0
      TabIndex        =   94
      Top             =   9360
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   0
      TabIndex        =   93
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   0
      TabIndex        =   92
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   0
      TabIndex        =   91
      Top             =   8640
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   0
      TabIndex        =   90
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   0
      TabIndex        =   89
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   0
      TabIndex        =   88
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   0
      TabIndex        =   87
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   0
      TabIndex        =   86
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   0
      TabIndex        =   85
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   0
      TabIndex        =   84
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   0
      TabIndex        =   83
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   0
      TabIndex        =   82
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   0
      TabIndex        =   81
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   0
      TabIndex        =   80
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   0
      TabIndex        =   79
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   0
      TabIndex        =   78
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   0
      TabIndex        =   77
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   0
      TabIndex        =   76
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   0
      TabIndex        =   75
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   0
      TabIndex        =   74
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   0
      TabIndex        =   73
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   0
      TabIndex        =   72
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   0
      TabIndex        =   71
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   0
      TabIndex        =   70
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   0
      TabIndex        =   69
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   0
      TabIndex        =   68
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   0
      TabIndex        =   67
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   0
      TabIndex        =   66
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   65
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   64
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   63
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   62
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   61
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   60
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   59
      Top             =   960
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   58
      Top             =   720
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   57
      Top             =   480
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   56
      Top             =   240
      Width           =   495
   End
   Begin VB.Label ReadInput
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "AD1216"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Buffer(32) As Integer

Private Sub Channels_Click(Index As Integer)
    LoopRead.Value = False
    For x = 0 To 47
        AD(x).Value = 0
        ReadInput(x + 47).Caption = ""
        ReadInput(x + 96).Caption = ""
    Next x
End Sub
Private Sub Form_Load()
    For x = 0 To 47
        ReadInput(x).Caption = x
    Next x
End Sub
Private Sub All_At_A_Time_Click()

    Form1.GetDataQuiet
    Form1.GetDataQuiet

    If RES(0).Value = True Then     'If 8-Bit Resolution is Selected
        Do
            If Channels(0).Value = True Then MaxDevice = 0
            If Channels(1).Value = True Then MaxDevice = 1
            If Channels(2).Value = True Then MaxDevice = 2
            OUT = 0
            For Device = 0 To MaxDevice                        'Count Through All Available Devices (3 Max)
Retry:
                 If Method(0).Value = True Then MethodCalc = 0
                 If Method(1).Value = True Then MethodCalc = 8
                 Form1.PutData (254)              'Enter Command Mode
                 Form1.PutData (192 + Device + MethodCalc)     'Get 8-Bit Value
                 If MethodCalc = 8 Then                        'If the Framed Option is Selected
                    CKSUM = Form1.GetDataQuiet2                'Get Header as Base Checksum Value
                 End If
                 For Channel = 0 To 15                         'Get Each Channel Data Value
                    MSB = Form1.GetDataQuiet                   'Read from Controller
                    If MSB <> -1 Then                          'Make sure Data is in Range
                        Buffer(Channel) = MSB                  'Store Data byte
                        CKSUM = CKSUM + MSB                    'Add to Checksum
                        If MethodCalc = 0 Then
                            Display_Data_8Bit Buffer(Channel), OUT 'Display Data
                            OUT = OUT + 1
                        End If
                    End If
                 Next Channel                                  'Get Next Channel
                 If MethodCalc = 8 Then                        'If the Framed Option is Selected
                    HW_Cksum = Form1.GetDataQuiet2             'Store the Hardware Checksum
                    CKSUM = (CKSUM And 255)                    'Extract Only Lower 8 Bits from Computed Checksum
                    If HW_Cksum = CKSUM Then                      'If the Checksums Match
                       Debug.Print "CK ok"
                       For Channel = 0 To 15                      'Display Data
                           Display_Data_8Bit Buffer(Channel), OUT 'Display Data
                           OUT = OUT + 1                          'Choose a New Position on the Display
                       Next Channel
                    Else
                       Debug.Print "CK ERROR"
                       GoTo Retry
                    End If
                 End If
            Next Device


        Loop Until LoopRead.Value = 0 Or AD1216.Visible = False                           'Repeat Operation Until User Clocks the Loop Button Off
    Else
        'If 12-Bit Resolution is Selected
        Do
            If Channels(0).Value = True Then MaxDevice = 0
            If Channels(1).Value = True Then MaxDevice = 1
            If Channels(2).Value = True Then MaxDevice = 2
            OUT = 0
            For Device = 0 To MaxDevice                     'Count Through All Available Devices
Retry2:
                 If Method(0).Value = True Then MethodCalc = 0
                 If Method(1).Value = True Then MethodCalc = 8
                 Form1.PutData (254)           'Enter Command Mode
                 Form1.PutData (196 + Device + MethodCalc)  'Get 12-Bit Value
                 If MethodCalc = 8 Then                     'If the Framed Option is Selected
                    CKSUM = Form1.GetDataQuiet2             'Get Header as Base Checksum Value
                 End If
                 For Channel = 0 To 15                      'Count Through All 16 Channels
                    LSB = Form1.GetDataQuiet                'Get LSB from Controller
                    MSB = Form1.GetDataQuiet                'Get MSB from Controller
                    If MSB <> -1 And LSB <> -1 Then
                        Buffer((Channel * 2)) = LSB
                        Buffer((Channel * 2) + 1) = MSB
                        CKSUM = CKSUM + LSB
                        CKSUM = CKSUM + MSB
                        If MethodCalc = 0 Then
                            Display_Data_12Bit Buffer(Channel * 2), Buffer((Channel * 2) + 1), OUT 'Display Results on AD1216 Form
                            OUT = OUT + 1
                        End If
                    End If
                 Next Channel
                 If MethodCalc = 8 Then                        'If the Framed Option is Selected
                    HW_Cksum = Form1.GetDataQuiet2             'Store the Hardware Checksum
                    CKSUM = (CKSUM And 255)                    'Extract Only Lower 8 Bits from Computed Checksum
                    If HW_Cksum = CKSUM Then                      'If the Checksums Match
                       Debug.Print "CK ok"
                       For Channel = 0 To 15
                          Display_Data_12Bit Buffer(Channel * 2), Buffer((Channel * 2) + 1), OUT 'Display Results on AD1216 Form
                          OUT = OUT + 1                       'Choose a New Position on the Display
                       Next Channel
                    Else
                       Debug.Print "CK ERROR"
                       GoTo Retry2
                    End If
                 End If
            Next Device
        Loop Until LoopRead.Value = 0 Or AD1216.Visible = False                      'Repeat Operation Until User Clocks the Loop Button Off
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    LoopRead.Value = 0
    Form1.AD1216_Button.Enabled = True
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HELP.Caption = "Select the number of input channels you would like to monitor.  The fewer the channels, the faster the refresh rate.  Choose the option that matches the number of inputs on your controller."
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HELP.Caption = "Resolution allows you to choose the type of A/D data that is sent to your computer.  When 8-Bit is chosen, a 0-5VDC voltage will be assigned a value from 0-255, this value is sent directly to your computer.  When 12-Bit is selected, voltages from 0-5VDC will be converted into a value from 0 to 4095.  Each channel will be sent to your computer using 2 bytes in the order of Least Significant Byte, then Most Significant Byte.  Your software will need to use the formula Value=(MSB*256)+LSB to rebuild the value into a 0-4095 range.  12-Bit values are more acurate, but take longer to transfer from the controller to the PC."
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HELP.Caption = "The Data Method Only pertains to 16-Channel A/D data requests.  When RAW is selected, only 16 bytes of A/D data is sent to your computer (8-bit mode), and 32 bytes will be sent for 12-bit mode.  When Framed is selected, the first byte will be a 254 header byte, the last byte sent will be an 8-bit checksum including the 254 header."
End Sub
Private Sub LoopRead_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    HELP.Caption = "Check the Loop option if you would like to constantly query the controller for A/D data.  Choosing any of the radio buttons above will cancel the Loop option."
End Sub
Private Sub Method_Click(Index As Integer)
    LoopRead.Value = 0
End Sub
Private Sub One_At_A_Time_Click()
    Do
        If Channels(0).Value = True Then MaxDevice = 0
        If Channels(1).Value = True Then MaxDevice = 1
        If Channels(2).Value = True Then MaxDevice = 2
        OUT = 0
        For Device = 0 To MaxDevice
            For Channel = 0 To 15
                If RES(0).Value = True Then                     '8-Bit Resolution Selected
                    If Device = 0 Then Options = 3              'Device 0 8-Bit
                    If Device = 1 Then Options = 11             'Device 1 8-Bit
                    If Device = 2 Then Options = 16             'Device 2 8-Bit
                    Form1.PutData (254)            'Enter Command Mode
                    Form1.PutData (192 + Options)  'Get 8-Bit Value + Options
                    Form1.PutData (Channel)        'Choose a Channel
                    MSB = Form1.GetDataQuiet2                   'Get 8-Bit Value (LSB)
                    Debug.Print Device; Channel; MSB
                    If MSB <> -1 Then                           'Make Sure Data Arrived
                        Display_Data_8Bit MSB, OUT              'Display Results on AD1216 Form
                    End If
                    OUT = OUT + 1                               'Choose a New Location for Next Data
                Else                                            '12-Bit Resolution Selected
                    If Device = 0 Then Options = 7              'Device 0 12-Bit
                    If Device = 1 Then Options = 15             'Device 1 12-Bit
                    If Device = 2 Then Options = 17             'Device 2 12-Bit
                    Form1.PutData (254)            'Enter Command Mode
                    Form1.PutData (192 + Options)  'Get 8-Bit Value + Options
                    Form1.PutData (Channel)        'Choose a Channel
                    LSB = Form1.GetDataQuiet2                   'Get 12-Bit LSB
                    MSB = Form1.GetDataQuiet2                   'Get 12-Bit MSB
                    If LSB <> -1 And MSB <> -1 Then             'Make Sure Data Arrived
                        Display_Data_12Bit LSB, MSB, OUT        'Display Results on AD1216 Form
                    End If
                    OUT = OUT + 1                               'Choose a New Location for Next Data
                End If
            Next Channel
        Next Device
    Loop Until LoopRead.Value = 0 Or AD1216.Visible = False
End Sub
Public Sub Display_Data_8Bit(MSB, OUT)
    AD(OUT).Max = 255
    AD(OUT).Value = MSB
    ReadInput((48 - OUT) + 47).Caption = str$(AD(OUT).Value) + "  /  255     "
    ReadInput((48 - OUT) + 95).Caption = Mid$(str$(AD(OUT).Value * 0.019607), 2, 3) + " Volts    "
/sssasdasdasdasdaasdasdasdasdEnd Sub
Public Sub Display_Data_12Bit(LSB, MSB, OUT)
    DATS = (MSB * 256) + LSB
    If DATS > 4095 Then DATS = 4095
    If DATS < 0 Then DATS = 0
    AD(OUT).Max = 4095
    AD(OUT).Value = DATS
    ReadInput((48 - OUT) + 47).Caption = str$(AD(OUT).Value) + "  /  4095 "
    ReadInput((48 - OUT) + 95).Caption = Mid$(str$(DATS * 0.001221001221001), 2, 6) + " Volts   "
End Sub
Private Sub RES_Click(Index As Integer)
    LoopRead.Value = False
End Sub
Public Sub dela()
'Debug.Print "DELA"
'    For N = 1 To 5000
'        DoEvents
'    Next N
End Sub
