VERSION 5.00
Begin VB.Form ZigBeeMirror 
   Caption         =   "ZigBee Mirror Module Configuration"
   ClientHeight    =   15585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form4"
   ScaleHeight     =   15585
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "Step 3: How Many Inputs will you be Using on your Mirror Module?"
      Height          =   735
      Left            =   120
      TabIndex        =   32
      Top             =   2640
      Width           =   5055
      Begin VB.HScrollBar MirrorInputs 
         Height          =   255
         Left            =   120
         Max             =   32
         Min             =   1
         TabIndex        =   33
         Top             =   360
         Value           =   1
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "8"
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
         Left            =   4320
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Step 2: Manage Serial Numbers"
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Top             =   1680
      Width           =   5055
      Begin VB.CommandButton Command3 
         Caption         =   "Manage ZigBee Serial Numbers"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Step 1: Select a ZigBee Mirror Mode"
      Height          =   1455
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   11775
      Begin VB.OptionButton MODE 
         Caption         =   $"ZigBeeMirror.frx":0000
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   11415
      End
      Begin VB.OptionButton MODE 
         Caption         =   $"ZigBeeMirror.frx":009B
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   10935
      End
      Begin VB.OptionButton MODE 
         Caption         =   "Mode 1: ZigBee Mirror Module Talks to a Specific ZigBee Device using a Single Serial Number"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   8895
      End
      Begin VB.OptionButton MODE 
         Caption         =   "Mode 0: ZigBee Mirror Module Talks to All ZigBee Series Devices within Range - Broadcasts to All Serial Numbers"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   8895
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Map a Mirror Module Input to a ZigBee Controller and Relay Number:"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   11400
      Width           =   11415
      Begin VB.Frame Frame1 
         Caption         =   "Send the Following Command when Input is Turned ON"
         Height          =   615
         Left            =   5160
         TabIndex        =   15
         Top             =   240
         Width           =   6135
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   23
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   22
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1920
            TabIndex        =   21
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   2520
            TabIndex        =   20
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   3120
            TabIndex        =   19
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   3720
            TabIndex        =   18
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   7
            Left            =   4320
            TabIndex        =   17
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Help Me"
            Height          =   255
            Left            =   4920
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Send the Following Command when Input is Turned OFF"
         Height          =   615
         Left            =   5160
         TabIndex        =   5
         Top             =   960
         Width           =   6135
         Begin VB.CommandButton Command2 
            Caption         =   "Help Me"
            Height          =   255
            Left            =   4920
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   15
            Left            =   120
            TabIndex        =   13
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   14
            Left            =   720
            TabIndex        =   12
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   1320
            TabIndex        =   11
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   1920
            TabIndex        =   10
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   2520
            TabIndex        =   9
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   10
            Left            =   3120
            TabIndex        =   8
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   9
            Left            =   3720
            TabIndex        =   7
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox ONN 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   8
            Left            =   4320
            TabIndex        =   6
            Text            =   "---"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Text            =   "16-Digit Serial Number"
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "HELP"
      Height          =   1695
      Left            =   5280
      TabIndex        =   0
      Top             =   1680
      Width           =   6615
      Begin VB.Label HELP 
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Mirror Module Input #1 Talks To ZigBee Device:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   13560
      Width           =   5055
   End
End
Attribute VB_Name = "ZigBeeMirror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
    ZigBeeSerials.Visible = True
    ZigBeeSerials.ZOrder 0
End Sub

Private Sub MirrorInputs_Change()
    Label1.Caption = MirrorInputs * 8
End Sub
Private Sub MirrorInputs_Scroll()
    Label1.Caption = MirrorInputs * 8
End Sub
