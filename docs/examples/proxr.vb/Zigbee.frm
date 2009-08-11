VERSION 5.00
Begin VB.Form Zigbee 
   Caption         =   "Zigbee Special Features"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form3"
   ScaleHeight     =   7860
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton DeleteZigBeeDevice 
      Caption         =   "Delete this ZigBee Device"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add ZigBee Device to the List Above"
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   8775
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton AddSerial 
         Caption         =   "Add ZigBee Device to the List"
         Height          =   855
         Left            =   4800
         TabIndex        =   6
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label2 
         Caption         =   "Serial Number Part 2:"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Serial Number Part 1:"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Location 
         Caption         =   "Where is the ZigBee Device Located (for easy reference)"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2640
      Width           =   5895
   End
   Begin VB.CommandButton TargetZigBee 
      Caption         =   "Send Commands to a Specific ZigBee Device with the Following Serial Number:"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2160
      Width           =   5775
   End
   Begin VB.CommandButton AllZigBee 
      Caption         =   "Send Commands to All ZigBee Devices"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label DID 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Zigbee.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "Zigbee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Form1.ZigbeeFeatures.Enabled = True
End Sub

