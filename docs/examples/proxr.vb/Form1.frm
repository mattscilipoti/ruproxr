VERSION 5.00
Object = "{02E0654E-AAC5-4BBF-A1DE-45576B24DFC1}#2.1#0"; "ProXR.ocx"
Begin VB.Form Form1 
   Caption         =   "ProXR Example Software for Visual Basic 6                                                WWW.CONTROLANYTHING.COM"
   ClientHeight    =   12900
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   860
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   3  'Windows Default
   Begin NCDProXR.ProXR ProXR1 
      Left            =   4080
      Top             =   12480
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Bluetooth  Features"
      Height          =   495
      Left            =   8400
      TabIndex        =   171
      Top             =   10920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Fiber Optic  Features"
      Height          =   255
      Left            =   7920
      TabIndex        =   170
      Top             =   11160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton PWM12_Button 
      Caption         =   "Light Dimmer 12-Channel"
      Height          =   495
      Left            =   3360
      TabIndex        =   169
      Top             =   10560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Test Misc Commands"
      Height          =   375
      Left            =   120
      TabIndex        =   168
      Top             =   12480
      Width           =   3015
   End
   Begin VB.CommandButton TestCycle 
      Caption         =   "Relay Test Cycle"
      Height          =   255
      Left            =   5640
      TabIndex        =   167
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton RS232Baud 
      Caption         =   "RS-232 Features"
      Height          =   255
      Left            =   6000
      TabIndex        =   166
      Top             =   11160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton NetworkFeatures 
      Caption         =   "Ethernet / 802.11b Features"
      Height          =   495
      Left            =   9960
      TabIndex        =   164
      Top             =   10560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton ZigbeeFeatures 
      Caption         =   "Zigbee Features"
      Height          =   255
      Left            =   120
      TabIndex        =   163
      Top             =   11160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton USBFeatures 
      Caption         =   "USB Features"
      Height          =   255
      Left            =   4320
      TabIndex        =   162
      Top             =   11160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton VINFeatures 
      Caption         =   "Voltage Detection Features"
      Height          =   495
      Left            =   6720
      TabIndex        =   161
      Top             =   10560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CCFeatures 
      Caption         =   "External Input Features"
      Height          =   495
      Left            =   5520
      TabIndex        =   160
      Top             =   10560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton DACFeatures 
      Caption         =   "DAC Features"
      Height          =   495
      Left            =   4560
      TabIndex        =   159
      Top             =   10560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton AD1216_Button 
      Caption         =   "12-Bit A/D Features"
      Height          =   495
      Left            =   1080
      TabIndex        =   158
      Top             =   10560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton POTFeatures 
      Caption         =   "Potentiometer Features"
      Height          =   495
      Left            =   2160
      TabIndex        =   157
      Top             =   10560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton ADFeatures 
      Caption         =   "A/D Features"
      Height          =   495
      Left            =   120
      TabIndex        =   156
      Top             =   10560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame24 
      Caption         =   "Device Identification"
      Height          =   855
      Left            =   120
      TabIndex        =   134
      Top             =   11520
      Width           =   10215
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Read Device Identification Data"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label DID 
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
         Index           =   4
         Left            =   9360
         TabIndex        =   138
         Top             =   480
         Width           =   735
      End
      Begin VB.Label DID 
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
         Index           =   3
         Left            =   8520
         TabIndex        =   139
         Top             =   480
         Width           =   735
      End
      Begin VB.Label DID 
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
         Index           =   2
         Left            =   7560
         TabIndex        =   140
         Top             =   480
         Width           =   735
      End
      Begin VB.Label DID 
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
         Index           =   1
         Left            =   5400
         TabIndex        =   141
         Top             =   480
         Width           =   615
      End
      Begin VB.Label DID 
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
         Index           =   0
         Left            =   2880
         TabIndex        =   142
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Firmware:"
         Height          =   255
         Left            =   2880
         TabIndex        =   147
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Year:"
         Height          =   255
         Left            =   7560
         TabIndex        =   146
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Version:"
         Height          =   255
         Left            =   8400
         TabIndex        =   145
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Interface:"
         Height          =   255
         Left            =   5280
         TabIndex        =   144
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "E3C:"
         Height          =   255
         Left            =   9240
         TabIndex        =   143
         Top             =   240
         Width           =   855
      End
      Begin VB.Label INTER 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ProXR Core CPU"
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   0
         Left            =   3600
         TabIndex        =   137
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label INTER 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   1
         Left            =   6000
         TabIndex        =   136
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton TimerCalibrationButton 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Timer Calibration"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   10080
      Width           =   1935
   End
   Begin VB.CommandButton AdvancedFeatureButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Advanced Feature Settings"
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   10080
      Width           =   2415
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0080FF80&
      Caption         =   "Relay Timer Commands"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   9600
      Width           =   1935
   End
   Begin VB.Frame Frame23 
      Caption         =   "Set the Status of Relays by Number"
      Height          =   1335
      Left            =   120
      TabIndex        =   115
      Top             =   9120
      Width           =   7215
      Begin VB.HScrollBar HScroll7 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   123
         Top             =   960
         Width           =   5055
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   119
         Top             =   600
         Width           =   5055
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   116
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label15 
         Caption         =   "One Relay On"
         Height          =   255
         Left            =   6000
         TabIndex        =   125
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   124
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Turn On Relay"
         Height          =   255
         Left            =   6000
         TabIndex        =   121
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Turn Off Relay"
         Height          =   255
         Left            =   6000
         TabIndex        =   120
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   118
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "---"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   117
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Manually Refresh All Relay Banks"
      Height          =   495
      Left            =   9960
      TabIndex        =   114
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Store Relay Refresh Settings as Power-Up Default Setting"
      Height          =   495
      Left            =   7440
      TabIndex        =   113
      Top             =   8160
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"Form1.frx":0000
      Height          =   735
      Left            =   7440
      TabIndex        =   112
      Top             =   8760
      Width           =   4335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bank Specified Commands"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Setup Mode Commands:"
      Height          =   1215
      Left            =   7440
      TabIndex        =   89
      Top             =   6840
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "Program E3C Device #"
         Height          =   495
         Left            =   3120
         TabIndex        =   98
         Top             =   120
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   91
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Index           =   10
         Left            =   2400
         TabIndex        =   92
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0FF&
         Caption         =   "This command works in configuration mode ONLY.       Turn off all DIP switches and power cycle the controller."
         Height          =   375
         Left            =   120
         TabIndex        =   90
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.Frame Frame18 
      Caption         =   "Current Relay Bank Memory"
      Height          =   4575
      Left            =   7440
      TabIndex        =   72
      Top             =   2280
      Width           =   4455
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 32"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   2280
         TabIndex        =   131
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 31"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   2280
         TabIndex        =   130
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 30"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   2280
         TabIndex        =   129
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 29"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   2280
         TabIndex        =   128
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 28"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   2280
         TabIndex        =   127
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 27"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   2280
         TabIndex        =   126
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Items in Red Show Power-Up Default Status of Relays"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   4200
         Width           =   4215
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 26"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   2280
         TabIndex        =   108
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 25"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   2280
         TabIndex        =   107
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 24"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   2280
         TabIndex        =   106
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 23"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   2280
         TabIndex        =   105
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 22"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   2280
         TabIndex        =   104
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 21"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   2280
         TabIndex        =   103
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 20"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   2280
         TabIndex        =   102
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 19"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   2280
         TabIndex        =   101
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 18"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   2280
         TabIndex        =   100
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 17"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   2280
         TabIndex        =   99
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 16"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   88
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 15"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   87
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 14"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   86
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 13"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   85
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 12"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   84
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 11"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   83
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 10"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   82
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 9"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   81
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 8"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   80
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 7"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   79
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 6"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   78
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 5"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   77
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 4"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   76
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 3"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   75
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 2"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Banks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bank 1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "E3C Device Networking"
      Height          =   2175
      Left            =   7440
      TabIndex        =   71
      Top             =   0
      Width           =   4455
      Begin VB.Frame Frame21 
         Caption         =   "Select an E3C Device Number"
         Height          =   615
         Left            =   120
         TabIndex        =   148
         Top             =   240
         Width           =   4215
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   149
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Index           =   11
            Left            =   3480
            TabIndex        =   150
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Disable All Devices"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Enable All Devices"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   1560
         Width           =   4215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Enable Selected Device ONLY"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   1320
         Width           =   4215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Enable Selected Device"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Read E3C Device Number from Controller"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   840
         Width           =   4215
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "Select a Relay Bank to Control"
      Height          =   735
      Left            =   120
      TabIndex        =   41
      Top             =   0
      Width           =   7215
      Begin VB.HScrollBar HScroll4 
         Height          =   375
         Left            =   120
         Max             =   32
         TabIndex        =   42
         Top             =   240
         Value           =   32
         Width           =   4935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0 (All Banks)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   5160
         TabIndex        =   43
         Top             =   240
         Width           =   1935
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
      Height          =   4575
      Left            =   120
      TabIndex        =   29
      Top             =   4560
      Width           =   7215
      Begin VB.Frame Frame19 
         Caption         =   "Store Current Relay Settings as Powerup Default"
         Height          =   2415
         Left            =   2280
         TabIndex        =   152
         Top             =   2040
         Width           =   4815
         Begin VB.CommandButton GetBank 
            Caption         =   "Get Stored Powerup Default Relay Pattern"
            Height          =   375
            Left            =   120
            TabIndex        =   154
            Top             =   720
            Width           =   4575
         End
         Begin VB.CommandButton StoreBank 
            Caption         =   "Store Current Relay Pattern as Powerup Default"
            Height          =   375
            Left            =   120
            TabIndex        =   153
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label5 
            Caption         =   $"Form1.frx":009A
            Height          =   975
            Left            =   120
            TabIndex        =   155
            Top             =   1200
            Width           =   4575
         End
      End
      Begin VB.CommandButton Recover 
         Caption         =   "Attempt Recovery"
         Height          =   495
         Left            =   5880
         TabIndex        =   151
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Baud Rate Select"
         Height          =   2415
         Left            =   120
         TabIndex        =   61
         Top             =   2040
         Width           =   2055
         Begin VB.OptionButton BaudSet 
            Caption         =   "115200"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   68
            Top             =   1680
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton BaudSet 
            Caption         =   "57600"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   67
            Top             =   1440
            Width           =   855
         End
         Begin VB.OptionButton BaudSet 
            Caption         =   "38400"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   66
            Top             =   1200
            Width           =   855
         End
         Begin VB.OptionButton BaudSet 
            Caption         =   "19200"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   65
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton BaudSet 
            Caption         =   "9600"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton BaudSet 
            Caption         =   "4800"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   63
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton BaudSet 
            Caption         =   "2400"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
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
            Index           =   9
            Left            =   120
            TabIndex        =   110
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "8 Data Bits 1 Stop Bit   No Parity   (8,N,1)."
            Height          =   975
            Left            =   1080
            TabIndex        =   70
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Set Status of All Relays at Once"
         Height          =   975
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   2535
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   39
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.CommandButton Test2Way 
         Caption         =   "Test 2-Way Communication"
         Height          =   735
         Left            =   5640
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Reverse 
         Caption         =   "Reverse On/Off Relay Pattern"
         Height          =   495
         Left            =   4080
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Invert 
         Caption         =   "Invert Status of All Relays"
         Height          =   495
         Left            =   2760
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton AllOn 
         Caption         =   "All Relays On"
         Height          =   495
         Left            =   1440
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton AllOff 
         Caption         =   "All Relays Off"
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1215
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
         Height          =   255
         Left            =   5880
         TabIndex        =   37
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Relay On/Off Pattern is Reversed, Relay 12345678 status is coppied to Relay 87654321"
         Height          =   975
         Left            =   4080
         TabIndex        =   36
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Relays that are on turn off.  Relays that are off turn on."
         Height          =   855
         Left            =   2760
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Read_All 
      Caption         =   "Read Status of All Relays"
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Read_Relay 
      Caption         =   "Read"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   18
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Relay 1"
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   1080
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7215
      Begin VB.Frame Frame10 
         Caption         =   "Relay 8"
         Height          =   855
         Left            =   6120
         TabIndex        =   58
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Relay 7"
         Height          =   855
         Left            =   5280
         TabIndex        =   56
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Relay 6"
         Height          =   855
         Left            =   4440
         TabIndex        =   54
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Relay 5"
         Height          =   855
         Left            =   3600
         TabIndex        =   52
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Relay 4"
         Height          =   855
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Relay 3"
         Height          =   855
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Relay 2"
         Height          =   855
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   6
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
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   7215
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   7
         Left            =   6120
         TabIndex        =   25
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 8"
         Height          =   855
         Index           =   6
         Left            =   6120
         TabIndex        =   50
         Top             =   360
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
            TabIndex        =   51
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   6
         Left            =   5280
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 7"
         Height          =   855
         Index           =   5
         Left            =   5280
         TabIndex        =   48
         Top             =   360
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
            TabIndex        =   49
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   23
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 6"
         Height          =   855
         Index           =   4
         Left            =   4440
         TabIndex        =   46
         Top             =   360
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
            TabIndex        =   47
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 5"
         Height          =   855
         Index           =   3
         Left            =   3600
         TabIndex        =   44
         Top             =   360
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
            TabIndex        =   45
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 4"
         Height          =   855
         Index           =   2
         Left            =   2760
         TabIndex        =   16
         Top             =   360
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
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 3"
         Height          =   855
         Index           =   1
         Left            =   1920
         TabIndex        =   14
         Top             =   360
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
            TabIndex        =   15
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Relay 2"
         Height          =   855
         Left            =   1080
         TabIndex        =   12
         Top             =   360
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
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 1"
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
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
            TabIndex        =   11
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
      TabIndex        =   27
      Top             =   3480
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "Empty Serial Buffer"
         Height          =   255
         Left            =   5520
         TabIndex        =   69
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Reporting 
         Caption         =   "OFF"
         Height          =   255
         Left            =   5520
         TabIndex        =   60
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   $"Form1.frx":01C3
         Height          =   615
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "----"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   165
      Top             =   11640
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public parentForm As Form

Public Sub LoadForm()
    frmLog.Show
    Set frmLog.proxrObj = ProXR1
    'Form3.Show vbModal, Me
    ProXR1.timeout = 1000
    'ProXR1.OpenPort
    'frmSelectDevice.Show vbModal, Me
    Dim s As String
    s = GetSetting("ProXR_V4", "Settings", "DeviceType", 0)
    Dim i As Integer
    i = Val(s)
    If i = MostNcdDevicesProXr Then   ' zigbee device
        BaudScan    'Find Out What Baud Rate the Controller is On
    End If
    
    Command12_Click
    
    Form1.Visible = True
    'Timers.Visible = True
    Label1(9).Caption = ProXR1.PortName
    ProXR1.ClearBuffer
    Form1.Visible = True
    'Do
    '    DoEvents
    'Loop
    
'    If HScroll4.Max > 1 Then
'        HScroll4.Value = 0  'Set Relay Bank to 0    Triggers the Read Function to Recall Status of All 256 Relays from Controller
'    Else
'        HScroll4.Value = 1
'    End If

 '  Form1.Caption = "ProXR Example Software for Visual Basic 6         COM" + Str$(MSComm1.CommPort) + "    " + MSComm1.Settings + "              WWW.CONTROLANYTHING.COM"
    Form1.Caption = "ProXR Example Software for Visual Basic 6  " + ProXR1.PortName + "  " + str$(ProXR1.BaudRate) + "  WWW.CONTROLANYTHING.COM"
  
      
  
    'Make Sure Relay Bank is Set to 1
'    CC = 0
    Do
        CC = CC + 1
        If CC > 5 Then GoTo EXT
        HScroll4.Value = 1  'Set Relay Bank to 1    Sets the Current Relay Control Bank to 1
        ProXR1.Sleep (100)
        Form1.PutData (254)  'Command Mode
        Form1.PutData (34)   'Ask for Relay Bank
    Loop Until GetData = 1          'Repeat Until Relay Bank is Se]t to 1
'
    'PWM12.Visible = True
    'PWM12.ZOrder 0
    
EXT:
    'Command3_Click  'Update E3C Device Number on Interface

'    GetData 'Clear the Serial Buffer
'    Form1.PutData (254)  'Enter Command Mode
'    Form1.PutData (36)   'Get Refresh Settings Command
'    If GetData = 1 Then
'        Check1.Value = 1
'    Else
'        Check1.Value = 0
'    End If
    'Command8_Click  'Bring Up Specify Bank Commands
'    GetData
    
    Form1.Test2Way_Click
    
    'Do
    '    For X = 0 To 255
    '
    '        HScroll1.Value = X
     '   Next X
    'Loop
    
    
    
    Form1.Command12_Click
        
    
    
    'TimerMultiTest
    
'    Dim i As Long
'    i = GetData
'    Do
'        Do
'            i = GetData
'            DoEvents
'        Loop Until i > 0
'        Debug.Print "DUMPING:"; Asc(i); Timer
'    Loop

End Sub
Private Sub AD1216_Button_Click()
    AD1216.Visible = True
    AD1216.ZOrder 0
    AD1216.Left = Form1.Left + Form1.Width
    AD1216.Top = Form1.Top
    'AD1216_Button.Enabled = False
End Sub
Private Sub ADFeatures_Click()
    AD.Visible = True
    AD.ZOrder 0
    AD.Visible = True
    AD.Left = Form1.Left + Form1.Width
    AD.Top = Form1.Top
    'Form1.ADFeatures.Enabled = False
End Sub
Private Sub AdvancedFeatureButton_Click()
    AdvancedFeatures.Visible = True
    AdvancedFeatures.ZOrder 0
End Sub
Private Sub AllOff_Click()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (29)               'Send Command Turn All Relays Off
    GetData
    Read_All_Click
End Sub
Private Sub AllOn_Click()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (30)               'Send Command Turn All Relays On
    GetData
    Read_All_Click
End Sub
Private Sub Banks_Click(Index As Integer)
    HScroll4.Value = Index + 1
End Sub

Private Sub CCFeatures_Click()
    ScanSwitch.Visible = True
    ScanSwitch.ZOrder 0
    ScanVolt.Visible = True
    ScanVolt.Top = ScanSwitch.Top + ScanSwitch.Height
    ScanVolt.Left = ScanSwitch.Left
    ScanVolt.ZOrder 0
    'CCFeatures.Enabled = False
End Sub
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Form1.PutData (254)  'Enter Command Mode
        Form1.PutData (25)   'Turn ON Refresh Command
    Else
        Form1.PutData (254)  'Enter Command Mode
        Form1.PutData (26)   'Turn Off Refresh Command
    End If
End Sub
Private Sub Command10_Click()
    Form1.PutData (254)  'Enter Command Mode
    Form1.PutData (37)   'Manually Refresh Relays Command
    GetData
End Sub
Private Sub Command11_Click()
    Timers.Visible = True       'Load the Timers Window
    Timers.ZOrder 0             'Bring it the Front of the Screen
End Sub
Public Sub Command12_Click()
    Debug.Print "Setup Interface...."
    
    For ReTest = 1 To 3
        
        Debug.Print "Retry:"; ReTest
        Debug.Print "Getting Device ID"
        Form1.PutData (254)          'Enter Command Mode
        Form1.PutData (246)          'Get Device ID Data
        DID(0).Caption = GetData            'Display Device ID 1
        DID(1).Caption = GetData            'Display Device ID 2
        DID(2).Caption = GetData            'Display Device Year of Manufacture
        DID(3).Caption = GetData            'Display Firmware Build Version
        DID(4).Caption = GetData            'Display E3C Device Number
        Debug.Print "Device ID Complete"
        
    If DID(0).Caption = "1" Then INTER(0).Caption = "ProXR" ': End If
    If DID(0).Caption = "2" Then INTER(0).Caption = "ProXR + 8AD" ': End If
    If DID(0).Caption = "3" Then INTER(0).Caption = "ProXR + POT" ': End If
    If DID(0).Caption = "4" Then INTER(0).Caption = "ProXR + AD1216" ': End If
    If DID(0).Caption = "5" Then INTER(0).Caption = "ProXR + USCS" ': End If
    If DID(0).Caption = "6" Then INTER(0).Caption = "ProXR + UXP" ': End If
    If DID(0).Caption = "7" Then INTER(0).Caption = "ProXR + PWM12" ': End If
    If DID(0).Caption = "8" Then INTER(0).Caption = "ProXR + FPWM"
    If DID(0).Caption = "128" Then INTER(0).Caption = "W + R1/R2"

    
    If DID(1).Caption = "0" Then INTER(1).Caption = "RS-232" ': End If
    If DID(1).Caption = "1" Then INTER(1).Caption = "USB" ': End If
    If DID(1).Caption = "2" Then INTER(1).Caption = "Zigbee" ': End If
    If DID(1).Caption = "3" Then INTER(1).Caption = "Ethernet" ': End If
    If DID(1).Caption = "4" Then INTER(1).Caption = "802.11b Wi-Fi" ': End If
    If DID(1).Caption = "5" Then INTER(1).Caption = "Bluetooth" ': End If
    If DID(1).Caption = "6" Then INTER(1).Caption = "Fiber Optic" ': End If

    
    If INTER(1).Caption = "RS-232" And INTER(0).Caption = "ProXR + AD1216" Then
        Form1.AD1216_Button.Visible = True
        AD1216.Visible = True
        AD1216.ZOrder 0
        AD1216.Left = Form1.Left + Form1.Width
        AD1216.Top = Form1.Top
    End If
    
    If INTER(1).Caption = "USB" And INTER(0).Caption = "ProXR + AD1216" Then
        Form1.AD1216_Button.Visible = True
        AD1216.Visible = True
        AD1216.ZOrder 0
        AD1216.Left = Form1.Left + Form1.Width
        AD1216.Top = Form1.Top
    End If

    If INTER(1).Caption = "USB" And INTER(0).Caption = "ProXR + PWM12" Then
        PWM12_Button.Visible = True
        PWM12.Visible = True
        PWM12.ZOrder 0
        PWM12.Left = Form1.Left + Form1.Width
        PWM12.Top = Form1.Top
    End If

    If INTER(0).Caption = "ProXR + FPWM" Then
        'PWM12_Button.Visible = True
        PWM.Visible = True
        PWM.ZOrder 0
        PWM.Left = Form1.Left + Form1.Width
        PWM.Top = Form1.Top
    End If

    If INTER(1).Caption = "Fiber Optic" Then
        FiberOpticOptions.Visible = True
        FiberOpticOptions.ZOrder 0
        FiberOpticOptions.Left = Form1.Left + Form1.Width
        'FiberOpticOptions.Top = Form1.Top
    End If
    
    If INTER(1).Caption = "RS-232" And INTER(0).Caption = "ProXR" Then
        Label21.Visible = False
    End If

    If INTER(1).Caption = "RS-232" And INTER(0).Caption <> "ProXR" Then
        RS232Special.Visible = True
        RS232Special.Top = Form1.Top + Form1.Height
        RS232Special.Left = Form1.Left
        RS232Baud.Visible = True
        RS232Baud.Enabled = False
    End If
    
    If INTER(1).Caption <> "RS-232" And INTER(1).Caption <> "Zigbee" Then
        Frame20.Visible = False     'Turn Off E3C Programming Option
        Frame14.Visible = False     'Turn Off E3C Options
        AdvancedFeatures.Label7.Visible = False 'Turn Off Baud Change Note
        AdvancedFeatures.ChangeBaud.Visible = False 'Turn Off Baud Change Button
        AdvancedFeatures.Label9.Visible = False
        AdvancedFeatures.Label8.Visible = False
        AdvancedTimers.Label11.Visible = False
        DID(4).Visible = False
        Label20.Visible = False
        RS232Special.Frame4.Visible = False
        
        For x = 1 To 6
            BaudSet(x).Enabled = False
        Next x
    End If
    
   ' End

    If INTER(1).Caption = "USB" Then
        USBFeatures.Visible = True
        USB_Features.Visible = True
        USB_Features.Left = Form1.Left + Form1.Width
        USB_Features.Top = Form1.Top + AD.Height
        'AD.Left = Form1.Left + Form1.Width
        'AD.Top = Form1.Top
    End If

    If INTER(1).Caption = "Zigbee" Then
        DID(4).Visible = False
        Label20.Visible = False
        ZigbeeFeatures.Visible = True
        Zigbee.Visible = True
        Zigbee.Left = Form1.Left + Form1.Width
        Zigbee.Top = Form1.Top + AD.Height
        'AD.Left = Form1.Left + Form1.Width
        'AD.Top = Form1.Top
    End If
    
    If INTER(0).Caption = "ProXR + 8AD" Then Form1.ADFeatures.Visible = True
    If INTER(0).Caption = "ProXR + POT" Then Form1.POTFeatures.Visible = True
    If INTER(0).Caption = "ProXR + AD1216" Then Form1.AD1216_Button.Visible = True
    
    If INTER(0).Caption = "ProXR + POT" Or INTER(0).Caption = "ProXR + UXP" Then
        POT.Visible = True
        POT.Left = Form1.Left + Form1.Width
        POT.Top = Form1.Top
        Form1.Recover.Visible = False
    End If
    
    If DID(0).Caption = "2" Then 'And DID(1).Caption = "2" Then
        AD.Visible = True
        AD.Left = Form1.Left + Form1.Width
        AD.Top = Form1.Top
        Form1.Recover.Visible = False
    End If
    
    If INTER(0).Caption = "ProXR + UXP" Then
        PWM.Visible = True
        Form1.CCFeatures.Visible = True
        Form1.AD1216_Button.Visible = True
        AD1216.Visible = True
        Form1.POTFeatures.Visible = True
        ScanSwitch.Visible = True
        ScanVolt.Visible = True
        ScanVolt.Top = ScanSwitch.Top + ScanSwitch.Height
        ScanVolt.Left = ScanSwitch.Left
        If DID(3).Caption = "1" Then
            POT.Read.Visible = False
            POT.Store.Visible = False
        End If
    End If
    
    Next ReTest
    
    Debug.Print "EXIT Setup Interface"
        
End Sub

Private Sub Command13_Click()
    MISC_Command_Testing
End Sub

Private Sub Command2_Click()
    BaudSet(5).Value = True
    Form1.PutData (254)          'Enter Command Mode
    Form1.PutData (255)          'Store E3C Device Number (Controller Must Be in Configuration Mode for this to work)
    Form1.PutData (HScroll2)     'Set Value of E3C Device Number
    GetData
    Command3_Click                      'Read Stored Settings
End Sub
Public Sub Command3_Click()
    Form1.PutData (254)  'Enter Command Mode
    Form1.PutData (247)  'Get E3C Device Number
    HScroll3.Value = GetData    'Update Interface with returned value
    HScroll2.Value = HScroll3   'Update Slider on Lower Part of Interface
End Sub
Private Sub Command4_Click()
    Form1.PutData (254)  'Enter Command Mode
    Form1.PutData (252)  'Enable Selected Device ONLY (All Others Disabled)
    Form1.PutData (HScroll3.Value)   'Device to Enable
End Sub

Private Sub Command5_Click()
    Form1.PutData (254)  'Enter Command Mode
    Form1.PutData (248)  'Enable All Devices
End Sub
Private Sub Command6_Click()
    Form1.PutData (254)  'Enter Command Mode
    Form1.PutData (249)  'Disable All Devices
End Sub

Private Sub Command7_Click()
    Form1.PutData (254)  'Enter Command Mode
    Form1.PutData (250)  'Enable Selected Device (Other Devices will Not Be Affected)
    Form1.PutData (HScroll3.Value)   'Device to Enable
End Sub

Public Sub Command8_Click()
    Form2.Visible = True
    Form2.ZOrder 0
    Form2.BANK = HScroll4.Value
End Sub

Public Sub Command9_Click()
    Form1.PutData (254)  'Enter Command Mode
    Form1.PutData (35)   'Store Refresh Settings Command as Power Up Default
    GetData
End Sub


Public Sub dela2()
    For tt = 0 To 20000
        DoEvents
    Next tt
End Sub

Public Sub TimerMultiTest()

For TimerSlider = 0 To 15
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (70 + TimerSlider) 'Activate Relay Timer (Index) for Duration
    Form1.PutData (0)   'Hours
    Form1.PutData (0) 'Minutes
    Form1.PutData ((5 * TimerSlider)) 'Seconds
    Form1.PutData (TimerSlider + 8) 'Relay
    Form1.GetData

Next TimerSlider
    
End Sub
Public Sub MISC_Command_Testing()
    Debug.Print "254,33 Testing 2-Way Communications"
    OUT 254
    OUT 33
    GetData
    Debug.Print "254,34 Send Currently Selected Relay Bank"
    OUT 254
    OUT 34
    GetData
   Debug.Print "254,36 Automatic Refresh Setting"
    OUT 254
    OUT 36
    GetData
   Debug.Print "254,28 Turn Off Reporting Mode SETTING STORED"
    OUT 254
    OUT 28
    GetData
'   Debug.Print "254,27 Turn ON Reporting Mode SETTING STORED"
'    OUT 254
'    OUT 27
'    GetData
    
End Sub
Public Sub OUT(i As Integer)
    Form1.PutData (i)
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If INTER(1).Caption = "RS-232" And INTER(0).Caption = "ProXR" Then
    Else
        Form1.Test2Way_Click
        If Form1.Label21.Caption = "CONFIG MODE" Then
            MsgBox ("This device is currently in Configuration Mode.  Before exiting this program, please change the Program/Run Jumper to Run mode.")
        End If
    End If
    ProXR1.ClosePort
    
    AD1216.Visible = False
    PWM12.Visible = False
    AdvancedFeatures.Visible = False
    FiberOpticOptions.Visible = False
    AdvancedTimers.Visible = False
    Form2.Visible = False
    Form3.Visible = False
    
    Unload frmLog
    ProXR1.ClosePort
    If Not parentForm Is Nothing Then
        parentForm.Show
    End If
    
    
    
End Sub

Private Sub GetBank_Click()
  '  GetData
    GetBank.Enabled = False
    Debug.Print "----    "
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (43)               'Get Power-Up Status of Relays for Selected Bank
    'Form1.PutData (0)
    'GetData
    If HScroll4 > 0 Then
        temp = GetData
        Banks(HScroll4 - 1).BackColor = &HFF&
        Banks(HScroll4 - 1).Caption = str$(HScroll4) + ":" + str$(temp) + ":" + BIN$(temp)
        'Debug.Print "Extra:"; GetData
    Else
        For N = 0 To 31
            temp = GetData
            Banks(N).BackColor = &HFF&
            Banks(N).Caption = str$(N + 1) + ":" + str$(temp) + ":" + BIN$(temp)
        Next N
    End If
    GetBank.Enabled = True
End Sub
Public Sub HScroll1_Change()
    Label6.Caption = HScroll1.Value
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (40)               'Set Status of All Relays Command
    Form1.PutData (HScroll1.Value)   'Set Status of All Relays Command
    GetData
    Read_All_Click
End Sub
Private Sub HScroll1_Scroll()
    Label6.Caption = HScroll1.Value
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (40)               'Set Status of All Relays Command
    Form1.PutData (HScroll1.Value)   'Set Status of All Relays Command
   GetData
    Read_All_Click
End Sub
Private Sub HScroll2_Change()
    Label1(10).Caption = HScroll2.Value
End Sub
Private Sub HScroll2_Scroll()
    Label1(10).Caption = HScroll2.Value
End Sub

Private Sub HScroll3_Change()
    Label1(11).Caption = HScroll3.Value
End Sub
Private Sub HScroll3_Scroll()
    Label1(11).Caption = HScroll3.Value
End Sub

Private Sub HScroll4_Change()
    Debug.Print "HScroll4 Change..."
    HScroll4.Enabled = False
    Label1(8).Caption = HScroll4.Value
    If HScroll4.Value = 0 Then Label1(8).Caption = "0 (All Banks)"
    Form1.PutData (254)
    Form1.PutData (49)
    Form1.PutData (HScroll4.Value)
    GetData
    'Do Not Read the Status of All Banks
    If HScroll4.Value <> 0 Then Read_All_Click
    If HScroll4.Value > 0 Then
        For N = 0 To 31
            Banks(N).BackColor = &H80000005
        Next N
        Banks(HScroll4.Value - 1).BackColor = &HFFC0C0
        'Read_All_Click
    Else
        For N = 0 To 31
            Banks(N).BackColor = &HFFC0C0
        Next N
        Read_All_Click
    End If
    HScroll4.Enabled = True
End Sub
Private Sub HScroll4_Scroll()
Debug.Print "HScroll4 Scroll..."
    Label1(8).Caption = HScroll4.Value
    If HScroll4.Value = 0 Then Label1(8).Caption = "0 (All Banks)"
    Form1.PutData (254)
    Form1.PutData (49)
    Form1.PutData (HScroll4.Value)
    GetData
    'Do Not Read the Status of All Banks
    If HScroll4.Value <> 0 Then Read_All_Click
    If HScroll4.Value > 0 Then
        For N = 0 To 31
            Banks(N).BackColor = &H80000005
        Next N
        Banks(HScroll4.Value - 1).BackColor = &HFFC0C0
    Else
        HScroll4.Enabled = False
        For N = 0 To 31
            Banks(N).BackColor = &HFFC0C0
        Next N
        Read_All_Click
        HScroll4.Enabled = True
    End If
    If HScroll4.Value <> 0 Then Read_All_Click

End Sub
Private Sub HScroll5_Change()
    Label8.Caption = HScroll5.Value
    Form1.PutData (254)
    Form1.PutData (47)
    Form1.PutData (HScroll5.Value)
    GetData
End Sub
Private Sub HScroll5_Scroll()
    Label8.Caption = HScroll5.Value
    Form1.PutData (254)
    Form1.PutData (47)
    Form1.PutData (HScroll5.Value)
    GetData
End Sub
Private Sub HScroll6_Change()
    Label11.Caption = HScroll6.Value
    Form1.PutData (254)
    Form1.PutData (48)
    Form1.PutData (HScroll6.Value)
    GetData
End Sub
Private Sub HScroll6_Scroll()
    Label11.Caption = HScroll6.Value
    Form1.PutData (254)
    Form1.PutData (48)
    Form1.PutData (HScroll6.Value)
    GetData
End Sub
Private Sub HScroll7_Change()
    Label14.Caption = HScroll7.Value
    Form1.PutData (254)
    Form1.PutData (46)
    Form1.PutData (HScroll7.Value)
    GetData
End Sub
Private Sub HScroll7_Scroll()
    Label14.Caption = HScroll7.Value
    Form1.PutData (254)
    Form1.PutData (46)
    Form1.PutData (HScroll7.Value)
    GetData
End Sub

Private Sub Invert_Click()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (31)               'Send Command Invert Relay Status
    GetData
    Read_All_Click
End Sub

Private Sub Label21_Click()
    Form1.Test2Way_Click
End Sub

' Private Sub MSComm1_OnComm()

'End Sub

Private Sub POTFeatures_Click()
        'POTFeatures.Enabled = False
        POT.Visible = True
        POT.Left = Form1.Left + Form1.Width
        POT.Top = Form1.Top
        Form1.Recover.Visible = False
End Sub

Private Sub ProXR1_OnDataReceived(ByVal data As Integer)
    frmLog.OnDataReceived data
End Sub

Private Sub ProXR1_OnDataSent(ByVal data As Integer)
    frmLog.OnDataSent data
End Sub

Private Sub PWM12_Button_Click()
    PWM12.Visible = True
    PWM12.ZOrder 0
End Sub

Public Sub Read_All_Click()
    Read_All.Enabled = False
    If HScroll4.Value = 0 Then Refresh_All
    If HScroll4.Value = 0 Then
        Read_All.Enabled = True
        Exit Sub
    End If
    For dela = 0 To 500
    DoEvents
    Next dela
    
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (24)               'Send Command to Read the Status of All Relays
    temp = GetData                          'Read Status from Relay Board
        
    'Set All to OFF
    For nn = 0 To 7
        Label1(nn).Caption = "OFF"
        RELAY(nn).Caption = "OFF"
    Next nn
    
    'Determine Which Relays are ON
    If (temp And 1) = 1 Then Label1(0).Caption = "ON": RELAY(0).Caption = "ON"
    If (temp And 2) = 2 Then Label1(1).Caption = "ON": RELAY(1).Caption = "ON"
    If (temp And 4) = 4 Then Label1(2).Caption = "ON": RELAY(2).Caption = "ON"
    If (temp And 8) = 8 Then Label1(3).Caption = "ON": RELAY(3).Caption = "ON"
    If (temp And 16) = 16 Then Label1(4).Caption = "ON": RELAY(4).Caption = "ON"
    If (temp And 32) = 32 Then Label1(5).Caption = "ON": RELAY(5).Caption = "ON"
    If (temp And 64) = 64 Then Label1(6).Caption = "ON": RELAY(6).Caption = "ON"
    If (temp And 128) = 128 Then Label1(7).Caption = "ON": RELAY(7).Caption = "ON"
    
    Banks(HScroll4 - 1).Caption = str$(HScroll4) + ":" + str$(temp) + ":" + BIN$(temp)
    Read_All.Enabled = True
End Sub
Public Sub Refresh_All()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (24)               'Send Command to Read the Status of All Relays
    Debug.Print "REFRESH ALL ROUTINE...."
    For N = 0 To 31
        Debug.Print N;
        temp = GetData                          'Read Status from Relay Board
        Banks(N).Caption = str$(N + 1) + ":" + str$(temp) + ":" + BIN$(temp)
    Next N
End Sub
Private Sub Read_Relay_Click(Index As Integer)
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (16 + Index)       'Send Command to Read the Status of a Relay (1-8)
    temp = GetData                          'Read Status from Relay Board
    If temp = 0 Then
        Label1(Index).Caption = "OFF"
    Else
        Label1(Index).Caption = "ON"
    End If
End Sub

Private Sub Recover_Click()
    
    Form1.PutData (254)  'Enter Command Mode
    Form1.PutData (248)  'Enable All Devices
    
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (147) 'Attempt Recovery Command

End Sub

Private Sub RELAY_Click(Index As Integer)
    RELAY(Index).Enabled = False
    If RELAY(Index).Caption = "ON" Then
        RELAY(Index).Caption = "OFF"
        Form1.PutData (254)          'Enter Command Mode
        Form1.PutData (Index)        'Turn Relay Off
    Else
        RELAY(Index).Caption = "ON"
        Form1.PutData (254)          'Enter Command Mode
        Form1.PutData (Index + 8)    'Turn Relay Off
    End If
    GetData
    Read_All_Click
    RELAY(Index).Enabled = True
End Sub
Private Sub Reporting_Click()
    If Reporting.Caption = "OFF" Then
        Reporting.Caption = "ON"
        Form1.PutData (254)          'Enter Command Mode
        Form1.PutData (27)           'Turn Reporting Mode ON
    Else
        Reporting.Caption = "OFF"
        Form1.PutData (254)          'Enter Command Mode
        Form1.PutData (28)           'Turn Reporting Mode OFF
    End If
End Sub
Private Sub Reverse_Click()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (32)               'Send Command to Reverse Relay Pattern
    GetData
    Read_All_Click
End Sub

Private Sub StoreBank_Click()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (42)               'Store Current Memory Bank as Power Up Defaults
    GetData
    If HScroll4 <> 0 Then
        Banks(HScroll4 - 1).BackColor = &HFF&
    Else
        For N = 0 To 31
            Banks(N).BackColor = &HFF&
        Next N
    End If
End Sub
Public Sub Test2Way_Click()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (33)               'Send Command to Test 2-Way Communication
    temp = GetData
    If temp = 85 Or temp = 86 Then          'Read Status from Relay Board
        T2Way.Caption = "PASS"              '2-Way Communication Test Passed
        If temp = 85 Then
            Label21.BackColor = &HC000&
            Label21.Caption = " RUN  MODE"
        End If
        If temp = 86 Then
            Label21.BackColor = &HFF&
            Label21.Caption = "CONFIG MODE"
        End If
    Else
        T2Way.Caption = "FAIL"
    End If
End Sub
Public Function BIN$(temp)
    BIN$ = ""
    'Determine Which Relays are ON
    If (temp And 128) = 128 Then BIN$ = BIN$ + "1" Else BIN$ = BIN$ + "0"
    If (temp And 64) = 64 Then BIN$ = BIN$ + "1" Else BIN$ = BIN$ + "0"
    If (temp And 32) = 32 Then BIN$ = BIN$ + "1" Else BIN$ = BIN$ + "0"
    If (temp And 16) = 16 Then BIN$ = BIN$ + "1" Else BIN$ = BIN$ + "0"
    If (temp And 8) = 8 Then BIN$ = BIN$ + "1" Else BIN$ = BIN$ + "0"
    If (temp And 4) = 4 Then BIN$ = BIN$ + "1" Else BIN$ = BIN$ + "0"
    If (temp And 2) = 2 Then BIN$ = BIN$ + "1" Else BIN$ = BIN$ + "0"
    If (temp And 1) = 1 Then BIN$ = BIN$ + "1" Else BIN$ = BIN$ + "0"
End Function
Private Sub TestCycle_Click()
    If TestCycle.Caption = "Relay Test Cycle" Then
        TestCycle.Caption = "Cancel Test Cycle"
    Else
        TestCycle.Caption = "Relay Test Cycle"
    End If

    RELAYs = 1
    BANK = 1
    While TestCycle.Caption <> "Relay Test Cycle"
       ' For x = 1 To 50000
       '     ProXR1.ClearBuffer
       '     DoEvents
       ' Next x
    'ProXR1.Sleep (1000)
    ProXR1.ClearBuffer
        Form1.PutData (254)              'Enter Command Mode
        Form1.PutData (140)               'Set Status of All Relays Command
        Form1.PutData (RELAYs)   'Set Status of All Relays Command
        Form1.PutData (BANK)
        'GetData
        Debug.Print "--------------------> " & str(ProXR1.GetData2(3000))
        RELAYs = RELAYs * 2
        If RELAYs > 128 Then
            RELAYs = 1
            BANK = BANK + 1
            If BANK > 2 Then
                BANK = 1
            End If
        End If
    Wend
    
End Sub

Private Sub TimerCalibrationButton_Click()
    AdvancedTimers.Visible = True
    AdvancedTimers.ZOrder 0
End Sub
Public Sub BaudScan()
    Debug.Print "Baud Scan In Progress..."
    GetData
    'ProXR1.ClearBuffer
    For again = 1 To 3
    For test = 7 To 1 Step -1               'Test Baud Rates Backward for Optimum Speed
        'ProXR1.ClearBuffer
        'GetData
        Debug.Print "Testing Baud"; test
        BaudSet(test).Value = True 'Test Baud Rates
        BaudSet_Click (test)
        Form1.PutData (254)              'Enter Command Mode
        Form1.PutData (33)               'Send Command to Test 2-Way Communication
        temp = GetData
        If (temp = 85) Or (temp = 86) Then Exit Sub
        For dela = 0 To 5000 '5000
           DoEvents
        Next dela
    Next test
    GetData
    Next again
    Debug.Print "Controller Not Found..."
    MsgBox ("Controller Not Found")
    'End
End Sub
Private Sub USBFeatures_Click()
        USB_Features.Visible = True
        USB_Features.ZOrder 0
        USB_Features.Left = Form1.Left + Form1.Width
        USB_Features.Top = Form1.Top + AD.Height
        'USBFeatures.Enabled = False
End Sub
Private Sub ZigbeeFeatures_Click()
    Zigbee.Visible = True
    Zigbee.ZOrder 0
    Form1.ZigbeeFeatures.Enabled = False
    
        ZigbeeFeatures.Visible = True
        Zigbee.Visible = True
        Zigbee.Left = Form1.Left + Form1.Width
        Zigbee.Top = Form1.Top + AD.Height
End Sub
Public Sub PutData(DAT As Integer)
    Debug.Print "OUT:"; DAT
    'Form1.MSComm1.Output = Chr$(DAT)
    ProXR1.SendData DAT
End Sub
Public Function GetData()
        GetData = GetDataQuiet
        If GetData < 0 Then GetData = 0
        Debug.Print "IN:";
        Debug.Print GetData
End Function
Public Function GetDataQuiet()
        Dim a As Long
        a = ProXR1.GetData
        GetDataQuiet = a
        If a = -1 Then
            Exit Function
        End If
End Function
Public Function GetDataQuiet2()
        GetDataQuiet2 = ProXR1.GetData
End Function
Private Sub Command1_Click()
    Dim i As Long
    i = ProXR1.GetData
    While i >= 0
        DUMP = Asc(i) 'Read Data Byte from the Serial Port
        DoEvents
        Debug.Print "DUMPING:"; DUMP
    i = ProXR1.GetData
    Wend
End Sub
Public Sub BaudSet_Click(Index As Integer)
    If Index = 0 Then ProXR1.BaudRate = br_38400
    If Index = 1 Then ProXR1.BaudRate = br_2400
    If Index = 2 Then ProXR1.BaudRate = br_4800
    If Index = 3 Then ProXR1.BaudRate = br_9600
    If Index = 4 Then ProXR1.BaudRate = br_19200
    If Index = 5 Then ProXR1.BaudRate = br_38400
    If Index = 6 Then ProXR1.BaudRate = br_57600
    If Index = 7 Then ProXR1.BaudRate = br_115200
    ProXR1.OpenPort
    Form1.Caption = "ProXR Example Software for Visual Basic 6  " + ProXR1.PortName + "  " + str$(ProXR1.BaudRate) + "  WWW.CONTROLANYTHING.COM"
    
End Sub
