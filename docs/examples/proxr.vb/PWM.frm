VERSION 5.00
Begin VB.Form PWM 
   Caption         =   "PWM Demo Program"
   ClientHeight    =   11250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18900
   LinkTopic       =   "Form3"
   ScaleHeight     =   11250
   ScaleWidth      =   18900
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox BITCH 
      Caption         =   "Bit 0"
      Height          =   255
      Index           =   0
      Left            =   12360
      TabIndex        =   241
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   8415
      Left            =   12120
      TabIndex        =   240
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 15"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   256
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 14"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   255
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 13"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   254
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 12"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   253
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 11"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   252
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 10"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   251
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 9"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   250
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 8"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   249
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 7"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   248
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 6"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   247
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 5"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   246
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 4"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   245
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 3"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   244
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 2"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   243
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox BITCH 
         Caption         =   "Bit 1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   242
         Top             =   2760
         Width           =   1215
      End
   End
   Begin VB.CommandButton ClearBuffer 
      Caption         =   "Clear Serial Buffer"
      Height          =   495
      Left            =   240
      TabIndex        =   205
      Top             =   10080
      Width           =   6855
   End
   Begin VB.CommandButton ReadConfig 
      Caption         =   "Read Configuration Settings from Controller"
      Height          =   495
      Left            =   120
      TabIndex        =   204
      Top             =   120
      Width           =   11895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Set the Frequency of Each Channel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   5520
      TabIndex        =   102
      Top             =   720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Frame Frame4 
         Caption         =   "~Frequency"
         Height          =   8295
         Left            =   5160
         TabIndex        =   206
         Top             =   120
         Width           =   1215
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   239
            Top             =   240
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   238
            Top             =   480
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   237
            Top             =   720
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   236
            Top             =   960
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   235
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   234
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   233
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   232
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   231
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   230
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   229
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   228
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   227
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   226
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   225
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   224
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   223
            Top             =   4080
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   222
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   221
            Top             =   4560
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   220
            Top             =   4800
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   219
            Top             =   5040
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   218
            Top             =   5280
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   217
            Top             =   5520
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   216
            Top             =   5760
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   215
            Top             =   6000
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   214
            Top             =   6240
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   213
            Top             =   6480
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   212
            Top             =   6720
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   211
            Top             =   6960
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   210
            Top             =   7200
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   209
            Top             =   7440
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   208
            Top             =   7680
            Width           =   975
         End
         Begin VB.Label KHZ 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   207
            Top             =   7920
            Width           =   975
         End
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   12
         Left            =   1320
         Max             =   1023
         TabIndex        =   137
         Top             =   3240
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   11
         Left            =   1320
         Max             =   1023
         TabIndex        =   136
         Top             =   3000
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   10
         Left            =   1320
         Max             =   1023
         TabIndex        =   135
         Top             =   2760
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   9
         Left            =   1320
         Max             =   1023
         TabIndex        =   134
         Top             =   2520
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   0
         Left            =   1320
         Max             =   1023
         TabIndex        =   133
         Top             =   360
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   1
         Left            =   1320
         Max             =   1023
         TabIndex        =   132
         Top             =   600
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   2
         Left            =   1320
         Max             =   1023
         TabIndex        =   131
         Top             =   840
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   3
         Left            =   1320
         Max             =   1023
         TabIndex        =   130
         Top             =   1080
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   4
         Left            =   1320
         Max             =   1023
         TabIndex        =   129
         Top             =   1320
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   5
         Left            =   1320
         Max             =   1023
         TabIndex        =   128
         Top             =   1560
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   6
         Left            =   1320
         Max             =   1023
         TabIndex        =   127
         Top             =   1800
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   7
         Left            =   1320
         Max             =   1023
         TabIndex        =   126
         Top             =   2040
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   8
         Left            =   1320
         Max             =   1023
         TabIndex        =   125
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Frame Frame3 
         Caption         =   "Powerup Default Settings"
         Height          =   735
         Left            =   120
         TabIndex        =   123
         Top             =   8400
         Width           =   6255
         Begin VB.CommandButton Command1 
            Caption         =   "Store Current Settings as Powerup Default"
            Height          =   375
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   13
         Left            =   1320
         Max             =   1023
         TabIndex        =   122
         Top             =   3480
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   14
         Left            =   1320
         Max             =   1023
         TabIndex        =   121
         Top             =   3720
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   15
         Left            =   1320
         Max             =   1023
         TabIndex        =   120
         Top             =   3960
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   16
         Left            =   1320
         Max             =   1023
         TabIndex        =   119
         Top             =   4200
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   17
         Left            =   1320
         Max             =   1023
         TabIndex        =   118
         Top             =   4440
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   18
         Left            =   1320
         Max             =   1023
         TabIndex        =   117
         Top             =   4680
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   19
         Left            =   1320
         Max             =   1023
         TabIndex        =   116
         Top             =   4920
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   20
         Left            =   1320
         Max             =   1023
         TabIndex        =   115
         Top             =   5160
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   21
         Left            =   1320
         Max             =   1023
         TabIndex        =   114
         Top             =   5400
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   22
         Left            =   1320
         Max             =   1023
         TabIndex        =   113
         Top             =   5640
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   23
         Left            =   1320
         Max             =   1023
         TabIndex        =   112
         Top             =   5880
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   24
         Left            =   1320
         Max             =   1023
         TabIndex        =   111
         Top             =   6120
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   25
         Left            =   1320
         Max             =   1023
         TabIndex        =   110
         Top             =   6360
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   26
         Left            =   1320
         Max             =   1023
         TabIndex        =   109
         Top             =   6600
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   27
         Left            =   1320
         Max             =   1023
         TabIndex        =   108
         Top             =   6840
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   28
         Left            =   1320
         Max             =   1023
         TabIndex        =   107
         Top             =   7080
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   29
         Left            =   1320
         Max             =   1023
         TabIndex        =   106
         Top             =   7320
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   30
         Left            =   1320
         Max             =   1023
         TabIndex        =   105
         Top             =   7560
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   31
         Left            =   1320
         Max             =   1023
         TabIndex        =   104
         Top             =   7800
         Width           =   3015
      End
      Begin VB.HScrollBar FRQ 
         Height          =   255
         Index           =   32
         Left            =   1320
         Max             =   1023
         TabIndex        =   103
         Top             =   8040
         Width           =   3015
      End
      Begin VB.Label Label66 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 12"
         Height          =   255
         Left            =   120
         TabIndex        =   203
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 11"
         Height          =   255
         Left            =   120
         TabIndex        =   202
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 10"
         Height          =   255
         Left            =   120
         TabIndex        =   201
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 9"
         Height          =   255
         Left            =   120
         TabIndex        =   200
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   4440
         TabIndex        =   199
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   4440
         TabIndex        =   198
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   4440
         TabIndex        =   197
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   196
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label62 
         Caption         =   "Set All Channels"
         Height          =   255
         Left            =   120
         TabIndex        =   195
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 1"
         Height          =   255
         Left            =   120
         TabIndex        =   194
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 2"
         Height          =   255
         Left            =   120
         TabIndex        =   193
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 3"
         Height          =   255
         Left            =   120
         TabIndex        =   192
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 4"
         Height          =   255
         Left            =   120
         TabIndex        =   191
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 5"
         Height          =   255
         Left            =   120
         TabIndex        =   190
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 6"
         Height          =   255
         Left            =   120
         TabIndex        =   189
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 7"
         Height          =   255
         Left            =   120
         TabIndex        =   188
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 8"
         Height          =   255
         Left            =   120
         TabIndex        =   187
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   186
         Top             =   360
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   185
         Top             =   600
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   184
         Top             =   840
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   183
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   182
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   181
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   180
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   179
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   178
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   4440
         TabIndex        =   177
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   4440
         TabIndex        =   176
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   4440
         TabIndex        =   175
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   4440
         TabIndex        =   174
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   4440
         TabIndex        =   173
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   4440
         TabIndex        =   172
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   4440
         TabIndex        =   171
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   4440
         TabIndex        =   170
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   4440
         TabIndex        =   169
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   4440
         TabIndex        =   168
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   4440
         TabIndex        =   167
         Top             =   5880
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   4440
         TabIndex        =   166
         Top             =   6120
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   4440
         TabIndex        =   165
         Top             =   6360
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   4440
         TabIndex        =   164
         Top             =   6600
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   4440
         TabIndex        =   163
         Top             =   6840
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   4440
         TabIndex        =   162
         Top             =   7080
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   4440
         TabIndex        =   161
         Top             =   7320
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   4440
         TabIndex        =   160
         Top             =   7560
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   4440
         TabIndex        =   159
         Top             =   7800
         Width           =   615
      End
      Begin VB.Label FRQValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   4440
         TabIndex        =   158
         Top             =   8040
         Width           =   615
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 24"
         Height          =   255
         Left            =   120
         TabIndex        =   157
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 23"
         Height          =   255
         Left            =   120
         TabIndex        =   156
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 22"
         Height          =   255
         Left            =   120
         TabIndex        =   155
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 21"
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 13"
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 14"
         Height          =   255
         Left            =   120
         TabIndex        =   152
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 15"
         Height          =   255
         Left            =   120
         TabIndex        =   151
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 16"
         Height          =   255
         Left            =   120
         TabIndex        =   150
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 17"
         Height          =   255
         Left            =   120
         TabIndex        =   149
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 18"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 19"
         Height          =   255
         Left            =   120
         TabIndex        =   147
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 20"
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 25"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 31"
         Height          =   255
         Left            =   120
         TabIndex        =   144
         Top             =   7800
         Width           =   1095
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 30"
         Height          =   255
         Left            =   120
         TabIndex        =   143
         Top             =   7560
         Width           =   1095
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 29"
         Height          =   255
         Left            =   120
         TabIndex        =   142
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 28"
         Height          =   255
         Left            =   120
         TabIndex        =   141
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 26"
         Height          =   255
         Left            =   120
         TabIndex        =   140
         Top             =   6600
         Width           =   1095
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 27"
         Height          =   255
         Left            =   120
         TabIndex        =   139
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 32"
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   8040
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set the Duty Cycle of Each Output Channel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   5295
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   32
         Left            =   1320
         Max             =   255
         TabIndex        =   61
         Top             =   8040
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   31
         Left            =   1320
         Max             =   255
         TabIndex        =   60
         Top             =   7800
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   30
         Left            =   1320
         Max             =   255
         TabIndex        =   59
         Top             =   7560
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   29
         Left            =   1320
         Max             =   255
         TabIndex        =   58
         Top             =   7320
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   28
         Left            =   1320
         Max             =   255
         TabIndex        =   57
         Top             =   7080
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   27
         Left            =   1320
         Max             =   255
         TabIndex        =   56
         Top             =   6840
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   26
         Left            =   1320
         Max             =   255
         TabIndex        =   55
         Top             =   6600
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   25
         Left            =   1320
         Max             =   255
         TabIndex        =   54
         Top             =   6360
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   24
         Left            =   1320
         Max             =   255
         TabIndex        =   53
         Top             =   6120
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   23
         Left            =   1320
         Max             =   255
         TabIndex        =   52
         Top             =   5880
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   22
         Left            =   1320
         Max             =   255
         TabIndex        =   51
         Top             =   5640
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   21
         Left            =   1320
         Max             =   255
         TabIndex        =   50
         Top             =   5400
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   20
         Left            =   1320
         Max             =   255
         TabIndex        =   49
         Top             =   5160
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   19
         Left            =   1320
         Max             =   255
         TabIndex        =   48
         Top             =   4920
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   18
         Left            =   1320
         Max             =   255
         TabIndex        =   47
         Top             =   4680
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   17
         Left            =   1320
         Max             =   255
         TabIndex        =   46
         Top             =   4440
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   16
         Left            =   1320
         Max             =   255
         TabIndex        =   45
         Top             =   4200
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   15
         Left            =   1320
         Max             =   255
         TabIndex        =   44
         Top             =   3960
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   14
         Left            =   1320
         Max             =   255
         TabIndex        =   43
         Top             =   3720
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   13
         Left            =   1320
         Max             =   255
         TabIndex        =   42
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Frame Frame21 
         Caption         =   "Powerup Default Settings"
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   8400
         Width           =   5055
         Begin VB.CommandButton StorePattern 
            Caption         =   "Store Current Settings as Powerup Default"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   8
         Left            =   1320
         Max             =   255
         TabIndex        =   13
         Top             =   2280
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   7
         Left            =   1320
         Max             =   255
         TabIndex        =   12
         Top             =   2040
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   6
         Left            =   1320
         Max             =   255
         TabIndex        =   11
         Top             =   1800
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   5
         Left            =   1320
         Max             =   255
         TabIndex        =   10
         Top             =   1560
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   4
         Left            =   1320
         Max             =   255
         TabIndex        =   9
         Top             =   1320
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   3
         Left            =   1320
         Max             =   255
         TabIndex        =   8
         Top             =   1080
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   2
         Left            =   1320
         Max             =   255
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   1
         Left            =   1320
         Max             =   255
         TabIndex        =   6
         Top             =   600
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   0
         Left            =   1320
         Max             =   255
         TabIndex        =   5
         Top             =   360
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   9
         Left            =   1320
         Max             =   255
         TabIndex        =   4
         Top             =   2520
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   10
         Left            =   1320
         Max             =   255
         TabIndex        =   3
         Top             =   2760
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   11
         Left            =   1320
         Max             =   255
         TabIndex        =   2
         Top             =   3000
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   12
         Left            =   1320
         Max             =   255
         TabIndex        =   1
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 32"
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   8040
         Width           =   1095
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 27"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 26"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   6600
         Width           =   1095
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 28"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 29"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 30"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   7560
         Width           =   1095
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 31"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   7800
         Width           =   1095
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 25"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 20"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 19"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 18"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 17"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 16"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 15"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 14"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 13"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 21"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 22"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 23"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 24"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   4440
         TabIndex        =   81
         Top             =   8040
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   4440
         TabIndex        =   80
         Top             =   7800
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   4440
         TabIndex        =   79
         Top             =   7560
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   4440
         TabIndex        =   78
         Top             =   7320
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   4440
         TabIndex        =   77
         Top             =   7080
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   4440
         TabIndex        =   76
         Top             =   6840
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   4440
         TabIndex        =   75
         Top             =   6600
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   4440
         TabIndex        =   74
         Top             =   6360
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   4440
         TabIndex        =   73
         Top             =   6120
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   4440
         TabIndex        =   72
         Top             =   5880
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   22
         Left            =   4440
         TabIndex        =   71
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   4440
         TabIndex        =   70
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   4440
         TabIndex        =   69
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   4440
         TabIndex        =   68
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   4440
         TabIndex        =   67
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   4440
         TabIndex        =   66
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   4440
         TabIndex        =   65
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   4440
         TabIndex        =   64
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   4440
         TabIndex        =   63
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   4440
         TabIndex        =   62
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   41
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   40
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   39
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   38
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   37
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   36
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   35
         Top             =   840
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   34
         Top             =   600
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 8"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 7"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 6"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 5"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 4"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 3"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 2"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 1"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Set All Channels"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   23
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   4440
         TabIndex        =   22
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   4440
         TabIndex        =   21
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label PWMValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   4440
         TabIndex        =   20
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 9"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 10"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 11"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 12"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "PWM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BITCH_Click(Index As Integer)
    
    FOUT = 0
    If BITCH(0).Value = 1 Then FOUT = FOUT + 1
    If BITCH(1).Value = 1 Then FOUT = FOUT + 2
    If BITCH(2).Value = 1 Then FOUT = FOUT + 4
    If BITCH(3).Value = 1 Then FOUT = FOUT + 8
    If BITCH(4).Value = 1 Then FOUT = FOUT + 16
    If BITCH(5).Value = 1 Then FOUT = FOUT + 32
    If BITCH(6).Value = 1 Then FOUT = FOUT + 64
    If BITCH(7).Value = 1 Then FOUT = FOUT + 128
    If BITCH(8).Value = 1 Then FOUT = FOUT + 256
    If BITCH(9).Value = 1 Then FOUT = FOUT + 512
    If BITCH(10).Value = 1 Then FOUT = FOUT + 1024
    If BITCH(11).Value = 1 Then FOUT = FOUT + 2048
    If BITCH(12).Value = 1 Then FOUT = FOUT + 4096
    If BITCH(13).Value = 1 Then FOUT = FOUT + 8192
    If BITCH(14).Value = 1 Then FOUT = FOUT + 16384
    If BITCH(15).Value = 1 Then FOUT = FOUT + 32768
    
    
    LSB = (FOUT And 255)
    MSB = (FOUT And 65280) / 256
    
    Debug.Print FOUT; LSB; MSB
    
    Form1.ProXR1.SendData 254       'Set Frequency
    Form1.ProXR1.SendData 189
    Form1.ProXR1.SendData Index
    Form1.ProXR1.SendData LSB
    Form1.ProXR1.SendData MSB
    Form1.ProXR1.GetData            'Get Confirmation
    
    
End Sub

Private Sub ClearBuffer_Click()
    Form1.ProXR1.ClearBuffer
End Sub
Private Sub FRQ_Change(Index As Integer)
    FRQValue(Index).Caption = FRQ(Index).Value
    KHZ(Index).Caption = Mid$(str$((FRQ(Index).Value * 0.01762) + 1.23), 1, 6) + " KHz"

    LSB = (FRQ(Index) And 255)
    MSB = (FRQ(Index) And 65280) / 256
    
    Debug.Print LSB; MSB
    
    Form1.ProXR1.SendData 254       'Set Frequency
    Form1.ProXR1.SendData 189
    Form1.ProXR1.SendData Index
    Form1.ProXR1.SendData LSB
    Form1.ProXR1.SendData MSB
    Form1.ProXR1.GetData            'Get Confirmation
        
End Sub
Private Sub FRQ_Scroll(Index As Integer)
    FRQValue(Index).Caption = FRQ(Index).Value
    KHZ(Index).Caption = Mid$(str$((FRQ(Index).Value * 0.01762) + 1.23), 1, 6) + " KHz"
    
    LSB = (FRQ(Index) And 255)
    MSB = (FRQ(Index) And 65280) / 256

    Debug.Print LSB; MSB
    
    Form1.ProXR1.SendData 254   'Set Frequency
    Form1.ProXR1.SendData 189
    Form1.ProXR1.SendData Index
    Form1.ProXR1.SendData LSB
    Form1.ProXR1.SendData MSB
    Form1.ProXR1.GetData        'Get Confirmation
    
End Sub
Private Sub PWM_Change(Index As Integer)
    PWMValue(Index).Caption = PWM(Index).Value
    Form1.ProXR1.SendData 254
    Form1.ProXR1.SendData 188
    Form1.ProXR1.SendData Index
    Form1.ProXR1.SendData PWM(Index).Value
    RX = Form1.ProXR1.GetData
    'If RX <> 188 Then
    '    MsgBox ("Master Processor is Unable to Communicate with PWM Subprocessors.")
    'End If
End Sub
Private Sub PWM_Scroll(Index As Integer)
    PWMValue(Index).Caption = PWM(Index).Value
    Form1.ProXR1.SendData 254
    Form1.ProXR1.SendData 188
    Form1.ProXR1.SendData Index
    Form1.ProXR1.SendData PWM(Index).Value
    Form1.ProXR1.GetData
    'If RX <> 188 Then
    '    MsgBox ("Master Processor is Unable to Communicate with PWM Subprocessors.")
    'End If
End Sub

Private Sub ReadConfig_Click()

    Form1.ProXR1.SendData 254
    Form1.ProXR1.SendData 191
    RX = Form1.ProXR1.GetData
    
    
    Frame1.Visible = True
    Frame2.Visible = True
    
    For x = 1 To 32
        KHZ(x).Caption = "1.23 KHz"
    Next x
End Sub
