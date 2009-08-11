VERSION 5.00
Begin VB.Form FiberOpticOptions 
   Caption         =   "Fiber Optic Interface Tuning"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form4"
   ScaleHeight     =   4935
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Choose Your Own Custom Tuning Value (Lower=Faster, Higher Values=More Reliable)"
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   7815
      Begin VB.CommandButton ReadTiming 
         Caption         =   "Read Serial Timing Value from Controller"
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton WriteTiming 
         Caption         =   "Store Serial Timing Value Into the Controller"
         Height          =   495
         Left            =   5520
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.HScrollBar TimingSlider 
         Height          =   255
         Left            =   120
         Max             =   255
         Min             =   3
         TabIndex        =   8
         Top             =   240
         Value           =   5
         Width           =   2415
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
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   $"FiberOpticOptions.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label3 
         Caption         =   "This Setting Can ONLY Be Changed while in Configuration Mode."
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   5520
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.OptionButton Tune 
      Caption         =   "Tune for RS-232 Interface "
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4215
   End
   Begin VB.OptionButton Tune 
      Caption         =   "Tune for USB Interface "
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   4215
   End
   Begin VB.OptionButton Tune 
      Caption         =   "Tune for Bluetooth Interface"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4215
   End
   Begin VB.OptionButton Tune 
      Caption         =   "Tune for ZigBee Interface "
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.OptionButton Tune 
      Caption         =   "Tune for Wi-Fi 802.11b and Ethernet Interface"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.OptionButton Tune 
      Caption         =   "Use a Safe and Conservative Timing Compatible with ALL Interface Options (Not the Best Performance)"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   $"FiberOpticOptions.frx":01E2
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "FiberOpticOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AdvancedFeatureButton_Click()
    AdvancedFeatures.Visible = True
    AdvancedFeatures.ZOrder 0
End Sub
Private Sub ReadTiming_Click()
    Form1.ProXR1.ClearBuffer
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
Private Sub TimingSlider_Change()
    TimingTag.Caption = TimingSlider
    If TimingSlider.Value = 125 Then Tune(0).Value = True
    If TimingSlider.Value = 120 Then Tune(1).Value = True
    If TimingSlider.Value = 80 Then Tune(2).Value = True
    If TimingSlider.Value = 37 Then Tune(3).Value = True
    If TimingSlider.Value = 36 Then Tune(4).Value = True
    If TimingSlider.Value = 35 Then Tune(5).Value = True
End Sub
Private Sub TimingSlider_Scroll()
    TimingTag.Caption = TimingSlider
    If TimingSlider.Value = 125 Then Tune(0).Value = True
    If TimingSlider.Value = 120 Then Tune(1).Value = True
    If TimingSlider.Value = 80 Then Tune(2).Value = True
    If TimingSlider.Value = 37 Then Tune(3).Value = True
    If TimingSlider.Value = 36 Then Tune(4).Value = True
    If TimingSlider.Value = 35 Then Tune(5).Value = True
End Sub
Private Sub Tune_Click(Index As Integer)
    If Index = 0 Then TimingSlider.Value = 125  'ALL
    If Index = 1 Then TimingSlider.Value = 120  'NETWORK
    If Index = 2 Then TimingSlider.Value = 80   'ZIGBEE
    If Index = 3 Then TimingSlider.Value = 37   'BLUETOOTH
    If Index = 4 Then TimingSlider.Value = 36   'USB
    If Index = 5 Then TimingSlider.Value = 35   'RS232
End Sub
Private Sub WriteTiming_Click()
    If AdvancedFeatures.Check_Store = 1 Then
        Form1.ProXR1.ClearBuffer
        Form1.PutData (254)            'Enter Command Mode
        Form1.PutData (50)             'Timer/Setup Branch Commands
        Form1.PutData (139)            'Set Serial Timing Command
        Form1.PutData (TimingSlider)   'Send Value to Controller
        Temp = Form1.GetData
    End If
End Sub
