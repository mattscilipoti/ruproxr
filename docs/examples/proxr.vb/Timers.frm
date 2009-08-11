VERSION 5.00
Begin VB.Form Timers 
   Caption         =   "Background Timers"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form3"
   ScaleHeight     =   8985
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Timer Component Functions"
      Height          =   3975
      Left            =   120
      TabIndex        =   25
      Top             =   4560
      Width           =   10935
      Begin VB.CommandButton QueryTimer 
         Caption         =   "Query Selected Timer"
         Height          =   735
         Left            =   2520
         TabIndex        =   67
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton WriteTimingData2 
         Caption         =   "Write as Pulse Timer"
         Height          =   735
         Left            =   1320
         TabIndex        =   65
         Top             =   3120
         Width           =   1095
      End
      Begin VB.HScrollBar TimerSlider2 
         Height          =   375
         Left            =   840
         Max             =   15
         TabIndex        =   49
         Top             =   720
         Width           =   3615
      End
      Begin VB.HScrollBar Hours2 
         Height          =   375
         Left            =   840
         Max             =   255
         TabIndex        =   48
         Top             =   1200
         Width           =   3615
      End
      Begin VB.HScrollBar Minutes2 
         Height          =   375
         Left            =   840
         Max             =   255
         TabIndex        =   47
         Top             =   1680
         Width           =   3615
      End
      Begin VB.HScrollBar Seconds2 
         Height          =   375
         Left            =   840
         Max             =   255
         TabIndex        =   46
         Top             =   2160
         Value           =   1
         Width           =   3615
      End
      Begin VB.HScrollBar Relay2 
         Height          =   375
         Left            =   840
         Max             =   208
         TabIndex        =   45
         Top             =   2640
         Width           =   3615
      End
      Begin VB.CommandButton WriteTimingData 
         Caption         =   "Write as Duration Timer"
         Height          =   735
         Left            =   120
         TabIndex        =   44
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Triggering Timers"
         Height          =   3255
         Left            =   5880
         TabIndex        =   26
         Top             =   600
         Width           =   4935
         Begin VB.CommandButton SelectNone 
            Caption         =   "Select None"
            Height          =   255
            Left            =   3600
            TabIndex        =   71
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CommandButton SelectAll 
            Caption         =   "Select All"
            Height          =   255
            Left            =   2520
            TabIndex        =   70
            Top             =   2880
            Width           =   1095
         End
         Begin VB.CommandButton TriggerTimer 
            Caption         =   "Trigger/Halt Selected Timers"
            Height          =   375
            Left            =   120
            TabIndex        =   57
            Top             =   2160
            Width           =   2295
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 15"
            Height          =   255
            Index           =   15
            Left            =   3720
            TabIndex        =   42
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 14"
            Height          =   255
            Index           =   14
            Left            =   3720
            TabIndex        =   41
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 13"
            Height          =   255
            Index           =   13
            Left            =   3720
            TabIndex        =   40
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 12"
            Height          =   255
            Index           =   12
            Left            =   3720
            TabIndex        =   39
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 11"
            Height          =   255
            Index           =   11
            Left            =   2520
            TabIndex        =   38
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 10"
            Height          =   255
            Index           =   10
            Left            =   2520
            TabIndex        =   37
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 9"
            Height          =   255
            Index           =   9
            Left            =   2520
            TabIndex        =   36
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 8"
            Height          =   255
            Index           =   8
            Left            =   2520
            TabIndex        =   35
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 7"
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   34
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 6"
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   33
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 5"
            Height          =   255
            Index           =   5
            Left            =   1320
            TabIndex        =   32
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 4"
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   31
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 3"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   1800
            Width           =   975
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 2"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   1560
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 1"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox TimerTrig 
            Caption         =   "Timer 0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "NOTE: Selected Timers will be Started, Unselected Timers will be Halted."
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   2520
            TabIndex        =   69
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Command Sent to Controller:"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            TabIndex        =   66
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label Label14 
            Caption         =   $"Timers.frx":0000
            Height          =   855
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Command Sent to Controller:"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3600
         TabIndex        =   63
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   61
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   60
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   59
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   58
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   $"Timers.frx":00FF
         Height          =   495
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   10455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TIMER:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   54
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "HOURS:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "MINS:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SECS:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "RELAY:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   2760
         Width           =   615
      End
   End
   Begin VB.Frame TimerFrame 
      Caption         =   "Timer Testing and Examples"
      Height          =   4335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.CommandButton Command12 
         BackColor       =   &H0080FF80&
         Caption         =   "Timer Calibration and Test Examples"
         Height          =   495
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3720
         Width           =   4575
      End
      Begin VB.HScrollBar TimerSlider 
         Height          =   375
         Left            =   840
         Max             =   15
         TabIndex        =   14
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton PulseTimer 
         Caption         =   "Server Reboot Pulse Timer"
         Height          =   375
         Left            =   3000
         TabIndex        =   13
         Top             =   2640
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Activate Relay for Duration of Timer"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   2775
      End
      Begin VB.HScrollBar Relay 
         Height          =   375
         Left            =   840
         Max             =   255
         TabIndex        =   7
         Top             =   2160
         Width           =   3615
      End
      Begin VB.HScrollBar Seconds 
         Height          =   375
         Left            =   840
         Max             =   255
         TabIndex        =   6
         Top             =   1680
         Value           =   1
         Width           =   3615
      End
      Begin VB.HScrollBar Minutes 
         Height          =   375
         Left            =   840
         Max             =   255
         TabIndex        =   3
         Top             =   1200
         Width           =   3615
      End
      Begin VB.HScrollBar Hours 
         Height          =   375
         Left            =   840
         Max             =   255
         TabIndex        =   1
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Command Sent to Controller:"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8160
         TabIndex        =   64
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   $"Timers.frx":01D1
         Height          =   975
         Left            =   3000
         TabIndex        =   24
         Top             =   3120
         Width           =   3135
      End
      Begin VB.Label Label10 
         Caption         =   "When this button is pressed, the selected relay will be activated.  The relay will turn off when the timer expires."
         Height          =   855
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "The timer can be applied to any of 256 relays.  Set the relay you would like this timer to be tied to using this slider."
         Height          =   495
         Left            =   5280
         TabIndex        =   22
         Top             =   2160
         Width           =   5535
      End
      Begin VB.Label Label8 
         Caption         =   "Set the Seconds of the timer here."
         Height          =   255
         Left            =   5280
         TabIndex        =   21
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label Label7 
         Caption         =   "Set the Minutes of the timer here."
         Height          =   255
         Left            =   5280
         TabIndex        =   20
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label Label6 
         Caption         =   "Set the Hours of the timer here."
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   840
         Width           =   5535
      End
      Begin VB.Label Label5 
         Caption         =   $"Timers.frx":02A6
         Height          =   615
         Left            =   5280
         TabIndex        =   18
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TIMER:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.Label TIM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.Label HR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.Label RE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label SC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label MN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "RELAY:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SECS:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "MINS:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "HOURS:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Label Label23 
      Caption         =   $"Timers.frx":0393
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   8640
      Width           =   10935
   End
End
Attribute VB_Name = "Timers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (50 + TimerSlider) 'Activate Relay Timer (Index) for Duration
    Form1.PutData (Hours)   'Hours
    Form1.PutData (Minutes) 'Minutes
    Form1.PutData (Seconds) 'Seconds
    Form1.PutData (RELAY)   'Relay
    Form1.GetData
    Label22.Caption = "Command Sent to Controller: 254," + Str$(50 + TimerSlider) + "," + Str$(Hours) + "," + Str$(Minutes) + "," + Str$(Seconds) + "," + Str$(RELAY)
End Sub
Private Sub Command12_Click()
    AdvancedTimers.Visible = True
    AdvancedTimers.ZOrder 0
End Sub

Private Sub Hours_Change()
    HR.Caption = Hours.Value
End Sub
Private Sub Hours_Scroll()
    HR.Caption = Hours.Value
End Sub

Private Sub Hours2_Change()
    Label15.Caption = Hours2.Value
End Sub
Private Sub Hours2_Scroll()
    Label15.Caption = Hours2.Value
End Sub

Private Sub Minutes_Change()
    MN.Caption = Minutes.Value
End Sub
Private Sub Minutes_Scroll()
    MN.Caption = Minutes.Value
End Sub

Private Sub Minutes2_Change()
    Label16.Caption = Minutes2.Value
End Sub
Private Sub Minutes2_Scroll()
    Label16.Caption = Minutes2.Value
End Sub
Private Sub PulseTimer_Click()
    Form1.PutData (254)            'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (70 + TimerSlider) 'Activate Relay Timer for Duration
    Form1.PutData (Hours)   'Hours
    Form1.PutData (Minutes) 'Minutes
    Form1.PutData (Seconds) 'Seconds
    Form1.PutData (RELAY)   'Relay
    Form1.GetData
    Label22.Caption = "Command Sent to Controller: 254," + Str$(70 + TimerSlider) + "," + Str$(Hours) + "," + Str$(Minutes) + "," + Str$(Seconds) + "," + Str$(RELAY)
End Sub
Private Sub QueryTimer_Click()
    QueryTimer.Enabled = False
    Form1.PutData (254)            'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (130) 'Query Timer Command
    Form1.PutData (TimerSlider2)
    Hours2 = Form1.GetData
    Minutes2 = Form1.GetData
    Seconds2 = Form1.GetData
    Relay2 = Form1.GetData
    QueryTimer.Enabled = True
    Label21.Caption = "Command Sent to Controller: 254, 130," + Str$(TimerSlider2)
End Sub
Private Sub Relay2_Change()
    Label18.Caption = Relay2.Value
End Sub
Private Sub Relay2_Scroll()
    Label18.Caption = Relay2.Value
End Sub
Private Sub Seconds_Change()
    SC.Caption = Seconds.Value
End Sub
Private Sub Seconds_Scroll()
    SC.Caption = Seconds.Value
End Sub
Private Sub Relay_Change()
    RE.Caption = RELAY.Value
End Sub
Private Sub Relay_Scroll()
    RE.Caption = RELAY.Value
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Timers.Visible = False
End Sub
Private Sub Seconds2_Change()
    Label17.Caption = Seconds2.Value
End Sub
Private Sub Seconds2_Scroll()
    Label17.Caption = Seconds2.Value
End Sub
Private Sub SelectAll_Click()
    For X = 0 To 15
        TimerTrig(X).Value = 1
    Next X
End Sub
Private Sub SelectNone_Click()
    For X = 0 To 15
        TimerTrig(X).Value = 0
    Next X
End Sub
Private Sub StoreCalibrationValue_Click()
    MsgBox ("Remember, this button will have no effect unless the controller is in PROGRAM mode.  Turn Off All DIP Switches and Power Cycle the Controller BEFORE using this button.  This is NOT an error message, just a reminder.")
End Sub
Private Sub TimerSlider_Change()
    TIM.Caption = TimerSlider.Value
End Sub
Private Sub TimerSlider_Scroll()
    TIM.Caption = TimerSlider.Value
End Sub
Private Sub TimerSlider2_Change()
    Label13.Caption = TimerSlider2.Value
End Sub
Private Sub TimerSlider2_Scroll()
    Label13.Caption = TimerSlider2.Value
End Sub
Private Sub TriggerTimer_Click()
    'Calculate the First Trigger Byte to Send to the Controller
    Part1 = 0
    If TimerTrig(0).Value = 1 Then Part1 = Part1 + 1
    If TimerTrig(1).Value = 1 Then Part1 = Part1 + 2
    If TimerTrig(2).Value = 1 Then Part1 = Part1 + 4
    If TimerTrig(3).Value = 1 Then Part1 = Part1 + 8
    If TimerTrig(4).Value = 1 Then Part1 = Part1 + 16
    If TimerTrig(5).Value = 1 Then Part1 = Part1 + 32
    If TimerTrig(6).Value = 1 Then Part1 = Part1 + 64
    If TimerTrig(7).Value = 1 Then Part1 = Part1 + 128
    'Calculate the Second Trigger Byte to Send to the Controller
    Part2 = 0
    If TimerTrig(8).Value = 1 Then Part2 = Part2 + 1
    If TimerTrig(9).Value = 1 Then Part2 = Part2 + 2
    If TimerTrig(10).Value = 1 Then Part2 = Part2 + 4
    If TimerTrig(11).Value = 1 Then Part2 = Part2 + 8
    If TimerTrig(12).Value = 1 Then Part2 = Part2 + 16
    If TimerTrig(13).Value = 1 Then Part2 = Part2 + 32
    If TimerTrig(14).Value = 1 Then Part2 = Part2 + 64
    If TimerTrig(15).Value = 1 Then Part2 = Part2 + 128
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (131) 'Configure Timer Register
    Form1.PutData (Part1) 'Lower 8 Bits
    Form1.PutData (Part2) 'Upper 8 Bits
    Form1.GetData
    Label20.Caption = "Command Sent to Controller: 254, 131," + Str$(Part1) + "," + Str$(Part2)
End Sub
Private Sub WriteTimingData_Click()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (90 + TimerSlider2) 'Set Relay Timer for Duration, But Do NOT Activate
    Form1.PutData (Hours2)   'Hours
    Form1.PutData (Minutes2) 'Minutes
    Form1.PutData (Seconds2) 'Seconds
    Form1.PutData (Relay2)   'Relay
    Form1.GetData
    Label21.Caption = "Command Sent to Controller: 254," + Str$(90 + TimerSlider2) + "," + Str$(Hours2) + "," + Str$(Minutes2) + "," + Str$(Seconds2) + "," + Str$(Relay2)
End Sub
Private Sub WriteTimingData2_Click()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (110 + TimerSlider2) 'Set Relay Timer for Pulse, But Do NOT Activate
    Form1.PutData (Hours2)   'Hours
    Form1.PutData (Minutes2) 'Minutes
    Form1.PutData (Seconds2) 'Seconds
    Form1.PutData (Relay2)   'Relay
    Form1.GetData
    Label21.Caption = "Command Sent to Controller: 254," + Str$(110 + TimerSlider2) + "," + Str$(Hours2) + "," + Str$(Minutes2) + "," + Str$(Seconds2) + "," + Str$(Relay2)
End Sub
