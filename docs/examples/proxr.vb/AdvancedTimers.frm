VERSION 5.00
Begin VB.Form AdvancedTimers 
   Caption         =   "Advanced Timing Functions                                                WWW.CONTROLANYTHING.COM"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form3"
   ScaleHeight     =   8070
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Calibrate the Timer: 1 Second = xxx Counts"
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   10095
      Begin VB.OptionButton TimeTest 
         Caption         =   "48 Hours"
         Height          =   255
         Index           =   11
         Left            =   8760
         TabIndex        =   34
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "24 Hours"
         Height          =   255
         Index           =   10
         Left            =   7440
         TabIndex        =   33
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "10 Hours"
         Height          =   255
         Index           =   9
         Left            =   6120
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "5 Hours"
         Height          =   255
         Index           =   8
         Left            =   4920
         TabIndex        =   31
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "1 Hour"
         Height          =   255
         Index           =   7
         Left            =   8760
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "30 Minutes"
         Height          =   255
         Index           =   6
         Left            =   7440
         TabIndex        =   29
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "10 Minutes"
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   28
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "5 Minutes"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   27
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "5 Seconds"
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   24
         Top             =   1080
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "1 Minute"
         Height          =   255
         Index           =   3
         Left            =   8760
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "30 Seconds"
         Height          =   255
         Index           =   2
         Left            =   7440
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton TimeTest 
         Caption         =   "10 Seconds"
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton T16Timer 
         Caption         =   "Test Calibration Value with 16 Timers"
         Height          =   495
         Left            =   8280
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton T8Timer 
         Caption         =   "Test Calibration Value with 8 Timers"
         Height          =   495
         Left            =   6600
         TabIndex        =   17
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton T1Timer 
         Caption         =   "Test Calibration Value with 1 Timer"
         Height          =   495
         Left            =   4920
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
      Begin VB.HScrollBar TimeCalSlider 
         Height          =   255
         Left            =   120
         Min             =   100
         TabIndex        =   15
         Top             =   240
         Value           =   255
         Width           =   4455
      End
      Begin VB.CommandButton ReadCalibrationValue 
         Caption         =   "Read the Current Calibration Value"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton StoreCalibrationValue 
         Caption         =   "Store Above Calibration Value into Controller"
         Height          =   495
         Left            =   2400
         TabIndex        =   13
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   $"AdvancedTimers.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   4800
         TabIndex        =   26
         Top             =   1920
         Width           =   5175
      End
      Begin VB.Label Label10 
         Caption         =   $"AdvancedTimers.frx":00A0
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label9 
         Caption         =   "Seconds . Jiffies (1/100 Second)"
         Height          =   255
         Left            =   7200
         TabIndex        =   23
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.000"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server Reboot Features"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   10095
      Begin VB.CommandButton PulseTimer 
         Caption         =   "Server Timer: Click this button to start the timer, keep clicking this button to keep the server from rebooting."
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel Timers"
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   $"AdvancedTimers.frx":0186
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.CommandButton CancelTimers 
      Caption         =   "Cancel Timers"
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   $"AdvancedTimers.frx":03AA
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   10095
      Begin VB.CommandButton Unprogram 
         Caption         =   "Unprogram (Clear) Startup Status of Relay 1"
         Height          =   495
         Left            =   8040
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton KeepAlive 
         Caption         =   "Keep Alive Timer...You have to Keep Pressing this Button to Keep the Relay Alive...otherwise, it times out in 5 seconds."
         Height          =   495
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   5175
      End
      Begin VB.CommandButton ProgramRelay1 
         Caption         =   "Program Relay 1 So it is On when Power is Applied to the Board"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label23 
         Caption         =   $"AdvancedTimers.frx":0442
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   9855
      End
   End
   Begin VB.CommandButton Test2 
      Caption         =   "Test 2 - Demo 8 timers Activating Bank 2 Relays, Deactivating 1 automatically every second while Handling Serial Communications."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   10095
   End
   Begin VB.CommandButton Test1 
      Caption         =   "Test 1 - Demo All 16 timers Activating the First 16 Relays, Deactivating 1 automatically every second."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "AdvancedTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DAT As Long
Private Sub CancelTimers_Click()
    'When this window is closed, the following commands are set to turn off
    'the "timing marks" that are transmitted along with relay control commands
    'that are associated using the relay timers.  These timing marks allows
    'this program to help you calibrate the timers.
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)     'Timer/Setup Branch Commands
    Form1.PutData (135)    'Turn OFF Calibrator Command
    Form1.GetData
    
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (131) 'Configure Timer Register
    Form1.PutData (0)   'Lower 8 Bits
    Form1.PutData (0)   'Upper 8 Bits
    Form1.GetData
End Sub
Private Sub Command1_Click()
    'When this window is closed, the following commands are set to turn off
    'the "timing marks" that are transmitted along with relay control commands
    'that are associated using the relay timers.  These timing marks allows
    'this program to help you calibrate the timers.
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)     'Timer/Setup Branch Commands
    Form1.PutData (135)    'Turn OFF Calibrator Command
    Form1.GetData
    
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (131) 'Configure Timer Register
    Form1.PutData (0)   'Lower 8 Bits
    Form1.PutData (0)   'Upper 8 Bits
    Form1.GetData
End Sub
Private Sub Form_Load()
    ReadCalibrationValue_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'When this window is closed, the following commands are set to turn off
    'the "timing marks" that are transmitted along with relay control commands
    'that are associated using the relay timers.  These timing marks allows
    'this program to help you calibrate the timers.
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)     'Timer/Setup Branch Commands
    Form1.PutData (135)    'Turn OFF Calibrator Command
    Form1.GetData
End Sub

Private Sub ProgramRelay1_Click()
    'Activate Relay 1 Only on Bank 1
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (140) 'Set the Status of Relays
    Form1.PutData (1)   'Value to Write to Relay Bank
    Form1.PutData (1)   'Direct this Command to Relay Bank 1
    Form1.GetData
    'Store Relay Setting as Powerup Default Status
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (142) 'Store Current Memory Bank as Power Up Defaults
    Form1.PutData (1)   'Direct this Command to Relay Bank 1
    Form1.GetData
End Sub
Private Sub PulseTimer_Click()
    Form1.PutData (254)     'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (70)      'Activate Relay Timer 0 for Duration Timing
    Form1.PutData (0)       'Hours
    Form1.PutData (0)       'Minutes
    Form1.PutData (5)       'Seconds
    Form1.PutData (0)       'Relay
    Form1.GetData
End Sub
Private Sub T16Timer_Click()
    'Turn On Calibrator
    'This Command tells the Controller to Signal the Starting and Stopping of a Timer
    'When the Calibrator is ON:
    'The Controller will Send 90 at the beginning of any timer.
    'The Controller will send 91 at the end of any timer.
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (134)    'Turn On Calibrator Command
    Form1.GetData
    'Determine Duration of Test
    If TimeTest(0).Value = True Then HR = 0: MN = 0: SC = 5
    If TimeTest(1).Value = True Then HR = 0: MN = 0: SC = 10
    If TimeTest(2).Value = True Then HR = 0: MN = 0: SC = 30
    If TimeTest(3).Value = True Then HR = 0: MN = 0: SC = 60
    If TimeTest(4).Value = True Then HR = 0: MN = 5: SC = 0
    If TimeTest(5).Value = True Then HR = 0: MN = 10: SC = 0
    If TimeTest(6).Value = True Then HR = 0: MN = 30: SC = 0
    If TimeTest(7).Value = True Then HR = 0: MN = 60: SC = 0
    If TimeTest(8).Value = True Then HR = 5: MN = 0: SC = 0
    If TimeTest(9).Value = True Then HR = 10: MN = 0: SC = 0
   If TimeTest(10).Value = True Then HR = 24: MN = 0: SC = 0
   If TimeTest(11).Value = True Then HR = 48: MN = 0: SC = 0
   
    'Setup 8 Timers for Selected Duration
    For x = 0 To 15
        Form1.PutData (254)     'Enter Command Mode
        Form1.PutData (50)  'Timer/Setup Branch Commands
        Form1.PutData (90 + x)  'Activate Relay Timer x for Duration Timing
        Form1.PutData (HR)      'Hours
        Form1.PutData (MN)      'Minutes
        Form1.PutData (SC)      'Seconds
        Form1.PutData (x)       'Relay
        Form1.GetData
    Next x
    
    'Activate All 16 Timers
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (131) 'Configure Timer Register
    Form1.PutData (255) 'Lower 8 Bits (First 8 Timers Activated)
    Form1.PutData (255) 'Lower 8 Bits (Last 8 Timers Activated)
    
    Dim i As Long
    Do
        DoEvents
        i = Form1.ProXR1.GetData
    Loop Until i >= 0
    
    Start = i
    STime = Timer
    Do
        DoEvents
        i = Form1.ProXR1.GetData
    Loop Until i >= 0
    StopT = Asc(i)
    TimeS = Timer

    If Start = 90 Then
        If StopT = 91 Then
            Duration = TimeS - STime
            Label7.Caption = Str$(Duration)
        End If
    End If
End Sub

Private Sub T1Timer_Click()
    'Turn On Calibrator
    'This Command tells the Controller to Signal the Starting and Stopping of a Timer
    'When the Calibrator is ON:
    'The Controller will Send 90 at the beginning of any timer.
    'The Controller will send 91 at the end of any timer.
    'The elapsed time between 90 and 91 is used to help you calibrate the timer
    Debug.Print "Turn On Calibrator"
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (134)    'Turn On Calibrator Command
    Form1.GetData
    Debug.Print "------------------"
    'Determine Duration of Test
    If TimeTest(0).Value = True Then HR = 0: MN = 0: SC = 5
    If TimeTest(1).Value = True Then HR = 0: MN = 0: SC = 10
    If TimeTest(2).Value = True Then HR = 0: MN = 0: SC = 30
    If TimeTest(3).Value = True Then HR = 0: MN = 0: SC = 60
    If TimeTest(4).Value = True Then HR = 0: MN = 5: SC = 0
    If TimeTest(5).Value = True Then HR = 0: MN = 10: SC = 0
    If TimeTest(6).Value = True Then HR = 0: MN = 30: SC = 0
    If TimeTest(7).Value = True Then HR = 0: MN = 60: SC = 0
    If TimeTest(8).Value = True Then HR = 5: MN = 0: SC = 0
    If TimeTest(9).Value = True Then HR = 10: MN = 0: SC = 0
   If TimeTest(10).Value = True Then HR = 24: MN = 0: SC = 0
   If TimeTest(11).Value = True Then HR = 48: MN = 0: SC = 0
    Debug.Print "Setup Timer"
    Form1.PutData (254)     'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (50)      'Activate Relay Timer 0 for Duration Timing
    Form1.PutData (HR)      'Hours
    Form1.PutData (MN)      'Minutes
    Form1.PutData (SC)      'Seconds
    Form1.PutData (0)       'Relay
    Debug.Print "Waiting for Start Time";
    Dim i As Long
    Do
        DoEvents
        i = Form1.ProXR1.GetData
    Loop Until i >= 0
    Start = i
    STime = Timer
    Do
        DoEvents
        i = Form1.ProXR1.GetData
    Loop Until i >= 0
    StopT = Asc(i)
    TimeS = Timer
    
    Debug.Print StopT
        'Debug.Print StopT; TimeS

    If Start = 90 Then
        If StopT = 91 Then
            Duration = TimeS - STime
            Label7.Caption = Str$(Duration)
        End If
    End If
    
End Sub
Private Sub T8Timer_Click()
    'Turn On Calibrator
    'This Command tells the Controller to Signal the Starting and Stopping of a Timer
    'When the Calibrator is ON:
    'The Controller will Send 90 at the beginning of any timer.
    'The Controller will send 91 at the end of any timer.
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (134)    'Turn On Calibrator Command
    Form1.GetData
    'Determine Duration of Test
    If TimeTest(0).Value = True Then HR = 0: MN = 0: SC = 5
    If TimeTest(1).Value = True Then HR = 0: MN = 0: SC = 10
    If TimeTest(2).Value = True Then HR = 0: MN = 0: SC = 30
    If TimeTest(3).Value = True Then HR = 0: MN = 0: SC = 60
    If TimeTest(4).Value = True Then HR = 0: MN = 5: SC = 0
    If TimeTest(5).Value = True Then HR = 0: MN = 10: SC = 0
    If TimeTest(6).Value = True Then HR = 0: MN = 30: SC = 0
    If TimeTest(7).Value = True Then HR = 0: MN = 60: SC = 0
    If TimeTest(8).Value = True Then HR = 5: MN = 0: SC = 0
    If TimeTest(9).Value = True Then HR = 10: MN = 0: SC = 0
   If TimeTest(10).Value = True Then HR = 24: MN = 0: SC = 0
   If TimeTest(11).Value = True Then HR = 48: MN = 0: SC = 0
   
    'Setup 8 Timers for Selected Duration
    For x = 0 To 7
        Form1.PutData (254)     'Enter Command Mode
        Form1.PutData (50)  'Timer/Setup Branch Commands
        Form1.PutData (90 + x)  'Activate Relay Timer x for Duration Timing
        Form1.PutData (HR)      'Hours
        Form1.PutData (MN)      'Minutes
        Form1.PutData (SC)      'Seconds
        Form1.PutData (x)       'Relay
        Form1.GetData
    Next x
    
    'Activate All 8 Timers
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (131) 'Configure Timer Register
    Form1.PutData (255) 'Lower 8 Bits (First 8 Timers Activated)
    Form1.PutData (0)   'Upper 8 Bits (Last 8 Timers Not Activated)
    
    Dim i As Long
    Do
        DoEvents
        i = Form1.ProXR1.GetData
    Loop Until i >= 0
    Start = i
    STime = Timer
    Do
        DoEvents
        i = Form1.ProXR1.GetData
    Loop Until i >= 0
    StopT = Asc(i)
    TimeS = Timer

    If Start = 90 Then
        If StopT = 91 Then
            Duration = TimeS - STime
            Label7.Caption = Str$(Duration)
        End If
    End If
End Sub

Private Sub Test1_Click()
    For Temp = 0 To 15
        Form1.PutData (254)        'Enter Command Mode
        Form1.PutData (50)  'Timer/Setup Branch Commands
        Form1.PutData (90 + Temp)  'Set Relay Timer for Duration, But Do NOT Activate
        Form1.PutData (0)          'Hours
        Form1.PutData (0)          'Minutes
        Form1.PutData (1 + Temp)   'Seconds
        Form1.PutData (Temp)       'Relay
        Form1.GetData
    Next Temp
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (50)  'Timer/Setup Branch Commands
    Form1.PutData (131) 'Activate Timers
    Form1.PutData (255) 'Lower 8 Bits
    Form1.PutData (255) 'Upper 8 Bits
    Form1.GetData
End Sub
Private Sub Test2_Click()
    
    'Program Timers 8-15 to Automatically Fall across 8 seconds
    For Temp = 8 To 15
        Form1.PutData (254)        'Enter Command Mode
        Form1.PutData (50)         'Timer/Setup Branch Commands
        Form1.PutData (90 + Temp)  'Set Relay Timer for Duration, But Do NOT Activate
        Form1.PutData (0)          'Hours
        Form1.PutData (0)          'Minutes
        Form1.PutData (Temp - 7)   'Seconds
        Form1.PutData (Temp)       'Relay
        Form1.GetData
    Next Temp
  
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (50)         'Timer/Setup Branch Commands
    Form1.PutData (131) 'Activate Timers
    Form1.PutData (0)   'Lower 8 Bits
    Form1.PutData (255) 'Upper 8 Bits
    Form1.GetData
        
  'Demo Relay Controls in the Background
   For N = 255 To 0 Step -1
        Form1.PutData (254) 'Enter Command Mode
        Form1.PutData (140) 'Set the Status of Relays
        Form1.PutData (N)   'Value to Write to Relay Bank
        Form1.PutData (1)   'Direct this Command to Relay Bank 1
        Form1.GetData
        For dela = 1 To 12000
            DoEvents
        Next dela
    Next N

End Sub
Private Sub KeepAlive_Click()
    Form1.PutData (254)     'Enter Command Mode
    Form1.PutData (50)      'Timer/Setup Branch Commands
    Form1.PutData (50)      'Activate Relay Timer 0 for Duration Timing
    Form1.PutData (0)       'Hours
    Form1.PutData (0)       'Minutes
    Form1.PutData (5)       'Seconds
    Form1.PutData (0)       'Relay
    Form1.GetData
End Sub
Public Sub TimeCalSlider_Change()
    Temp = TimeCalSlider.Value
    Frame3.Caption = "Calibrate the Timer: 1 Second =" + Str$(Temp * 2) + " Counts."
End Sub
Public Sub TimeCalSlider_Scroll()
    Temp = TimeCalSlider.Value
    Frame3.Caption = "Calibrate the Timer: 1 Second =" + Str$(Temp * 2) + " Counts."
End Sub
Private Sub Unprogram_Click()
    'DeActivate Relay 1 Only on Bank 1
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (140) 'Set the Status of Relays
    Form1.PutData (0)   'Value to Write to Relay Bank
    Form1.PutData (1)   'Direct this Command to Relay Bank 1
    Form1.GetData
    'Store Relay Setting as Powerup Default Status
    Form1.PutData (254) 'Enter Command Mode
    Form1.PutData (142) 'Store Current Memory Bank as Power Up Defaults
    Form1.PutData (1)   'Direct this Command to Relay Bank 1
    Form1.GetData
End Sub

Public Sub ReadCalibrationValue_Click()
        Form1.PutData (254)        'Enter Command Mode
        Form1.PutData (50)  'Timer/Setup Branch Commands
        Form1.PutData (133)        'Get Calibration Value Stored in EEPROM
        LSB = Form1.GetData
        MSB = Form1.GetData
        Slide = (LSB + (MSB * 256)) / 2
        If Slide < TimeCalSlider.Min Then
            TimeCalSlider.Value = TimeCalSlider.Min
        Else
            TimeCalSlider.Value = Slide
        End If
End Sub
Private Sub StoreCalibrationValue_Click()
        Temp = TimeCalSlider.Value
        DAT = Temp * 2
        LSB = (DAT And 255)
        MSB = (DAT And 65280) / 256
        Debug.Print LSB
        Debug.Print MSB
        Form1.PutData (254)        'Enter Command Mode
        Form1.PutData (50)  'Timer/Setup Branch Commands
        Form1.PutData (132)        'Change Calibration Value Stored in EEPROM
        Form1.PutData (LSB)        'LSB
        Form1.PutData (MSB)        'MSB
        Form1.GetData
End Sub
