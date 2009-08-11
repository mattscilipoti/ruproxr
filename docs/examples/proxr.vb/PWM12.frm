VERSION 5.00
Begin VB.Form PWM12 
   Caption         =   "PWM12"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form3"
   ScaleHeight     =   6135
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Chase FX"
      Height          =   495
      Left            =   6840
      TabIndex        =   53
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame LightingEffects 
      Caption         =   "Lighting Effects"
      Height          =   5895
      Left            =   5520
      TabIndex        =   42
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton Command12 
         Caption         =   "Soft Random"
         Height          =   495
         Left            =   4920
         TabIndex        =   66
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Outer Lights"
         Height          =   495
         Left            =   3720
         TabIndex        =   65
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Inner Lights"
         Height          =   495
         Left            =   2520
         TabIndex        =   64
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Random"
         Height          =   495
         Left            =   4920
         TabIndex        =   63
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Three-Step"
         Height          =   495
         Left            =   1320
         TabIndex        =   62
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Two-Step"
         Height          =   495
         Left            =   120
         TabIndex        =   61
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cylon"
         Height          =   495
         Left            =   3720
         TabIndex        =   60
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Chase FX2"
         Height          =   495
         Left            =   2520
         TabIndex        =   59
         Top             =   3840
         Width           =   1095
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   56
         Top             =   3480
         Value           =   8
         Width           =   4815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stop FX"
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   3840
         Width           =   1095
      End
      Begin VB.HScrollBar FXSpeed 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   54
         Top             =   3120
         Value           =   128
         Width           =   4815
      End
      Begin VB.Frame Frame4 
         Caption         =   "Starlight: Select a Bright Spot (with speed control)"
         Height          =   975
         Left            =   120
         TabIndex        =   50
         Top             =   2040
         Width           =   5895
         Begin VB.HScrollBar HScroll5 
            Height          =   255
            Left            =   120
            Max             =   11
            TabIndex        =   52
            Top             =   240
            Value           =   1
            Width           =   5655
         End
         Begin VB.HScrollBar HScroll4 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   51
            Top             =   600
            Value           =   128
            Width           =   5655
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Starlight: Select a Bright Spot (with speed control)"
         Height          =   975
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   5895
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   49
            Top             =   600
            Value           =   128
            Width           =   5655
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   120
            Max             =   11
            TabIndex        =   48
            Top             =   240
            Value           =   1
            Width           =   5655
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Starlight: Select a Bright Spot"
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   5895
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   120
            Max             =   11
            TabIndex        =   46
            Top             =   240
            Value           =   1
            Width           =   5655
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "All On"
         Height          =   495
         Left            =   4920
         TabIndex        =   44
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "All Off"
         Height          =   495
         Left            =   3720
         TabIndex        =   43
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Dim Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "FX Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   3120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Pulse Width Modulation Values for Each Output Channel"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   12
         Left            =   1320
         Max             =   255
         TabIndex        =   33
         Top             =   4680
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   11
         Left            =   1320
         Max             =   255
         TabIndex        =   32
         Top             =   4320
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   10
         Left            =   1320
         Max             =   255
         TabIndex        =   31
         Top             =   3960
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   9
         Left            =   1320
         Max             =   255
         TabIndex        =   30
         Top             =   3600
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   0
         Left            =   1320
         Max             =   255
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   1
         Left            =   1320
         Max             =   255
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   2
         Left            =   1320
         Max             =   255
         TabIndex        =   9
         Top             =   1080
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   3
         Left            =   1320
         Max             =   255
         TabIndex        =   8
         Top             =   1440
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   4
         Left            =   1320
         Max             =   255
         TabIndex        =   7
         Top             =   1800
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   5
         Left            =   1320
         Max             =   255
         TabIndex        =   6
         Top             =   2160
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   6
         Left            =   1320
         Max             =   255
         TabIndex        =   5
         Top             =   2520
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   7
         Left            =   1320
         Max             =   255
         TabIndex        =   4
         Top             =   2880
         Width           =   3015
      End
      Begin VB.HScrollBar PWM 
         Height          =   255
         Index           =   8
         Left            =   1320
         Max             =   255
         TabIndex        =   3
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Frame Frame21 
         Caption         =   "Powerup Default Settings"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   5040
         Width           =   5055
         Begin VB.CommandButton StorePattern 
            Caption         =   "Store Current Settings as Powerup Default"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 12"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 11"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 10"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 9"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3600
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
         Index           =   12
         Left            =   4440
         TabIndex        =   37
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
         Index           =   11
         Left            =   4440
         TabIndex        =   36
         Top             =   4320
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
         TabIndex        =   35
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
         Index           =   9
         Left            =   4440
         TabIndex        =   34
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Set All Channels"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 1"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 2"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 3"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 4"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 5"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 6"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 7"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Channel 8"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3240
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
         Index           =   0
         Left            =   4440
         TabIndex        =   20
         Top             =   360
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
         TabIndex        =   19
         Top             =   720
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
         TabIndex        =   18
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
         Index           =   3
         Left            =   4440
         TabIndex        =   17
         Top             =   1440
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   2160
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
         TabIndex        =   14
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
         Index           =   7
         Left            =   4440
         TabIndex        =   13
         Top             =   2880
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
         TabIndex        =   12
         Top             =   3240
         Width           =   615
      End
   End
End
Attribute VB_Name = "PWM12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (0)                'Select a Channel 0 = All or 1-8
    Form1.PutData (0)     'Set Channel to Value 0-255
    Form1.GetData
End Sub


Private Sub Command10_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (157)                  'Lighting Effects Command
    Form1.PutData (FXSpeed)              'Set Speed of Lighting Effects
    Form1.PutData (HScroll6)             'Set Speed of Dimmers
    Form1.GetData
End Sub

Private Sub Command11_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (158)                  'Lighting Effects Command
    Form1.PutData (FXSpeed)              'Set Speed of Lighting Effects
    Form1.PutData (HScroll6)             'Set Speed of Dimmers
    Form1.GetData
End Sub

Private Sub Command12_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (159)                  'Lighting Effects Command
    Form1.PutData (FXSpeed)              'Set Speed of Lighting Effects
    Form1.PutData (HScroll6)             'Set Speed of Dimmers
    Form1.GetData
End Sub

Private Sub Command2_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (0)                'Select a Channel 0 = All or 1-8
    Form1.PutData (255)     'Set Channel to Value 0-255
    Form1.GetData
End Sub
Private Sub Command3_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (151)                  'Lighting Effects Command
    Form1.PutData (FXSpeed)              'Set Speed of Lighting Effects
    Form1.PutData (HScroll6)             'Set Speed of Dimmers
    Form1.GetData
End Sub
Private Sub Command4_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (150)                  'Lighting Effects Command
    Form1.GetData
End Sub

Private Sub Command5_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (152)                  'Lighting Effects Command
    Form1.PutData (FXSpeed)              'Set Speed of Lighting Effects
    Form1.PutData (HScroll6)             'Set Speed of Dimmers
    Form1.GetData
End Sub

Private Sub Command6_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (153)                  'Lighting Effects Command
    Form1.PutData (FXSpeed)              'Set Speed of Lighting Effects
    Form1.PutData (HScroll6)             'Set Speed of Dimmers
    Form1.GetData
End Sub

Private Sub Command7_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (154)                  'Lighting Effects Command
    Form1.PutData (FXSpeed)              'Set Speed of Lighting Effects
    Form1.PutData (HScroll6)             'Set Speed of Dimmers
    Form1.GetData
End Sub

Private Sub Command8_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (155)                  'Lighting Effects Command
    Form1.PutData (FXSpeed)              'Set Speed of Lighting Effects
    Form1.PutData (HScroll6)             'Set Speed of Dimmers
    Form1.GetData
End Sub

Private Sub Command9_Click()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (156)                  'Lighting Effects Command
    Form1.PutData (FXSpeed)              'Set Speed of Lighting Effects
    Form1.PutData (HScroll6)             'Set Speed of Dimmers
    Form1.GetData
End Sub



Private Sub HScroll1_Change()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (100)                  'Lighting Effects Command
    Form1.PutData (HScroll1.Value)       'Select a Brights Spot
    Form1.GetData
End Sub
Private Sub HScroll1_Scroll()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (100)                  'Lighting Effects Command
    Form1.PutData (HScroll1.Value)       'Select a Brights Spot
    Form1.GetData
End Sub
Private Sub HScroll2_Change()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (101)                  'Lighting Effects Command
    Form1.PutData (HScroll2.Value)       'Spot
    Form1.PutData (HScroll3.Value)       'Speed
    Form1.GetData
End Sub
Private Sub HScroll2_Scroll()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (101)                  'Lighting Effects Command
    Form1.PutData (HScroll2.Value)       'Spot
    Form1.PutData (HScroll3.Value)       'Speed
    Form1.GetData
End Sub
Private Sub HScroll5_Change()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (102)                  'Lighting Effects Command
    Form1.PutData (HScroll5.Value)       'Spot
    Form1.PutData (HScroll4.Value)       'Speed
    Form1.GetData
End Sub
Private Sub HScroll5_Scroll()
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (102)                  'Lighting Effects Command
    Form1.PutData (HScroll5.Value)       'Spot
    Form1.PutData (HScroll4.Value)       'Speed
    Form1.GetData
End Sub

Private Sub PWM_Change(Index As Integer)
    Form1.ProXR1.ClearBuffer
    PWMValue(Index) = PWM(Index).Value
    'Debug.Print Index
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (Index)                'Select a Channel 0 = All or 1-8
    Form1.PutData (PWM(Index).Value)     'Set Channel to Value 0-255
    Form1.GetData
End Sub
Private Sub PWM_Scroll(Index As Integer)
    Form1.ProXR1.ClearBuffer
    PWMValue(Index) = PWM(Index).Value
    'Debug.Print Index
    Form1.PutData (253)                  'Enter Command Mode
    Form1.PutData (Index)                'Select a Channel 0 = All or 1-8
    Form1.PutData (PWM(Index).Value)     'Set Channel to Value 0-255
    Form1.GetData
End Sub
