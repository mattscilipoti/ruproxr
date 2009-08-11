VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "The Commands on this Page Allow you to Specify the Relay Bank AND Command at the Same Time"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form2"
   ScaleHeight     =   5625
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame19 
      Caption         =   "Store Current Relay Settings as Powerup Default"
      Height          =   2415
      Left            =   7440
      TabIndex        =   56
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton GetBank 
         Caption         =   "Get Stored Powerup Default Relay Pattern"
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   720
         Width           =   4575
      End
      Begin VB.CommandButton StoreBank 
         Caption         =   "Store Current Relay Pattern as Powerup Default"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label5 
         Caption         =   $"Form2.frx":0000
         Height          =   975
         Left            =   120
         TabIndex        =   59
         Top             =   1200
         Width           =   4575
      End
   End
   Begin VB.CommandButton Reverse 
      Caption         =   "Reverse On/Off Relay Pattern"
      Height          =   495
      Left            =   6360
      TabIndex        =   52
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Invert 
      Caption         =   "Invert Status of All Relays"
      Height          =   495
      Left            =   7440
      TabIndex        =   50
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton AllOn 
      Caption         =   "All Relays On"
      Height          =   495
      Left            =   9960
      TabIndex        =   49
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton AllOff 
      Caption         =   "All Relays Off"
      Height          =   495
      Left            =   7440
      TabIndex        =   48
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Frame Frame17 
      Caption         =   "Set Status of All Relays at Once"
      Height          =   615
      Left            =   120
      TabIndex        =   45
      Top             =   3840
      Width           =   6135
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   46
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   5400
         TabIndex        =   47
         Top             =   240
         Width           =   615
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
      TabIndex        =   20
      Top             =   2280
      Width           =   7215
      Begin VB.CommandButton Read_All 
         Caption         =   "Read Status of All Relays"
         Height          =   255
         Left            =   4440
         TabIndex        =   54
         Top             =   0
         Width           =   2535
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   7
         Left            =   6120
         TabIndex        =   21
         Top             =   1080
         Width           =   855
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
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   27
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   30
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   33
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   35
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Read_Relay 
         Caption         =   "Read"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 1"
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   42
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
            TabIndex        =   43
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Relay 2"
         Height          =   855
         Left            =   1080
         TabIndex        =   40
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
            TabIndex        =   41
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 3"
         Height          =   855
         Index           =   1
         Left            =   1920
         TabIndex        =   38
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
            TabIndex        =   39
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 4"
         Height          =   855
         Index           =   2
         Left            =   2760
         TabIndex        =   36
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
            TabIndex        =   37
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 5"
         Height          =   855
         Index           =   3
         Left            =   3600
         TabIndex        =   31
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
            TabIndex        =   32
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 6"
         Height          =   855
         Index           =   4
         Left            =   4440
         TabIndex        =   28
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
            TabIndex        =   29
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 7"
         Height          =   855
         Index           =   5
         Left            =   5280
         TabIndex        =   25
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
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Relay 8"
         Height          =   855
         Index           =   6
         Left            =   6120
         TabIndex        =   22
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
            TabIndex        =   23
            Top             =   360
            Width           =   615
         End
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
      TabIndex        =   3
      Top             =   840
      Width           =   7215
      Begin VB.Frame Frame3 
         Caption         =   "Relay 1"
         Height          =   855
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Relay 2"
         Height          =   855
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Relay 3"
         Height          =   855
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Relay 4"
         Height          =   855
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Relay 5"
         Height          =   855
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Relay 6"
         Height          =   855
         Left            =   4440
         TabIndex        =   8
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Relay 7"
         Height          =   855
         Left            =   5280
         TabIndex        =   6
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Relay 8"
         Height          =   855
         Left            =   6120
         TabIndex        =   4
         Top             =   360
         Width           =   855
         Begin VB.CommandButton RELAY 
            Caption         =   "OFF"
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "Select a Relay Bank to Control"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.HScrollBar BANK 
         Height          =   375
         Left            =   120
         Max             =   32
         TabIndex        =   1
         Top             =   240
         Value           =   26
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
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"Form2.frx":0129
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   55
      Top             =   4560
      Width           =   12135
   End
   Begin VB.Label Label4 
      Caption         =   "Relay On/Off Pattern is Reversed, Relay 12345678 status is coppied to Relay 87654321"
      Height          =   495
      Left            =   7920
      TabIndex        =   53
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Relays that are on turn off.  Relays that are off turn on."
      Height          =   495
      Left            =   8760
      TabIndex        =   51
      Top             =   3240
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AllOff_Click()
    Form1.PutData (254)        'Enter Command Mode
    Form1.PutData (129)        'Send Command Turn All Relays Off
    Form1.PutData (BANK)       'Specify a Relay Bank
    Form1.GetData
End Sub
Private Sub AllOn_Click()
    Form1.PutData (254)        'Enter Command Mode
    Form1.PutData (130)        'Send Command Turn All Relays Off
    Form1.PutData (BANK)       'Specify a Relay Bank
    Form1.GetData
End Sub
Private Sub BANK_Change()
    Form1.HScroll4 = BANK
    Label1(8).Caption = BANK.Value
    If BANK.Value = 0 Then Label1(8).Caption = "0 (All Banks)"
End Sub
Private Sub BANK_Scroll()
    Form1.HScroll4 = BANK
    Label1(8).Caption = BANK.Value
    If BANK.Value = 0 Then Label1(8).Caption = "0 (All Banks)"
End Sub
Private Sub GetBank_Click()
    Form1.PutData (254)        'Enter Command Mode
    Form1.PutData (143)        'Store Current Memory Bank as Power Up Defaults
    Form1.PutData (BANK)       'Specify a Relay Bank
    If BANK > 0 Then
        Temp = Form1.GetData
        Form1.Banks(BANK - 1).BackColor = &HFF&
        Form1.Banks(BANK - 1).Caption = Str$(BANK) + ":" + Str$(Temp) + ":" + Form1.BIN$(Temp)
    Else
        For N = 0 To 31
            Temp = Form1.GetData
            Form1.Banks(N).BackColor = &HFF&
            Form1.Banks(N).Caption = Str$(N + 1) + ":" + Str$(Temp) + ":" + Form1.BIN$(Temp)
        Next N
    End If
End Sub
Private Sub HScroll1_Change()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (140)              'Set Status of All Relays Command
    Form1.PutData (HScroll1.Value)   'Set Status of All Relays Command
    Form1.PutData (BANK)             'Specify a Relay Bank
    Form1.GetData
    Label6.Caption = HScroll1
End Sub
Private Sub HScroll1_Scroll()
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (140)              'Set Status of All Relays Command
    Form1.PutData (HScroll1.Value)   'Set Status of All Relays Command
    Form1.PutData (BANK)             'Specify a Relay Bank
    Form1.GetData
    Label6.Caption = HScroll1
End Sub
Private Sub Invert_Click()
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (131)     'Send Command Invert Relay Status
    Form1.PutData (BANK)   'Specify a Relay Bank
    Form1.GetData
End Sub

Private Sub Read_All_Click()
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (124)    'Send Command to Read the Status of All Relays
    Form1.PutData (BANK)   'Specify a Relay Bank
    Temp = Form1.GetData                'Read Status from Relay Board
        
    'Set All to OFF
    For nn = 0 To 7
        Label1(nn).Caption = "OFF"
        RELAY(nn).Caption = "OFF"
    Next nn
    
    'Determine Which Relays are ON
    If (Temp And 1) = 1 Then Label1(0).Caption = "ON": RELAY(0).Caption = "ON"
    If (Temp And 2) = 2 Then Label1(1).Caption = "ON": RELAY(1).Caption = "ON"
    If (Temp And 4) = 4 Then Label1(2).Caption = "ON": RELAY(2).Caption = "ON"
    If (Temp And 8) = 8 Then Label1(3).Caption = "ON": RELAY(3).Caption = "ON"
    If (Temp And 16) = 16 Then Label1(4).Caption = "ON": RELAY(4).Caption = "ON"
    If (Temp And 32) = 32 Then Label1(5).Caption = "ON": RELAY(5).Caption = "ON"
    If (Temp And 64) = 64 Then Label1(6).Caption = "ON": RELAY(6).Caption = "ON"
    If (Temp And 128) = 128 Then Label1(7).Caption = "ON": RELAY(7).Caption = "ON"
End Sub

Private Sub Read_Relay_Click(Index As Integer)
    Form1.PutData (254)              'Enter Command Mode
    Form1.PutData (116 + Index)      'Send Command to Read the Status of a Relay (1-8)
    Form1.PutData (BANK)             'Specify a Relay Bank
    Temp = Form1.GetData                          'Read Status from Relay Board
    If Temp = 0 Then
        Label1(Index).Caption = "OFF"
    Else
        Label1(Index).Caption = "ON"
    End If
End Sub
Private Sub RELAY_Click(Index As Integer)
    If RELAY(Index).Caption = "ON" Then
        RELAY(Index).Caption = "OFF"
        Form1.PutData (254)            'Command Mode
        Form1.PutData (100 + Index)    'Relay Control Command
        Form1.PutData (BANK)           'Specify a Relay Bank
    Else
        RELAY(Index).Caption = "ON"
        Form1.PutData (254)            'Command Mode
        Form1.PutData (108 + Index)    'Relay Control Command
        Form1.PutData (BANK)           'Specify a Relay Bank
    End If
    Form1.GetData
End Sub
Private Sub Reverse_Click()
    Form1.PutData (254)    'Enter Command Mode
    Form1.PutData (132)     'Send Command to Reverse Relay Pattern
    Form1.PutData (BANK)   'Specify a Relay Bank
    Form1.GetData
End Sub
Private Sub StoreBank_Click()
    Form1.PutData (254)        'Enter Command Mode
    Form1.PutData (142)        'Store Current Memory Bank as Power Up Defaults
    Form1.PutData (BANK)       'Specify a Relay Bank
    Form1.GetData
    If BANK <> 0 Then
        Form1.Banks(BANK - 1).BackColor = &HFF&
        'Banks(HScroll4 - 1).Caption = Str$(HScroll4) + ":" + Str$(Temp) + ":" + BIN$(Temp)
    Else
        For N = 0 To 31
            Form1.Banks(N).BackColor = &HFF&
        Next N
    End If
End Sub
