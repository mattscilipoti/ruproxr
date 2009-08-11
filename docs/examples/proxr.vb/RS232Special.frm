VERSION 5.00
Begin VB.Form RS232Special 
   Caption         =   "Firmware Selected RS-232 Baud Rate"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form3"
   ScaleHeight     =   2235
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Baud Rate Settings for Runtime Mode"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.ComboBox Baud 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "RS232Special.frx":0000
         Left            =   120
         List            =   "RS232Special.frx":0019
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label ADC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"RS232Special.frx":006F
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   8
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   8895
      End
      Begin VB.Label ADC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"RS232Special.frx":01EE
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   10695
      End
   End
End
Attribute VB_Name = "RS232Special"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Baud_Click()
    Debug.Print "Changing Runtime Baud Rate"
    Debug.Print Baud.Text
    Debug.Print Baud.ListIndex
    
    Debug.Print "---------"
    Form1.PutData (254)                'Command Mode
    Form1.PutData (169)                'Write Baud Command
    Form1.PutData (Baud.ListIndex + 1) 'Baud Data
    RET = Form1.GetData 'Debug.Print "Stored Device";: Form1.GetData     'Echo back
    If RET = 0 Then
       RET = 3
    End If
    If RET - 1 <> 254 Then
    Debug.Print "Stored:"; RET - 1
    Baud.ListIndex = RET - 1
    'Debug.Print Baud.ListIndex
    Debug.Print "---------"
    End If
End Sub

Private Sub Form_Load()
    Form1.GetData
    Debug.Print "---------READING DEFAULT BAUD RATE"
    Form1.PutData (254) 'Command Mode
    Form1.PutData (168) 'Read Baud Command
    RET = Form1.GetData
    Debug.Print "Reading Device Baud Settings"; RET
    Debug.Print "---------END"
    
    If RET = 1 Then
        Baud.SelText = "2400 Baud"
    End If
    If RET = 2 Then
        Baud.SelText = "4800 Baud"
    End If
    If RET = 3 Then
        Baud.SelText = "9600 Baud"
    End If
    If RET = 4 Then
        Baud.SelText = "19.2K Baud"
    End If
    If RET = 5 Then
        Baud.SelText = "38.4K Baud"
    End If
    If RET = 6 Then
        Baud.SelText = "57.6K Baud"
    End If
    If RET = 7 Then
        Baud.SelText = "115.2K Baud"
    End If
    'If Form1.BaudSet(1).Value = True Then
    'End If
    'If Form1.BaudSet(2).Value = True Then
    'End If
    'If Form1.BaudSet(3).Value = True Then
    'End If
    'If Form1.BaudSet(4).Value = True Then
    'End If
    'If Form1.BaudSet(5).Value = True Then
    'End If
    'If Form1.BaudSet(6).Value = True Then
    'End If
    'If Form1.BaudSet(7).Value = True Then
    'End If
    
    If Form1.INTER(1).Caption <> "RS-232" Then
        'Frame20.Visible = False     'Turn Off E3C Programming Option
        'Frame14.Visible = False     'Turn Off E3C Options
        'AdvancedFeatures.Label7.Visible = False 'Turn Off Baud Change Note
        'AdvancedFeatures.ChangeBaud.Visible = False 'Turn Off Baud Change Button
        'AdvancedFeatures.Label9.Visible = False
        'AdvancedFeatures.Label8.Visible = False
        'AdvancedTimers.Label11.Visible = False
        'DID(4).Visible = False
        'Label20.Visible = False
        RS232Special.Frame4.Visible = False
        'For X = 1 To 6
        '    BaudSet(X).Enabled = False
        'Next X
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Form1.RS232Baud.Enabled = True
End Sub
Private Sub Command1_Click()
    Debug.Print "---------"
    Form1.PutData (254) 'Command Mode
    Form1.PutData (168) 'Read Baud Command
    Debug.Print "Stored Device";: Form1.GetData
    Debug.Print "---------"
End Sub
Private Sub Command2_Click()
    Debug.Print "---------"
    Form1.PutData (254) 'Command Mode
    Form1.PutData (169) 'Write Baud Command
    Form1.PutData (5)   'Baud Data
    Debug.Print "Stored Device";: Form1.GetData
    Debug.Print "---------"
End Sub
