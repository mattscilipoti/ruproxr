VERSION 5.00
Begin VB.Form POT 
   Caption         =   "Potentiometer Features"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form3"
   ScaleHeight     =   2235
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Read 
      Caption         =   "Read Stored Powerup Default Value"
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton Store 
      Caption         =   "Store as Powerup Default Value"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select a Output Value for this Potentiometer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   6495
      Begin VB.HScrollBar PotValue 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   5
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5760
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select a Potentiometer to Read or Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CheckBox SelectAll 
         Caption         =   "Select All"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.HScrollBar PotSelect 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5760
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "POT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Form1.Visible = True
    Form1.POTFeatures.Enabled = True
End Sub
Private Sub PotSelect_Change()
    Label8.Caption = PotSelect
    If PotSelect > 47 Then
        Store.Enabled = False
        Read.Enabled = False
    Else
        Store.Enabled = True
        Read.Enabled = True
    End If
End Sub
Private Sub PotSelect_Scroll()
    Label8.Caption = PotSelect
    If PotSelect > 47 Then
        Store.Enabled = False
        Read.Enabled = False
    Else
        Store.Enabled = True
        Read.Enabled = True
    End If
End Sub
Private Sub PotValue_Change()
    Label1.Caption = PotValue
    If SelectAll.Value = False Then
        Form1.PutData (254)
        Form1.PutData (170)
        Form1.PutData (PotSelect)
        Form1.PutData (PotValue)
        Form1.GetData
    Else
        Form1.PutData (254)
        Form1.PutData (171)
        Form1.PutData (PotValue)
        Form1.GetData
    End If
End Sub
Private Sub PotValue_Scroll()
    Label1.Caption = PotValue
    If SelectAll.Value = False Then
        Form1.PutData (254)
        Form1.PutData (170)
        Form1.PutData (PotSelect)
        Form1.PutData (PotValue)
        Form1.GetData
    Else
        Form1.PutData (254)
        Form1.PutData (171)
        Form1.PutData (PotValue)
        Form1.GetData
    End If
End Sub
Private Sub Read_Click()
    Form1.PutData (254)
    Form1.PutData (173)
    Form1.PutData (PotSelect)
    PotValue = Form1.GetData
End Sub
Private Sub SelectAll_Click()
    If SelectAll.Value = 1 Then
        Store.Enabled = False
        Read.Enabled = False
    Else
        Store.Enabled = True
        Read.Enabled = True
    End If
End Sub
Private Sub Store_Click()
    Form1.PutData (254)
    Form1.PutData (172)
    Form1.PutData (PotSelect)
    Form1.PutData (PotValue)
    Form1.GetData
End Sub
