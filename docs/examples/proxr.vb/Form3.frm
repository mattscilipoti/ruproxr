VERSION 5.00
Object = "{6B4D88AD-6CB8-4A1F-8989-E7B2C6D53220}#4.0#0"; "ProXR.ocx"
Begin VB.Form Form3 
   Caption         =   "Select COM Port"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4380
   LinkTopic       =   "Form3"
   ScaleHeight     =   1410
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin NCDProXR.ProXR ProXR1 
      Left            =   2400
      Top             =   600
      _ExtentX        =   926
      _ExtentY        =   926
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Dim ports As Collection
Set ports = ProXR1.EnumComPorts
Combo1.Clear
For Each obj In ports
Combo1.AddItem obj
Next
If (ports.Count > 0) Then Combo1.Text = ports(1)
End Sub
