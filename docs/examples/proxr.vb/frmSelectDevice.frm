VERSION 5.00
Begin VB.Form frmSelectDevice 
   Caption         =   "Select Device"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   2220
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Other"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Zigbee"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmSelectDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function GetSelect() As Integer
    For i = 0 To 1
        If (Option1(i).Value = True) Then
            SaveSetting "ProXR_V4", "Settings", "DeviceType", i
        End If
    Next
End Function

Private Sub cmdSelect_Click()
    GetSelect
    Hide
End Sub

Private Sub Form_Load()
    Dim s As String
    s = GetSetting("ProXR_V4", "Settings", "DeviceType", 0)
    Dim i As Integer
    i = Val(s)
    Option1(i).Value = True
    Debug.Print s
End Sub
