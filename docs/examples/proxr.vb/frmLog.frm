VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Log Window"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReceive 
      Caption         =   "Receive"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtSendData 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "254 33"
      Top             =   2760
      Width           =   4815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hex"
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Dec"
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   2
      Top             =   840
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public proxrObj As ncdproxr.ProXR

Public Sub OnDataSent(data As Integer)
' show data here
    txtLog.Text = txtLog.Text + "Send: " + GetShowString(data) + vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.SelLength = 0
End Sub

Public Sub OnDataReceived(data As Integer)
    txtLog.Text = txtLog.Text + "Rec:  " + GetShowString(data) + vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.SelLength = 0
End Sub

Private Sub cmdClear_Click()
    txtLog.Text = ""
End Sub

' convert the data to correct format
Private Function GetShowString(data As Integer) As String
    If data = -1 Then
        GetShowString = ""
        Exit Function
    End If
    If Option1(0).Value Then    ' decmal format
        GetShowString = str(data)
    Else        ' hex format
        GetShowString = Hex(data)
    End If
End Function

Private Sub cmdReceive_Click()
    If Not proxrObj Is Nothing Then
        proxrObj.GetData
    End If
End Sub

' return -1 for wrong string
Private Function GetSendValue(ByVal str As String) As Integer
    
    If Option1(0).Value Then    ' decmal format
        GetSendValue = Val(str)
    Else        ' hex format
        GetSendValue = Val("&H" & str & "&")
    End If
    If GetSendValue > 255 Or GetSendValue < 0 Then
        GetSendValue = -1
    End If
End Function

Private Sub cmdSend_Click()
On Error Resume Next
    If Trim(txtSendData.Text) = "" Then Exit Sub

    If Not proxrObj Is Nothing Then
        s = Split(txtSendData.Text)
        For i = 0 To UBound(s)
            sv = GetSendValue(s(i))
            If sv <> -1 Then
                proxrObj.SendData sv
            End If
        Next
    End If
End Sub

