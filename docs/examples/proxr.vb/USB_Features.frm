VERSION 5.00
Begin VB.Form USB_Features 
   Caption         =   "USB Special Features"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3135
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton AdvancedFeatureButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Advanced Feature Settings"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label DID 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"USB_Features.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "USB_Features"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AdvancedFeatureButton_Click()
    AdvancedFeatures.Visible = True
    AdvancedFeatures.ZOrder 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Form1.USBFeatures.Enabled = True
End Sub
