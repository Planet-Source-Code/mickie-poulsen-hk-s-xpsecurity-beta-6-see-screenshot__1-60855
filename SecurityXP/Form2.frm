VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "HK.SECURITY BETA. BUILD 189"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   -1200
      TabIndex        =   1
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROFESSIONAL EDITION. LICENSED TO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   -1560
      TabIndex        =   0
      Top             =   1680
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim hint
Dim streng As String


Label1.Left = Screen.Width - Label1.Width - 220
Label1.Top = Screen.Height - Label1.Height - 200

Label2.Left = Screen.Width - Label2.Width - 220
Label2.Top = Screen.Height - Label2.Height - 440

hint = GetSetting("Human Knowledge", "sec", "licens")
If hint = "" Then
Label1.Caption = "PRO EDITION. UNLICENSED"
Else
Label1.Caption = Label1.Caption & " " & hint
Label1.Caption = UCase(Label1.Caption)
End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.FadeOUT.Enabled = True
Form1.FadeIN.Enabled = False

Form1.FadeOUT.Enabled = True

End Sub

