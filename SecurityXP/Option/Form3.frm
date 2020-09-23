VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENTER PASSWORD"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   945
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter old password to proceed. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00303030&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim hint
hint = GetSetting("Human Knowledge", "sec", "password")




If hint = "" Then
Form1.Show
Form3.Visible = False
End If

Form3.Icon = Form1.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = "13" Then
CheckPass1
End If

End Sub


Private Sub CheckPass1()
Dim hintet
Dim hint
hintet = GetSetting("Human Knowledge", "sec", "password")
hint = Decrypt(hintet, "admhk32590254")

If Text1.Text = hint Then
Form1.Show
Form3.Visible = False
Else
MsgBox ("Wrong password!")
End If

End Sub
