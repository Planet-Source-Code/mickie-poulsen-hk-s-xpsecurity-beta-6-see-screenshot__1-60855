VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   7665
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Off"
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
      Left            =   2280
      TabIndex        =   15
      Top             =   3000
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Female"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   3000
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Male"
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
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE SETTINGS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   7080
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   " E-Mail Alarm "
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
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   3855
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable e-mail alam."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00303030&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00303030&
         Height          =   255
         Left            =   140
         TabIndex        =   11
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00303030&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3255
      End
   End
   Begin VB.TextBox Text2 
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
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
   End
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
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Voice Alarm:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "False"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Password:"
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
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      TabIndex        =   3
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text3.Enabled = True
Text4.Enabled = True
Else
Text3.Enabled = False
Text4.Enabled = False
End If
End Sub

Private Sub Command1_Click()
If Text1.Text = Text2.Text Then
Label4.Caption = "true"
Else
MsgBox ("Passwords are not identical!")
Exit Sub
End If

If Text1.Text = "" Then
MsgBox ("You MUST enter a password!")
Else
MsgBox ("Password saved")
EnCryPteT = Encrypt(Text1.Text, "admhk32590254")
SaveSetting "Human Knowledge", "sec", "password", EnCryPteT

SaveSetting "Human Knowledge", "sec", "value", Check1.Value
Call CheckVOICE
Call CheckCheck
End If



End Sub


Private Sub REGfun()
check = GetSetting("Human Knowledge", "sec", "licens")
If check = "" Then
Form2.Show
Else
SaveSetting "Human Knowledge", "sec", "msg", Text4.Text
SaveSetting "Human Knowledge", "sec", "number", Text3.Text
SaveSetting "Human Knowledge", "sec", "value", Check1.Value

MsgBox ("E-mail Alarm enabled")
End
End If
End Sub


Private Sub CheckCheck()

If Check1.Value = "1" Then
Call REGfun
Else
End
End If

End Sub

Private Sub Form_Load()
Text4.Text = GetSetting("Human knowledge", "sec", "msg")
Text3.Text = GetSetting("Human knowledge", "sec", "number")
checkme = GetSetting("Human knowledge", "sec", "value")
Call CheckIT


If checkme = 1 Then
Check1.Value = "1"
Else
Text3.Enabled = False
Text4.Enabled = False
End If
End Sub


Private Sub CheckVOICE()
If Option1.Value = True Then
SaveSetting "Human Knowledge", "sec", "voice", "male"
End If
If Option2.Value = True Then
SaveSetting "Human Knowledge", "sec", "voice", "female"
End If
If Option3.Value = True Then
SaveSetting "Human Knowledge", "sec", "voice", "off"
End If
End Sub

Private Sub CheckIT()

checkvoice1 = GetSetting("Human knowledge", "sec", "voice", "off")


If checkvoice1 = "male" Then
Option1.Value = True
End If

If checkvoice1 = "female" Then
Option2.Value = True
End If

If checkvoice1 = "off" Then
Option3.Value = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

