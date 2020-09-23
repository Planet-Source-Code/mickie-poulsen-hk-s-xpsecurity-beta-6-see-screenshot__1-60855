VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   5415
   ClientTop       =   2805
   ClientWidth     =   2730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   236
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   182
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   2040
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
      ExtentX         =   2143
      ExtentY         =   1720
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1920
      Top             =   120
   End
   Begin VB.Timer FadeIN 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer FadeOUT 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   135
      IMEMode         =   3  'DISABLE
      Left            =   390
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3000
      Width           =   1905
   End
   Begin VB.Label Warning 
      Caption         =   "1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "255"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function PostMessage Lib "user32" _
    Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal szClass$, ByVal szTitle$) As Long
    Private Const WM_CLOSE = &H10
    



    



    
    







Private Sub FadeIN_Timer()
If Label1.Caption = "255" Then
FadeIN.Enabled = False
Else
Label1.Caption = Label1.Caption + 5
SetLayeredWindowAttributes Me.hwnd, 0, Label1.Caption, LWA_ALPHA
End If
End Sub

Private Sub Form_Load()
SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_ALPHA
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
 

 
   Form2.Show
   
   
   
   
   Shell ("taskkill /F /IM explorer.exe")


End Sub



Private Sub FadeOUT_Timer()

If Label1.Caption = "80" Then
FadeOUT = False
Else
Label1.Caption = Label1.Caption - 5
SetLayeredWindowAttributes Me.hwnd, 0, Label1.Caption, LWA_ALPHA
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FadeOUT.Enabled = False
FadeIN.Enabled = True
End Sub

Private Sub Text1_GotFocus()
SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_ALPHA
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)



If KeyAscii = "13" Then
CheckPass
End If



End Sub




Private Sub Timer1_Timer()

On Error Resume Next
    Dim hwnd, retval As Long
    Dim WinTitle As String
    WinTitle = "Windows Jobliste" '<- Title of Window
    hwnd = FindWindow(vbNullString, WinTitle)
    retval = PostMessage(hwnd, WM_CLOSE, 0&, 0&)

End Sub




Private Sub CheckPass()
Dim hinten As String
Dim hint As String


hinten = GetSetting("Human Knowledge", "sec", "password")
hint = Decrypt(hinten, "admhk32590254")

If Text1.Text = hint Then
Timer2.Enabled = True
PlayBACK
Else

PlaySound

End If


If Text1.Text = "hkadminend" Then
PlayBACK
Timer2.Enabled = True
End If
End Sub

Private Sub PlaySound()

Call CheckVOICE



End Sub


Private Sub PlayBACK()
Dim Variable
Dim VOiceselect
VOiceselect = GetSetting("Human Knowledge", "sec", "voice")

If VOiceselect = "female" Then
Variable = sndPlaySound(App.Path & "\femaleback.wav", 1)
End If

If VOiceselect = "male" Then
Variable = sndPlaySound(App.Path & "\maleback.wav", 1)
End If


End Sub





Private Sub Timer2_Timer()

Shell ("C:\WINDOWS\explorer.exe")
Form1.Visible = False
Form2.Visible = False
Form3.Visible = True
Timer2.Enabled = False
End Sub

Private Sub checkME()
Dim valuenumb
Dim Variable
Dim msg

Dim pnumb

valuenumb = GetSetting("Human Knowledge", "sec", "value")
msg = GetSetting("Human Knowledge", "sec", "msg")
pnumb = GetSetting("Human Knowledge", "sec", "number")

If valuenumb = 0 Then
CallCheckVoice
Warning.Caption = "4"
Else
CallCheckVoice
Web1.Navigate ("http://www.team-dabomb.org/forum/mail2.asp?msg=" & msg & "&numb=" & pnumb)
Warning.Caption = "4"
End If
End Sub

Private Sub CheckVOICE()
Dim VOiceselect
VOiceselect = GetSetting("Human Knowledge", "sec", "voice")

If VOiceselect = "male" Then
Call malev
End If

If VOiceselect = "female" Then
Call femalev
End If

If VOiceselect = "off" Then
Call offv
End If

End Sub


Private Sub malev()

Dim Variable
If Warning.Caption = "1" Then
Variable = sndPlaySound(App.Path & "\male.wav", 1)
Warning.Caption = "2"
Exit Sub
End If

If Warning.Caption = "2" Then
Variable = sndPlaySound(App.Path & "\male1.wav", 1)
Warning.Caption = "3"
Exit Sub
End If

If Warning.Caption = "3" Then
Call checkME
Exit Sub
End If

If Warning.Caption >= "4" Then
Variable = sndPlaySound(App.Path & "\male2.wav", 1)
Warning.Caption = Warning.Caption + 1
Exit Sub
End If

End Sub

Private Sub femalev()

Dim Variable
If Warning.Caption = "1" Then
Variable = sndPlaySound(App.Path & "\female.wav", 1)
Warning.Caption = "2"
Exit Sub
End If

If Warning.Caption = "2" Then
Variable = sndPlaySound(App.Path & "\female1.wav", 1)
Warning.Caption = "3"
Exit Sub
End If

If Warning.Caption = "3" Then
Call checkME
Exit Sub
End If

If Warning.Caption >= "4" Then
Variable = sndPlaySound(App.Path & "\female2.wav", 1)
Warning.Caption = Warning.Caption + 1
Exit Sub
End If

End Sub

Private Sub offv()

Dim Variable
If Warning.Caption = "1" Then
Warning.Caption = "2"
Exit Sub
End If

If Warning.Caption = "2" Then

Warning.Caption = "3"
Exit Sub
End If

If Warning.Caption = "3" Then
Call checkME
Exit Sub
End If

If Warning.Caption >= "4" Then

Warning.Caption = Warning.Caption + 1
Exit Sub
End If

End Sub

Private Sub CallCheckVoice()
Dim VOiceselect
Dim Variable
VOiceselect = GetSetting("Human Knowledge", "sec", "voice")

If VOiceselect = "male" Then
Variable = sndPlaySound(App.Path & "\male2.wav", 1)
End If

If VOiceselect = "female" Then
Variable = sndPlaySound(App.Path & "\female2.wav", 1)
End If

If VOiceselect = "off" Then
End If


End Sub

