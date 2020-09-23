Attribute VB_Name = "Module1"
Option Explicit


Private Declare Function PostMessage Lib "user32" _
    Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal szClass$, ByVal szTitle$) As Long
    Private Const WM_CLOSE = &H10


    Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


    Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

