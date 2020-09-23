VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl Ticker 
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   LockControls    =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   1980
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H8000000A&
      ForeColor       =   &H00FFFFFF&
      Height          =   2505
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   1650
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      Begin SHDocVwCtl.WebBrowser Web 
         Height          =   4170
         Left            =   -195
         TabIndex        =   1
         Top             =   -1320
         Width           =   2385
         ExtentX         =   4207
         ExtentY         =   7355
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "res://C:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///"
      End
   End
End
Attribute VB_Name = "Ticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MouseText As String
Public Sub Refresh()
'refreshes the page
Web.Refresh
End Sub

Private Sub UserControl_Initialize()
Web.Navigate "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=-1"

End Sub

Private Sub Web_StatusTextChange(ByVal Text As String)
'This just changes the link info for the web downloadbegin to work
If Left(Text, 14) = "Web site found" Or Left(Text, 4) = "Open" Then
IsConnected = True
Menu = False
Exit Sub
End If

If Left(Text, 4) = "http" Then
MouseText = Text
ElseIf Right(Text, 9) = "error.htm" Then 'If there is no connection
IsConnected = False
If Menu = True Then 'If the clicked update
MsgBox "Sorry, Internet Connection is not available", vbInformation
Menu = False
End If
End If
End Sub
Private Sub Web_DownloadBegin()
'Makes it open the link you clicked in your default browser
If Left(MouseText, 4) = "http" Then
Web.Stop
ShellExecute hWnd, "open", MouseText, vbNullString, vbNullString, conSwNormal
End If
End Sub

