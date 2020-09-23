VERSION 5.00
Begin VB.Form Settings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   345
   ClientLeft      =   7845
   ClientTop       =   6690
   ClientWidth     =   3270
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   270
      Left            =   2565
      TabIndex        =   3
      Top             =   45
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Height          =   240
      Left            =   1395
      TabIndex        =   0
      Text            =   "5"
      Top             =   45
      Width           =   390
   End
   Begin VB.Label Label2 
      Caption         =   "minute(s)."
      Height          =   210
      Left            =   1830
      TabIndex        =   2
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Auto Update every"
      Height          =   225
      Left            =   15
      TabIndex        =   1
      Top             =   75
      Width           =   1815
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPOP"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Update Now!"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If IsNumeric(Text1) Then
UpdateTime = Text1
Unload Me
Else
Beep
End If
End Sub

Public Sub mnuUpdate_Click()
Main.Ticker.Refresh
Unload Me
'Menu = fales is not added because it first has to check
'     if there is a connection, if there isnt then a message will display
End Sub
Public Sub mnuQuit_Click()
RemoveIcon Me
End
End Sub
Public Sub mnuSettings_Click()
Me.Show
Menu = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> "" Then Command1_Click
End Sub
