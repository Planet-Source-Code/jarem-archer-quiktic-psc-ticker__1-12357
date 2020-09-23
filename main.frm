VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "QuikTic: PSC Ticker"
   ClientHeight    =   3090
   ClientLeft      =   10035
   ClientTop       =   5475
   ClientWidth     =   1905
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "main.frx":0C42
   ScaleHeight     =   3090
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUpdate 
      Interval        =   60000
      Left            =   2190
      Top             =   930
   End
   Begin VB.Timer MouseCheck 
      Interval        =   1000
      Left            =   2160
      Top             =   1500
   End
   Begin VB.PictureBox Picture1 
      Height          =   3090
      Left            =   0
      ScaleHeight     =   3030
      ScaleWidth      =   1845
      TabIndex        =   0
      Top             =   0
      Width           =   1905
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   3
         Top             =   2775
         Width           =   1185
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1230
         TabIndex        =   2
         Top             =   2775
         Width           =   585
      End
      Begin VB.Frame Frame1 
         Caption         =   "QuikTic: PSC Ticker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Left            =   15
         TabIndex        =   1
         Top             =   0
         Width           =   1800
         Begin QuikTic.Ticker Ticker 
            Height          =   2520
            Left            =   60
            TabIndex        =   4
            Top             =   195
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   4445
         End
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UpdateCount As Integer
Private Sub cmdSearch_Click()
ShellExecute hwnd, "open", "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=" & txtSearch & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=1&optSort=Alphabetical", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Form_Load()
Me.Hide
ShowIcon Me
UpdateTime = 5
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        ' Put your code here to emulate systray icons events. Note: If there is any
        ' code for MouseMove, MouseDown, or MouseUp then a Double Click won't be
        ' caught.
        ' Only uncomment the events that your app will use, so as to avoid any
        ' strange errors.
        Dim TempX
        Dim TempY
        
      If RunningInTray Then
        Select Case x
            Case 7680   ' MouseMove
            If Menu = False And IsConnected Then
                SysIcon.szTip = ""
                Me.WindowState = vbNormal   ' Or vbMaximized if you feel like it.
                If Me.Visible = False Then txtSearch.Text = ""
                Me.Show
                AlwaysOnTop Me, True
            ElseIf IsConnected = False Then
                SysIcon.szTip = "Not Connected"
            End If
                
                
                
            'Case 7695   ' Left MouseDown
            'Case 7710   ' Left MouseUp
           
            Case 7725   ' Left DoubleClick

            'Case 7740   ' Right MouseDown
            Case 7755   ' Right MouseUp
            Menu = True
            PopupMenu Settings.mnuPop
            'Case 7770   ' Right DoubleClick
            
        End Select
      End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Remove the icon when this form unload. Don't forget to unload this form!
    RemoveIcon Me 'Add your form's name here for the sub to work.
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        ' This code hides the Form and puts the icon in the tray. Feel free to move
        ' it around if you like.
        Me.Hide
        ShowIcon Me
    End If
End Sub

Private Sub MouseCheck_Timer()
Dim Temp
Temp = GetY

If GetX < 645 Then
Me.Hide
End If
If GetY < 250 Then
Me.Hide
End If

End Sub

Private Sub tmrUpdate_Timer()
If UpdateCount = UpdateTime Then
UpdateCount = 0
Ticker.Refresh
Else
UpdateCount = UpdateCount + 1
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtSearch.Text <> "" Then cmdSearch_Click

End Sub
