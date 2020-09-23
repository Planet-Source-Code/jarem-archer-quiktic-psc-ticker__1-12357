Attribute VB_Name = "TrayModule"
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
  
Public SysIcon As NOTIFYICONDATA, RunningInTray As Boolean

Public Sub ShowIcon(ByRef TrayForm As Form)
    ' Show the systray icon. Use from another form : "SystrayIcon.ShowIcon"
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = TrayForm.hwnd
    SysIcon.uID = vbNull
    SysIcon.uFlags = 7
    SysIcon.uCallbackMessage = 512
    SysIcon.hIcon = TrayForm.Icon
   ' SysIcon.szTip = TrayForm.Caption + Chr(0)
    Shell_NotifyIcon 0, SysIcon
    RunningInTray = True
End Sub

Public Sub RemoveIcon(TrayForm As Form)
    ' Remove the systray icon. Use from another form : "SystrayIcon.RemoveIcon"
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = TrayForm.hwnd
    SysIcon.uID = vbNull
    SysIcon.uFlags = 7
    SysIcon.uCallbackMessage = vbNull
    SysIcon.hIcon = TrayForm.Icon
    SysIcon.szTip = Chr(0)
    Shell_NotifyIcon 2, SysIcon
    RunningInTray = False
End Sub

