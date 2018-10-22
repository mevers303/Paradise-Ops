VERSION 5.00
Begin VB.Form frmTray 
   Caption         =   "Tray"
   ClientHeight    =   600
   ClientLeft      =   405
   ClientTop       =   675
   ClientWidth     =   2460
   Icon            =   "frmTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleWidth      =   2460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NotifyIconData
  Size As Long
  handle As Long
  ID As Long
  Flags As Long
  CallBackMessage As Long
  icon As Long
  Tip As String * 64
End Type
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NotifyIconData) As Boolean

Const NIM_ADD = &H0    ' Add an icon
Const NIM_MODIFY = &H1 ' Modify an icon
Const NIM_DELETE = &H2 ' Delete an icon

Const NIF_MESSAGE = &H1       ' To change uCallBackMessage member
Const NIF_ICON = &H2          ' To change the icon
Const NIF_TIP = &H4           ' To change the tooltip text

'nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE


Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201   ' Left click
Const WM_LBUTTONDBLCLK = &H203 ' Left double click
Const WM_RBUTTONDOWN = &H204   ' Right click
Const WM_RBUTTONDBLCLK = &H206 ' Right double click

Dim nid As NotifyIconData
Dim CurrentTooltip As String

Public Function GetTooltip() As String
    GetTooltip = CurrentTooltip
End Function

Public Sub SetTooltip(strTooltip)
    CurrentTooltip = strTooltip
    nid.Tip = strTooltip & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Public Sub SendToTray()
    nid.Size = Len(nid)
    nid.handle = hwnd
    nid.ID = 0
    nid.Flags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.CallBackMessage = WM_MOUSEMOVE
    nid.icon = frmMain.icon
    nid.Tip = CurrentTooltip & vbNullChar
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Msg As Long
    Msg = ScaleX(x, ScaleMode, vbPixels)
    
    Select Case Msg
        Case WM_LBUTTONDBLCLK
            frmMain.WindowState = vbNormal
            frmMain.Show
            Unload Me
        Case WM_RBUTTONDOWN
            PopupMenu mnuTray
    End Select
End Sub

Private Sub Form_Resize()
    mnuRestore_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuExit_Click()
    Shell_NotifyIcon NIM_DELETE, nid
    Unload frmMain
    Unload Me
    End
End Sub

Public Sub mnuRestore_Click()
    frmMain.WindowState = vbNormal
    frmMain.Show
    Unload Me
End Sub
