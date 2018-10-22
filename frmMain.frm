VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000016&
   Caption         =   "Paradise Ops by Spasm"
   ClientHeight    =   5790
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstRunners 
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstTimeBans 
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSWinsockLib.Winsock sckBNLS 
      Index           =   0
      Left            =   7440
      Tag             =   "WAR3/D2DV"
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrMisc 
      Interval        =   1000
      Left            =   9120
      Top             =   4200
   End
   Begin VB.ListBox lstIPBans 
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstBans 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ListView lstChannel 
      Height          =   3015
      Left            =   7320
      TabIndex        =   6
      Top             =   240
      Width           =   3550
      _ExtentX        =   6271
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   3572
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Ping"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Flags"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.ListBox lstQueue 
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3360
      Width           =   7335
   End
   Begin MSWinsockLib.Winsock sckBNET 
      Index           =   0
      Left            =   6600
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5953
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckBNLS 
      Index           =   1
      Left            =   8160
      Tag             =   "STAR/SEXP/W2BN"
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblBancount 
      Alignment       =   2  'Center
      Caption         =   "Bancount: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   4
      Top             =   3480
      Width           =   3550
   End
   Begin VB.Label lblChannel 
      Alignment       =   2  'Center
      Caption         =   "Disconnected"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   0
      Width           =   3550
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status: Offline"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   3240
      Width           =   3550
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "Connect Next Bot"
   End
   Begin VB.Menu mnuDisconnect 
      Caption         =   "Disconnect All"
   End
   Begin VB.Menu mnuReload 
      Caption         =   "Reload Settings"
   End
   Begin VB.Menu mnuGUI 
      Caption         =   "Disable GUI"
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuConfig 
         Caption         =   "Configuration"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open File"
         Begin VB.Menu mnuOpenFolder 
            Caption         =   "Bot Folder"
         End
         Begin VB.Menu mnuOpenConfig 
            Caption         =   "Configuration"
         End
      End
   End
   Begin VB.Menu test 
      Caption         =   "test"
   End
   Begin VB.Menu mnuChannel 
      Caption         =   "Channel Options"
      Visible         =   0   'False
      Begin VB.Menu mnuKick 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuIPBan 
         Caption         =   "IP Ban"
      End
      Begin VB.Menu mnuShitlist 
         Caption         =   "Shitlist"
      End
      Begin VB.Menu mnuBleh 
         Caption         =   "-----------"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDesignate 
         Caption         =   "Designate"
      End
      Begin VB.Menu mnuOp 
         Caption         =   "Op"
      End
      Begin VB.Menu mnuDDP 
         Caption         =   "DDP"
      End
      Begin VB.Menu mnuSafelist 
         Caption         =   "Safelist"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    ReDim BotVars(0)
    ReDim Shamans(0)
    GUI = True
    LoadSettings
    
    Status = "Offline"
    Channel = "Disconnected"
    
    frmTray.SetTooltip "Paradise Ops - " & Status
       
End Sub

Private Sub Form_Resize()

    txtSend.Top = Me.Height - 1190
    If txtSend.Top > 0 Then rtbText.Height = txtSend.Top
    lblChannel.Left = Me.Width - 3615
    lstChannel.Left = lblChannel.Left
    If rtbText.Height > 390 Then lstChannel.Height = rtbText.Height - 390
    lblStatus.Top = lstChannel.Height + 255
    lblBancount.Top = lblStatus.Top + 255
    lblStatus.Left = lstChannel.Left
    lblBancount.Left = lblStatus.Left
    If lstChannel.Left > 0 Then rtbText.Width = lstChannel.Left
    txtSend.Width = rtbText.Width
    
    If Me.WindowState = vbMinimized Then
        frmTray.SendToTray
        Me.Hide
    End If
    
End Sub

Private Sub lstChannel_DblClick()
    If lstChannel.ListItems.Count > 0 Then txtSend.Text = txtSend.Text & lstChannel.SelectedItem.Text
End Sub

Private Sub lstChannel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstChannel.ListItems.Count > 0 And Button = 2 Then
        PopupMenu mnuChannel
    End If
End Sub

Private Sub mnuBan_Click()
    AddQ "/ban " & lstChannel.SelectedItem.Key, 1
End Sub

Private Sub mnuConfig_Click()
    Load frmConfig
    frmConfig.Show
End Sub

Private Sub mnuConnect_Click()
    Dim i As Integer
    For i = 0 To UBound(BotVars)
        If Not BotVars(i).Connected Then
            LoadBot i
            Exit Sub
        End If
    Next
    i = UBound(BotVars) + 1
    LoadBot i
End Sub

Private Sub mnuDDP_Click()
    ParseCommand "ddp " & lstChannel.SelectedItem.Key, "§", "lol"
End Sub

Private Sub mnuDesignate_Click()
    AddQ "/designate " & lstChannel.SelectedItem.Key, 1
End Sub

Private Sub mnuDisconnect_Click()
    Dim i As Integer
    For i = UBound(BotVars) To 0 Step -1
        Closewinsock i
    Next
End Sub

Private Sub mnuGUI_Click()
    If GUI Then
        DisableGUI
    Else
        EnableGUI
    End If
End Sub

Private Sub mnuIPBan_Click()
    AddQ "/squelch " & lstChannel.SelectedItem.Key, 1
End Sub

Private Sub mnuKick_Click()
    AddQ "/kick " & lstChannel.SelectedItem.Key, 1
End Sub

Private Sub mnuOp_Click()
    AddQ "/designate " & lstChannel.SelectedItem.Key, 1
    AddQ "/rejoin"
End Sub

Private Sub mnuReload_Click()
    LoadSettings
End Sub

Private Sub mnuSafelist_Click()
    Dim strReturn As String
    ParseCommand "safelist " & lstChannel.SelectedItem.Key, "§", strReturn
    AddChat vbYellow, strReturn
End Sub

Private Sub mnuShitlist_Click()
    Dim strReturn As String
    ParseCommand "shitlist " & lstChannel.SelectedItem.Key, "§", strReturn
    AddChat vbYellow, strReturn
End Sub

Private Sub sckBNET_Close(Index As Integer)
    Closewinsock Index
    SetTimer frmMain.hWnd, CLng(Index + 2000), 60000, AddressOf Reconnect
    AddChat vbRed, "[" & Index & "] Will attempt to reconnect in one minute."
End Sub

Private Sub sckBNET_Connect(Index As Integer)
    ' When we called sckBNET.Connect, Winsock attempted to connect to the
    ' Battle.Net server specified in quotes. If it connected successfully,
    ' Winsock will automatically call this function in which we continue to
    ' send the needed packets to Battle.Net
    AddChat vbYellow, "[" & Index & "] Connected"
    
    ' Unless you send this, Battle.Net will not know you are attempting to connect
    ' and will thus not like you sending it random packets
    sckBNET(Index).SendData Chr$(1)
    
    ' Now we'll move on to send packet 0x50, which is our AUTH_INFO packet
    Send0x50 Index
End Sub

Private Sub sckBNET_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    ' Winsock calls this Sub whenever data is recieved from Battle.Net.  This Sub
    ' gets the data and passes it along to our ParseData Sub which in turn will
    ' decide what needs to be done with it\
    On Error GoTo Err
    
    Dim lngLen As Long
    Dim strTemp As String
    
    sckBNET(Index).GetData strTemp, vbString
    
    BotVars(Index).strBuffer = BotVars(Index).strBuffer & strTemp
    
    While Len(BotVars(Index).strBuffer) > 4
        lngLen = PBuffer.GetWORD(Mid(BotVars(Index).strBuffer, 3, 2))
        ParseData Left(BotVars(Index).strBuffer, lngLen), Index
        BotVars(Index).strBuffer = Mid(BotVars(Index).strBuffer, lngLen + 1)
    Wend
    Exit Sub
    
Err:

    If Index < UBound(BotVars) Then
        BotVars(Index).strBuffer = vbNullString
    End If
    
End Sub

Private Sub sckBNET_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'When Winsock encounters an error, it will display it in red
    AddChat vbRed, "[" & Index & "] Socket Error " & Number & ": " & Description
    Closewinsock Index
    SetTimer frmMain.hWnd, CLng(Index + 2000), 60000, AddressOf Reconnect
    AddChat vbRed, "[" & Index & "] Will attempt to reconnect in one minute."
End Sub

Private Sub sckBNLS_Close(Index As Integer)
    AddChat vbRed, " - [BNLS] Disconnected"
    sckBNLS(Index).Close
End Sub

Private Sub sckBNLS_Connect(Index As Integer)
    AddChat vbYellow, " - [BNLS] Connected"
End Sub

Private Sub sckBNLS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Static strBuffer As String
    Dim lngLen As Long
    Dim strTemp As String
    
    sckBNLS(Index).GetData strTemp, vbString
    
    strBuffer = strBuffer & strTemp
    
    While Len(strBuffer) > 3
        lngLen = PBuffer.GetWORD(Left(strBuffer, 2))
        ParseBNLSData Left(strBuffer, lngLen), Index
        strBuffer = Mid(strBuffer, lngLen + 1)
    Wend

End Sub

Private Sub sckBNLS_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'When Winsock encounters an error, it will display it in red
    AddChat vbRed, "[BNLS] Socket Error " & Number & ": " & Description
    AddChat vbRed, "[BNLS] Disconnected"
    sckBNLS(Index).Close
End Sub

Private Sub test_Click()
    Dim x As Integer
    Dim y As Integer
    
    AddChat &HC000C0, "Ban list"
    For x = 0 To lstBans.ListCount
        AddChat &HC000C0, x & ": " & lstBans.List(x)
    Next
    
    AddChat vbWhite, "++++++++++++++++++++++++++++++++++++++"
    
    AddChat &HC000C0, "Userlist"
    For x = 0 To UBound(Userlist)
        AddChat &HC000C0, x & ": " & Userlist(x).Username & " " & Userlist(x).Access
    Next
    
    AddChat vbWhite, "++++++++++++++++++++++++++++++++++++++"
    
    For y = 0 To UBound(BotVars)
        AddChat &HC000C0, "IP Bans (" & BotVars(y).Username & ")"
        For x = 0 To lstIPBans(y).ListCount
            AddChat &HC000C0, x & ": " & lstIPBans(y).List(x)
        Next
        AddChat vbWhite, "-----"
    Next
End Sub

Private Sub tmrMisc_Timer()
    Dim x As Integer
    Static LastIdle As Long
    Static LastClanList As Long
    Dim Tick As Long
    
    Tick = GetTickCount()
    
    If IdleKick Then
        For x = 1 To lstChannel.ListItems.Count
            If Tick - Val(lstChannel.ListItems(x).Tag) >= (IdleKickTime * 1000) Then
                If Not IsSafelisted(lstChannel.ListItems(x).Key) Then AddQ "/kick " & lstChannel.ListItems(x).Key & " Idle for " & IdleKickTime & " seconds"
            End If
        Next
    End If
    
    If Idle Then
        If ((Tick - LastIdle) / 1000) >= IdleTime Then
            AddQ IdleMessage
            LastIdle = Tick
        End If
    End If
    
    For x = 0 To (lstTimeBans.ListCount - 1)
        If lstTimeBans.ItemData(x) < Tick Then
            AddQ "/unban " & lstTimeBans.List(x)
            RemoveIPBan lstTimeBans.List(x)
            lstTimeBans.RemoveItem x
        End If
    Next
    
    If Tick - LastClanList > 60000 Then
        For x = 0 To UBound(BotVars)
            If BotVars(x).Product = "WAR3" And BotVars(x).Connected Then
                PBuffer.InsertDWORD 19
                PBuffer.sendPacket &H7D, x
                LastClanList = Tick
                Exit For
            End If
        Next
    End If
        
End Sub

Private Sub txtSend_Change()
    If InStr(txtSend.Text, vbCrLf) Then txtSend_KeyPress (13)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)

    Dim strTemp() As String
    Dim strReturn As String
    Dim x As Integer
    
    If Left(txtSend.Text, 2) = vbCrLf Then txtSend.Text = Mid(txtSend.Text, 3)
    
    If KeyAscii = 13 Then
        strTemp = Split(txtSend.Text, vbCrLf)
        For x = 0 To UBound(strTemp)
            If strTemp(x) <> vbNullString Then
                If Left(strTemp(x), 1) = "/" Then ParseCommand Mid(strTemp(x), 2), "§", strReturn
                If strReturn = "Spasm is cool." Then
                ElseIf strReturn <> vbNullString Then
                    AddChat vbYellow, strReturn
                Else
                    AddQ strTemp(x)
                End If
            End If
        Next
        txtSend.Text = vbNullString
    End If
    
End Sub
