Attribute VB_Name = "modFunctions"
' I like to keep my misc. functions in their own module. This makes it easier
' to know where to go when I need to edit one of them because they are all
' in this file
Option Explicit

'--------for INI file read/write
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'-------------------
Public Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Public Sub AddChat(ParamArray saElements() As Variant)
    ' This is a function to add text to a Rich Textbox. When calling it, you
    ' use the following syntax: AddChat("Example Text", vbRed) to display the text
    ' Example Text in the RTB red
    If Not GUI Then Exit Sub
    
    Dim Data As String
    Dim newText() As String
    Dim strTimeStamp As String
    
    strTimeStamp = "[" & Format(Time, "c") & "] "
    With frmMain.rtbText
        If Len(.Text) > 10000 Then
            .SelStart = 0
            .SelLength = InStr(5000, .Text, vbCrLf) + 2
            .SelText = vbNullString
        End If
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelColor = vbWhite
        .SelText = strTimeStamp
        .SelStart = Len(.Text)
        Data = strTimeStamp
        Dim i As Byte
        For i = LBound(saElements) To UBound(saElements) Step 2
            .SelStart = Len(.Text)
            .SelLength = 0
            .SelColor = saElements(i)
            .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
            .SelStart = Len(.Text)
            Data = Data & saElements(i + 1)
        Next i
    End With
End Sub

Sub ClearLists()
    
End Sub

Public Sub AddQ(ByVal Message As String, Optional ByVal Priority As Byte = 0, Optional ByVal SpecificBot As Integer = 1337)
    Static Index As Integer
    If Not Connected Then Exit Sub
    If SpecificBot <> 1337 Then Index = SpecificBot  'If a specific bot was specified, switch to that one
Begin:

    DoEvents
    
    If Index > UBound(BotVars) Then Index = 0
    
    If Not BotVars(Index).Connected Then Index = Index + 1: GoTo Begin 'Go to the next one
    
    If Left(Message, 1) = "/" Then
        Select Case Left(Message, 4)
            Case "/ban", "/kic", "/cla", "/des", "/unb", "/rej"
                If Not HasOps Then Exit Sub
                If Not BotVars(Index).HasOps Then Index = Index + 1: GoTo Begin
        End Select
    End If
    
    If Priority = 0 Then  'Normal message
        If frmMain.lstQueue(Index).ListCount = MaxQueue Then Exit Sub
        frmMain.lstQueue(Index).AddItem Message
    Else   'High priority message
        frmMain.lstQueue(Index).AddItem Message, 0
    End If
        
    If Not BotVars(Index).TimerEnabled Then SetTimer frmMain.hWnd, CLng(Index), 50, AddressOf TimerProc
    
    Index = Index + 1
    
End Sub

Public Function RealmUsername(ByVal Username As String, ByVal Index As Integer) As String

    Username = LCase(Username)
    If Not AssumeRealms Then RealmUsername = Username: Exit Function
    
    If BotVars(Index).Product = "WAR3" Then
        If Right(Username, Len(Realm)) <> Realm Then
            If Right(Username, Len(War3Realm)) <> War3Realm Then
                RealmUsername = Username & War3Realm
                Exit Function
            End If
        End If
    Else
        If Right(Username, Len(War3Realm)) <> War3Realm Then
            If Right(Username, Len(Realm)) <> Realm Then
                RealmUsername = Username & Realm
                Exit Function
            End If
        End If
    End If
    
    RealmUsername = Username
    
End Function

Public Function AddRealms(ByVal Message As String, ByVal Index As Integer) As String


    Dim strTemp() As String
    strTemp = Split(Message)
    If UBound(strTemp) < 1 Then AddRealms = Message: Exit Function
    Select Case strTemp(0)
        Case "/ban", "/kick", "/ignore", "/w", "/designate", "/unban", "/squelch", "/m", "/where", "/whois", "/unsquelch", "/unignore"
            Dim Username As String
            Username = LCase(strTemp(1))
            Message = Replace(Message, strTemp(1), Username)
            Select Case BotVars(Index).Product
                Case "WAR3"
                    If Right(Username, Len(War3Realm)) = War3Realm Then
                        Username = Left(Username, Len(Username) - Len(War3Realm))
                    ElseIf Right(Username, Len(Realm)) <> Realm Then
                        Username = Username & Realm
                    End If
                Case "STAR", "W2BN", "D2DV"
                    If Right(Username, Len(Realm)) = Realm Then
                        Username = Left(Username, Len(Username) - Len(Realm))
                    End If
            End Select
            If BotVars(Index).Product = "D2DV" Then
                If Left(Username, 1) <> "*" Then Username = "*" & Username
            End If
            AddRealms = Replace(Message, LCase(strTemp(1)), Username)
        Case Else
            AddRealms = Message
            Exit Function
    End Select

End Function

Public Sub DisableGUI()
    GUI = False
    With frmMain
        .rtbText.SelStart = 0
        .rtbText.SelLength = Len(frmMain.rtbText.Text)
        .rtbText.SelText = vbNullString
        .Width = 6000
        .Height = 840
        .BorderStyle = 1
        .mnuGUI.Caption = "Enable GUI"
        .rtbText.Visible = False
        .lstChannel.Visible = False
        .txtSend.Visible = False
        .lblBancount.Visible = False
        .lblChannel.Visible = False
        .lblStatus.Visible = False
    End With
End Sub

Public Sub EnableGUI()

    GUI = True
    
    With frmMain
        .mnuGUI.Caption = "Disable GUI"
        .BorderStyle = 2
        .Width = 13000
        .Height = 8000
        .rtbText.Visible = True
        .lstChannel.Visible = True
        .txtSend.Visible = True
        .lblBancount.Visible = True
        .lblBancount.Caption = "Bancount: " & frmMain.lstBans.ListCount
        .lblChannel.Visible = True
        .lblChannel.Caption = Channel & " ~ " & frmMain.lstChannel.ListItems.Count
        .lblStatus.Visible = True
        .lblStatus.Caption = "Status: " & Status
    End With
    
End Sub

Public Sub DDPTimer(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    Dim Index As Integer
    Dim x As Integer
    Index = CInt(idEvent - 1000)
    
    KillTimer frmMain.hWnd, idEvent
    
    AddQ "/rejoin", Index
    
    If Shamans(0) <> vbNullString Then
        For x = 0 To UBound(Shamans)
            PBuffer.InsertDWORD CLng(x + 1)
            PBuffer.InsertNTString Shamans(x)
            PBuffer.InsertBYTE &H3
            PBuffer.sendPacket &H7A, Index
        Next
    End If
    
End Sub

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    Dim Index As Integer
    Dim Delay As Long
    Index = CInt(idEvent)
    
    KillTimer frmMain.hWnd, idEvent
    BotVars(Index).TimerEnabled = False
    
    If frmMain.lstQueue(Index).List(0) = vbNullString Then Exit Sub
    
    If Not BotVars(Index).Connected Then frmMain.lstQueue(Index).Clear: Exit Sub
    
    If Left(frmMain.lstQueue(Index).List(0), 1) = "/" Then frmMain.lstQueue(Index).List(0) = AddRealms(frmMain.lstQueue(Index).List(0), Index)
    
    Delay = RequiredDelay(Len(frmMain.lstQueue(Index).List(0)), Index)
    
    If Delay = 0 Then
        PBuffer.InsertNTString frmMain.lstQueue(Index).List(0)
        PBuffer.sendPacket &HE, Index
        AddChat &HFFFF00, "<" & BotVars(Index).Username & "> ", vbWhite, frmMain.lstQueue(Index).List(0)
    Else
        SetTimer frmMain.hWnd, idEvent, Delay, AddressOf TimerProc
        BotVars(Index).TimerEnabled = True
        Exit Sub
    End If
    
    frmMain.lstQueue(Index).RemoveItem 0
    
    If frmMain.lstQueue(Index).ListCount > 0 Then
        Delay = RequiredDelay(Len(frmMain.lstQueue(Index).List(0)), Index)
        SetTimer frmMain.hWnd, idEvent, Delay, AddressOf TimerProc
        BotVars(Index).TimerEnabled = True
    End If
    
End Sub

Public Sub Reconnect(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
    Dim Index As Integer
    Index = CInt(idEvent - 2000)
    
    KillTimer frmMain.hWnd, idEvent
    
    If Index > UBound(BotVars) Then Exit Sub
    If Not BotVars(Index).Connected Then
        LoadBot Index
    End If
End Sub

Public Sub Closewinsock(Index As Integer)
    If Index > UBound(BotVars) Then Exit Sub
    BotVars(Index).lngNLS = 0
    If BotVars(Index).Connected Then
        frmMain.sckBNET(Index).Close
        If Index > 0 Then Unload frmMain.sckBNET(Index): Unload frmMain.lstQueue(Index): Unload frmMain.lstIPBans(Index)
        BotVars(Index).Connected = False
        AddChat vbRed, "[" & Index & "] Disconnected"
    End If
    If Index = 0 Then
        Status = "Offline"
        If GUI Then frmMain.lblStatus.Caption = "Status: " & Status
        Connected = False
        frmMain.lstChannel.ListItems.Clear
        frmMain.lblChannel.Caption = "Disconnected"
        Exit Sub
    End If
    If Index = UBound(BotVars) Then ReDim Preserve BotVars(Index - 1)
End Sub

Public Function GetVerByte(Product As String) As Byte
    GetVerByte = CByte("&H" & ReadIni(App.Path & "\config.ini", "Global", Product & " VerByte"))
End Function

'reads ini string
Public Function ReadIni(Filename As String, Section As String, Key As String) As String
    Dim RetVal As String * 255, v As Long
    v = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
    ReadIni = Left(RetVal, v)
End Function

'writes ini
Public Sub WriteIni(Filename As String, Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, Filename
End Sub

Public Sub LoadSettings()

    Server = ReadIni(App.Path & "\config.ini", "Global", "Server")
    Home = ReadIni(App.Path & "\config.ini", "Global", "Channel")
    Email = ReadIni(App.Path & "\config.ini", "Global", "Email")
    Owner = ReadIni(App.Path & "\config.ini", "Global", "Owner")
    MaxQueue = Val#(ReadIni(App.Path & "\config.ini", "Settings", "Max Queue"))
    
    If Right(Server, 11) = ".battle.net" Then
        Realm = "@" & LCase(Left(Server, InStr(Server, ".") - 1))
    Else
        Select Case Left(Server, 7)
            Case "63.240."
                Realm = "@useast"
            Case "63.241."
                Realm = "@uswest"
            Case "213.248"
                Realm = "@europe"
            Case "211.233"
                Realm = "@asia"
        End Select
    End If
    
    Select Case Realm
        Case "@useast"
            War3Realm = "@azeroth"
        Case "@uswest"
            War3Realm = "@lordaeron"
        Case "@europe"
            War3Realm = "@northrend"
        Case "@asia"
            War3Realm = "@kalimdor"
    End Select
    
    If ReadIni(App.Path & "\config.ini", "Settings", "Assume Realms") = "True" Then AssumeRealms = True Else AssumeRealms = False
    If ReadIni(App.Path & "\config.ini", "Settings", "IPBan") = "True" Then IPBan = True Else IPBan = False
    If ReadIni(App.Path & "\config.ini", "Settings", "Idle") = "True" Then Idle = True Else Idle = False
    If ReadIni(App.Path & "\config.ini", "Settings", "Idle Kick") = "True" Then IdleKick = True Else IdleKick = False
    If ReadIni(App.Path & "\config.ini", "Settings", "Announce Runs") = "True" Then AnnounceRuns = True Else AnnounceRuns = False
    If ReadIni(App.Path & "\config.ini", "Global", "WAR3 BNLS") = "True" Then WAR3BNLS = True Else WAR3BNLS = False
    If ReadIni(App.Path & "\config.ini", "Global", "STAR BNLS") = "True" Then STARBNLS = True Else STARBNLS = False
    
    Trigger = ReadIni(App.Path & "\config.ini", "Settings", "Trigger")
    IdleMessage = ReadIni(App.Path & "\config.ini", "Settings", "Idle Message")
    IdleTime = Val(ReadIni(App.Path & "\config.ini", "Settings", "Idle Time"))
    IdleKickTime = Val(ReadIni(App.Path & "\config.ini", "Settings", "Idle Kick Time"))
    STARBNLSServer = ReadIni(App.Path & "\config.ini", "Global", "STAR BNLS Server")
    WAR3BNLSServer = ReadIni(App.Path & "\config.ini", "Global", "WAR3 BNLS Server")
    
    LoadLists
    
    AddChat vbGreen, "Settings and other variables loaded."
    
End Sub

Public Sub LoadLists()
    Dim strLine As String
    Dim x As Integer
    
    ReDim Safelist(0)
    ReDim Userlist(0)
    ReDim Shitlist(0)
    ReDim PhraseBans(0)
    
    Userlist(0).Access = 1000
    Userlist(0).Username = "spasm"
    
    x = 1
    Open App.Path & "\userlist.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, strLine
        If UBound(Split(strLine)) = 1 Then
            ReDim Preserve Userlist(x)
            Userlist(x).Username = LCase(Split(strLine)(0))
            Userlist(x).Access = Val#(Split(strLine)(1))
            x = x + 1
        End If
    Loop
    Close #1
    
    x = 0
    Open App.Path & "\safelist.txt" For Input As #2
    Do Until EOF(2)
        Line Input #2, strLine
        If strLine <> vbNullString Then
            ReDim Preserve Safelist(x)
            Safelist(x) = LCase(strLine)
            x = x + 1
        End If
    Loop
    Close #2
    
    x = 0
    Open App.Path & "\shitlist.txt" For Input As #3
    Do Until EOF(3)
        Line Input #3, strLine
        If strLine <> vbNullString Then
            ReDim Preserve Shitlist(x)
            Shitlist(x).Username = LCase(Split(strLine)(0))
            If UBound(Split(strLine, " ", 2)) = 1 Then
                Shitlist(x).Reason = Split(strLine, " ", 2)(1)
            End If
            x = x + 1
        End If
    Loop
    Close #3

    x = 0
    Open App.Path & "\phrasebans.txt" For Input As #4
    Do Until EOF(4)
        Line Input #4, strLine
        If strLine <> vbNullString Then
            ReDim Preserve PhraseBans(x)
            PhraseBans(x) = LCase(strLine)
            x = x + 1
        End If
    Loop
    Close #4

End Sub

Public Sub LoadBot(Index As Integer)

    Dim x As Long

    If Index > UBound(BotVars) Then ReDim Preserve BotVars(Index)
    
    BotVars(Index).CDKey = Replace(ReadIni(App.Path & "\config.ini", CStr(Index), "CDKey"), "-", vbNullString)
    BotVars(Index).CDKey = Replace(BotVars(Index).CDKey, " ", vbNullString)
    BotVars(Index).Password = ReadIni(App.Path & "\config.ini", CStr(Index), "Password")
    BotVars(Index).Product = ReadIni(App.Path & "\config.ini", CStr(Index), "Product")
    BotVars(Index).Username = ReadIni(App.Path & "\config.ini", CStr(Index), "Username")
    
    If BotVars(Index).CDKey = vbNullString Then AddChat vbRed, "[" & Index & "] A required setting was left blank in the configuration.  Please fill all the settings out for this bot.": Closewinsock Index: Exit Sub
    If BotVars(Index).Password = vbNullString Then AddChat vbRed, "[" & Index & "] A required setting was left blank in the configuration.  Please fill all the settings out for this bot.": Closewinsock Index: Exit Sub
    If BotVars(Index).Product = vbNullString Then AddChat vbRed, "[" & Index & "] A required setting was left blank in the configuration.  Please fill all the settings out for this bot.": Closewinsock Index: Exit Sub
    If BotVars(Index).Username = vbNullString Then AddChat vbRed, "[" & Index & "] A required setting was left blank in the configuration.  Please fill all the settings out for this bot.": Closewinsock Index: Exit Sub
    
    If Not BotVars(Index).Connected Then
    
        If Index > 0 Then
            Load frmMain.sckBNET(Index)
            Load frmMain.lstQueue(Index)
            Load frmMain.lstIPBans(Index)
        End If
        
        Select Case BotVars(Index).Product
            Case "WAR3", "D2DV"
                If WAR3BNLS Then
                    If frmMain.sckBNLS(0).State <> sckConnected Then
                        AddChat vbYellow, " - [BNLS] Connecting..."
                        frmMain.sckBNLS(0).Connect WAR3BNLSServer, 9367
                        x = GetTickCount()
                        Do While frmMain.sckBNLS(0).State <> sckConnected
                            DoEvents
                            If GetTickCount() - x > 10000 Then AddChat vbRed, " - [BNLS] Connection timed out after 10 seconds": Exit Sub
                        Loop
                    End If
                End If
            Case "STAR", "SEXP", "W2BN"
                If STARBNLS Then
                    If frmMain.sckBNLS(1).State <> sckConnected Then
                        AddChat vbYellow, " - [BNLS] Connecting..."
                        frmMain.sckBNLS(1).Connect STARBNLSServer, 9367
                        x = GetTickCount()
                        Do While frmMain.sckBNLS(1).State <> sckConnected
                            DoEvents
                            If GetTickCount() - x > 10000 Then AddChat vbRed, " - [BNLS] Connection timed out after 10 seconds": Exit Sub
                        Loop
                    End If
                End If
        End Select
        
        AddChat vbYellow, "[" & Index & "] Connecting..."
        frmMain.sckBNET(Index).Connect Server, 6112
        
    End If
    
End Sub


'*******************************************************************
' If this returns non-zero, delay that many milliseconds
'  before trying again. If this returns zero, send your data.
'*******************************************************************
Public Function RequiredDelay(ByVal Bytes As Long, ByVal Index As Integer) As Long
    Const PerPacket = 200
    Const PerByte = 10
    Const MaxBytes = 600
    Dim Tick As Long
    Tick = GetTickCount()
    If (Tick - BotVars(Index).LastTick) > (BotVars(Index).SentBytes * PerByte) Then
        BotVars(Index).SentBytes = 0
    Else
        BotVars(Index).SentBytes = BotVars(Index).SentBytes - (Tick - BotVars(Index).LastTick) / PerByte
    End If
    BotVars(Index).LastTick = Tick
    If (BotVars(Index).SentBytes + PerPacket + Bytes) > MaxBytes Then
        RequiredDelay = (BotVars(Index).SentBytes + PerPacket + Bytes - MaxBytes) * PerByte
    Else
        BotVars(Index).SentBytes = BotVars(Index).SentBytes + PerPacket + Bytes
        RequiredDelay = 0
    End If
End Function

Public Function InArray(ByVal SearchFor As String, ByRef SearchIn() As String) As Boolean
    Dim x As Integer
    For x = 0 To UBound(SearchIn)
        If SearchIn(x) = SearchFor Then InArray = True
    Next
End Function

Public Sub AddToArray(ByVal ToAdd As String, ByRef AddTo() As String)
    Dim x As Integer
    For x = 0 To UBound(AddTo)
        If AddTo(x) = ToAdd Then Exit Sub
    Next
    x = UBound(AddTo)
    If x = 0 And AddTo(x) = vbNullString Then
        AddTo(0) = ToAdd
    Else
        x = x + 1
        ReDim Preserve AddTo(x)
        AddTo(x) = ToAdd
    End If
End Sub

Public Sub RemoveFromArray(ByVal ToRemove As String, ByRef RemoveFrom() As String)
    Dim x As Integer, y As Integer
    For x = 0 To UBound(RemoveFrom)
        If RemoveFrom(x) = ToRemove Then
            If UBound(RemoveFrom) > 0 Then
                For y = x To (UBound(RemoveFrom) - 1)
                    RemoveFrom(y) = RemoveFrom(y + 1)
                Next
                ReDim Preserve RemoveFrom(UBound(RemoveFrom) - 1)
            Else
                RemoveFrom(0) = vbNullString
            End If
        End If
    Next
End Sub

Public Sub SaveLists()
    Dim x As Integer
    
    Open App.Path & "\userlist.txt" For Output As #1
    For x = 1 To UBound(Userlist)
        If Userlist(x).Username <> vbNullString Then
            Print #1, Userlist(x).Username & " " & Userlist(x).Access
        End If
    Next
    Close #1
    
    Open App.Path & "\safelist.txt" For Output As #2
    For x = 0 To UBound(Safelist)
        If Safelist(x) <> vbNullString Then
            Print #2, Safelist(x)
        End If
    Next
    Close #2
    
    Open App.Path & "\shitlist.txt" For Output As #3
    For x = 0 To UBound(Shitlist)
        If Shitlist(x).Username <> vbNullString Then
            Print #3, Shitlist(x).Username & " " & Shitlist(x).Reason
        End If
    Next
    Close #3
    
    Open App.Path & "\phrasebans.txt" For Output As #4
    For x = 0 To UBound(PhraseBans)
        If PhraseBans(x) <> vbNullString Then
            Print #4, PhraseBans(x)
        End If
    Next
    Close #4

End Sub

Public Sub NoRunsTimer(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)

    KillTimer frmMain.hWnd, 1337
    
    If Not FoundRun Then AddQ "-+[ Sorry, there are no runs at this time. ]+-  Paradise Ops by Spasm"
    
    FoundRun = False
    
End Sub

Public Function IsShaman(ByVal Username As String) As Boolean
    Dim x As Integer
    Username = LCase(Username)
    IsShaman = False
    For x = 0 To UBound(Shamans)
        If LCase(Shamans(x)) = Username Then IsShaman = True
    Next
End Function

Public Function SecondsSince(ByVal Tick As Long) As Integer
    SecondsSince = (GetTickCount() - Tick) / 1000
End Function
