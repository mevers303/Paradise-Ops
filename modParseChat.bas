Attribute VB_Name = "modParseChat"
Sub UserInChannel(ByVal Flags As Long, ByVal Ping As Long, ByVal Username As String, ByVal Index As Integer)

    Dim strTemp As String

    strTemp = Username
    Username = RealmUsername(Username, 0)

    If Flags And 32 Then
        AddQ "/ban " & Username & " IP Match: " & BotVars(Index).Username
        frmMain.lstIPBans(Index).AddItem Username
    End If
    
    If Index > 0 Then GoTo Finish
    
    Dim x As Integer
    
    For x = 1 To frmMain.lstChannel.ListItems.Count
        If frmMain.lstChannel.ListItems(x) = strTemp Then Exit Sub
    Next
    
    If Flags And 2 Then
        frmMain.lstChannel.ListItems.Add 1, Username, strTemp
    Else
        frmMain.lstChannel.ListItems.Add , Username, strTemp
    End If
    
    x = frmMain.lstChannel.ListItems(Username).Index
    frmMain.lstChannel.ListItems(x).SubItems(1) = Ping
    frmMain.lstChannel.ListItems(x).SubItems(2) = Flags
    frmMain.lstChannel.ListItems(x).Tag = GetTickCount 'idle
    
    If GUI Then frmMain.lblChannel.Caption = Channel & " ~ " & frmMain.lstChannel.ListItems.Count
    
    x = IsShitlisted(Username)
    If x <> -1 Then AddQ "/ban " & Username & " " & Shitlist(x).Reason
    
Finish:
End Sub

Sub UserJoin(ByVal Flags As Long, ByVal Ping As Long, ByVal Username As String, ByVal Index As Integer)

    Dim strTemp As String

    strTemp = Username
    Username = RealmUsername(Username, 0)

    If Flags And 32 Then
        AddQ "/ban " & Username & " IP Match: " & BotVars(Index).Username
        frmMain.lstIPBans(Index).AddItem Username
    End If
    
    If Index > 0 Then GoTo Finish

    Dim x As Integer
    
    For x = 0 To frmMain.lstRunners.ListCount
        If Username = frmMain.lstRunners.List(x) Then
            WriteIni App.Path & "\runners.ini", "Time", Username, (Val(ReadIni(App.Path & "\runners.ini", "Time", Username)) + SecondsSince(frmMain.lstRunners.ItemData(x))) / 2
            WriteIni App.Path & "\runners.ini", "Runs", Username, (Val(ReadIni(App.Path & "\runners.ini", "Runs", Username))) + 1
            AddQ "/w " & Username & " Your run lasted " & SecondsSince(frmMain.lstRunners.ItemData(x)) & " seconds."
            frmMain.lstRunners.RemoveItem x
        End If
    Next
    
    frmMain.lstChannel.ListItems.Add , Username, strTemp
    x = frmMain.lstChannel.ListItems(Username).Index
    
    frmMain.lstChannel.ListItems(x).SubItems(1) = Ping
    frmMain.lstChannel.ListItems(x).SubItems(2) = Flags
    frmMain.lstChannel.ListItems(x).Tag = GetTickCount 'when they joined
        
    If GUI Then frmMain.lblChannel.Caption = Channel & " ~ " & frmMain.lstChannel.ListItems.Count
    
    x = IsShitlisted(Username)
    If x <> -1 Then AddQ "/ban " & Username & " " & Shitlist(x).Reason
    
Finish:
End Sub

Sub UserLeave(Username As String, ByVal Index As Integer)

    If Index > 0 Then Exit Sub

    frmMain.lstChannel.ListItems.Remove RealmUsername(Username, 0)
    
    If GUI Then frmMain.lblChannel.Caption = Channel & " ~ " & frmMain.lstChannel.ListItems.Count
    
End Sub

Sub UserTalk(ByVal Username As String, ByVal Text As String, ByVal Flags As Long, ByVal Index As Integer)

    If Index > 0 Then Exit Sub

    Dim x As Long
    
    If Flags And 2 Or Flags And 1 Then
        x = vbWhite
    Else
        x = vbYellow
    End If
    
    AddChat x, "<" & Username & "> ", vbWhite, Text
    
    Username = RealmUsername(Username, 0)
    
    If Not IsSafelisted(Username) Then
        If PhraseBans(0) <> vbNullString Then
            For x = 0 To UBound(PhraseBans)
                If PhraseBans(x) <> vbNullString Then
                    If InStr(LCase(Text), " " & PhraseBans(x) & " ") > 0 Then AddQ "/ban " & Username & " Phraseban - " & PhraseBans(x): Exit Sub
                End If
            Next
        End If
    End If
    
    frmMain.lstChannel.ListItems(Username).Tag = GetTickCount()
    
    If Text = Trigger Then Exit Sub
    
    If Left(Text, 1) = Trigger Or LCase(Text) = "?trigger" Then
        ParseCommand Mid(Text, 2), Username, Text
        If Text <> vbNullString And Text <> "Spasm is cool." Then AddQ Text
    End If
    
End Sub

Sub JoinedChannel(ByVal Text As String, ByVal Index As Integer)

    If Index > 0 Then Exit Sub

    Channel = Text
    
    AddChat vbGreen, "Joined Channel: " & Channel
    
    frmMain.lstChannel.ListItems.Clear
    
    HasOps = False
    
    Dim x As Integer
    For x = 0 To UBound(BotVars)
        BotVars(x).HasOps = False
    Next
    
    frmMain.lstBans.Clear
    
    If GUI Then frmMain.lblChannel.Caption = Channel & " ~ " & frmMain.lstChannel.ListItems.Count: frmMain.lblBancount.Caption = "Bancount: " & frmMain.lstBans.ListCount
    
End Sub

Sub UserEmote(ByVal Username As String, ByVal Text As String, ByVal Index As Integer)

    If Index > 0 Then Exit Sub
    
    Dim x As Integer

    AddChat vbYellow, "<" & Username & " " & Text & ">"
    
    Username = RealmUsername(Username, 0)
    
    If Not IsSafelisted(Username) Then
        For x = 0 To UBound(PhraseBans)
            If InStr(LCase(Text), " " & PhraseBans(x) & " ") Then AddQ "/ban " & Username & " Phraseban - " & PhraseBans(x): Exit Sub
        Next
    End If

    frmMain.lstChannel.ListItems(Username).Tag = GetTickCount()
    
End Sub

Sub UserWhisper(ByVal Username As String, ByVal Text As String, ByVal Index As Integer)

    Dim x As Integer

    AddChat vbYellow, "[" & Index & "] <From: " & Username & "> ", &H808080, Text
    
    If AnnounceRuns Then
        If InStr(Text, " entered a Diablo II Lord of Destruction game called ") > 0 Then
            Text = Split(Text, "called ", 2)(1)
            If InStr(LCase(Text), "baal") > 0 Or InStr(LCase(Text), "chaos") > 0 Then
                AddQ "-+[ *" & Username & " has just joined " & Text & " ]+- Paradise Ops by Spasm"
                Username = RealmUsername(Username, Index)
                WriteIni App.Path & "\runners.ini", "Count", Username, Val(ReadIni(App.Path & "\runners.ini", "Count", Username)) + 1
                frmMain.lstRunners.AddItem Username, 0
                frmMain.lstRunners.ItemData(0) = GetTickCount()
                For x = 1 To frmMain.lstRunners.ListCount
                    If Username = frmMain.lstRunners.List(x) Then
                        WriteIni App.Path & "\runners.ini", "Time", Username, (Val(ReadIni(App.Path & "\runners.ini", "Time", Username)) + SecondsSince(frmMain.lstRunners.ItemData(x))) / 2
                        AddQ "/w " & Username & " Your run lasted " & SecondsSince(frmMain.lstRunners.ItemData(x)) & " seconds."
                        frmMain.lstRunners.RemoveItem x
                    End If
                Next
            End If
        End If
    End If
    
    If InStr(Text, " has exited Battle.net.") > 0 Then
        For x = 0 To frmMain.lstRunners.ListCount
            If Username = frmMain.lstRunners.List(x) Then
                WriteIni App.Path & "\runners.ini", "Time", Username, (Val(ReadIni(App.Path & "\runners.ini", "Time", Username)) + SecondsSince(frmMain.lstRunners.ItemData(x))) / 2
                WriteIni App.Path & "\runners.ini", "Runs", Username, (Val(ReadIni(App.Path & "\runners.ini", "Runs", Username))) + 1
                frmMain.lstRunners.RemoveItem x
            End If
        Next
    End If
    
    Username = RealmUsername(Username, Index)
    
    If Text = Trigger Then Exit Sub
        
    If Left(Text, 1) = Trigger Or LCase(Text) = "?trigger" Then
        ParseCommand Mid(Text, 2), Username, Text
        If Text <> vbNullString And Text <> "Spasm is cool." Then AddQ "/w " & Username & " " & Text
    End If
        
End Sub

Sub ServerBroadcast(ByVal Text As String, ByVal Index As Integer)

    If Index > 0 Then Exit Sub

    AddChat &HFFFF00, "<Battle.net Broadcast> " & Text
    
End Sub

Sub FlagsUpdate(ByVal Flags As Long, ByVal Ping As Long, ByVal Username As String, ByVal Index As Integer)

    Dim strTemp As String

    strTemp = Username
    Username = RealmUsername(Username, Index)
    
    frmMain.lstChannel.ListItems(Username).SubItems(2) = Flags
    
    If Flags And 32 Then
        AddQ "/ban " & Username & " IP Match: " & BotVars(Index).Username
        frmMain.lstIPBans(Index).AddItem Username
    End If
    
    If Flags And 2 Then
        Dim Tick As Long
        Tick = frmMain.lstChannel.ListItems(Username).Tag
        frmMain.lstChannel.ListItems.Remove Username
        frmMain.lstChannel.ListItems.Add 1, Username, strTemp
        frmMain.lstChannel.ListItems(1).SubItems(1) = Ping
        frmMain.lstChannel.ListItems(1).SubItems(2) = Flags
        frmMain.lstChannel.ListItems(1).Tag = Tick
        If Right(LCase(strTemp), Len(Realm)) = Realm Then strTemp = Left(strTemp, Len(strTemp) - Len(Realm))
        If Right(LCase(strTemp), Len(War3Realm)) = War3Realm Then strTemp = Left(strTemp, Len(strTemp) - Len(War3Realm))
        For x = 0 To UBound(BotVars)
            If BotVars(x).Username = strTemp Then
                BotVars(x).HasOps = True
                HasOps = True
                Exit For
            End If
        Next
    End If
    
    AddChat &HFFFF00, "<" & strTemp & "> Flags update: " & Flags
    
Finish:
End Sub

Sub WhisperSent(ByVal Username As String, ByVal Text As String, ByVal Index As Integer)

    AddChat &HFFFF00, "[" & Index & "] <To: " & Username & "> ", &H808080, Text
    
End Sub

Sub ChannelFull(ByVal Text As String, ByVal Index As Integer)

    AddChat vbRed, "[" & Index & "] " & Text
    
End Sub

Sub ChannelDoesNotExist(ByVal Text As String, ByVal Index As Integer)

    AddChat vbRed, "[" & Index & "] " & Text
    
End Sub

Sub ChannelRestricted(ByVal Text As String, ByVal Index As Integer)

    AddChat vbRed, "[" & Index & "] " & Text
    
End Sub

Sub Info(ByVal Text As String, ByVal Index As Integer)

    Dim x As Integer
    Dim strTemp As String
    
    strTemp = vbNullString
        
    If InStr(Text, " was banned by ") > 0 Then
        If Index = 0 Then
            strTemp = RealmUsername(Split(Text)(0), 0)
            For x = 0 To frmMain.lstBans.ListCount
                If frmMain.lstBans.List(x) = strTemp Then
                    strTemp = vbNullString
                    GoTo Display
                End If
            Next
            frmMain.lstBans.AddItem strTemp
            If GUI Then frmMain.lblBancount.Caption = "Bancount: " & frmMain.lstBans.ListCount
            strTemp = vbNullString
            GoTo Display
        End If
        Exit Sub
    End If
        
    If InStr(Text, " was unbanned by ") > 0 Then
        If Index = 0 Then
            strTemp = RealmUsername(Split(Text)(0), 0)
            For x = 0 To frmMain.lstBans.ListCount
                If frmMain.lstBans.List(x) = strTemp Then
                    frmMain.lstBans.RemoveItem x
                End If
            Next
            If GUI Then frmMain.lblBancount.Caption = "Bancount: " & frmMain.lstBans.ListCount
            strTemp = vbNullString
            GoTo Display
        End If
        Exit Sub
    End If
            
    If InStr(Text, " has been squelched.") > 0 Then
        frmMain.lstIPBans(Index).AddItem RealmUsername(Split(Text)(0), Index)
        strTemp = "[" & Index & "] "
        GoTo Display
    End If
    
    If InStr(Text, " has been unsquelched.") > 0 Then
        strTemp = RealmUsername(Split(Text)(0), Index)
        For x = 0 To frmMain.lstIPBans(Index).ListCount
            If frmMain.lstIPBans(Index).List(x) = strTemp Then
                frmMain.lstIPBans(Index).RemoveItem x
            End If
        Next
        strTemp = "[" & Index & "] "
        GoTo Display
    End If
    
    If InStr(Text, " to your friends list.") > 0 Then
        AddQ Split(Text)(1) & " has successfully been added to " & BotVars(Index).Username & "'s friends list.  Make sure they add that bot to their friends list as well!"
        strTemp = "[" & Index & "] "
        GoTo Display
    End If
    
    If InStr(Text, " from your friends list.") > 0 Then
        AddQ Split(Text)(1) & " has been removed from " & BotVars(Index).Username & "'s friends list."
        strTemp = "[" & Index & "] "
        GoTo Display
    End If
    
Display:
    AddChat 13408512, strTemp & Text
    
End Sub

Sub BNETError(ByVal Text As String, ByVal Index As Integer)

    AddChat vbRed, "[" & Index & "] " & Text
    
End Sub
