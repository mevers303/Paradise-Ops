Attribute VB_Name = "ParseCommands"
Option Explicit

Private LastRunsCommand As Long

Public Sub ParseCommand(ByVal Command As String, ByVal Username As String, ByRef strReturn As String, Optional ByVal Index As Integer = 0)
    Dim Access As Integer
    Access = GetAccess(Username)
    Dim x As Integer
    Dim y As Integer
    Dim strTemp As String
    Dim strTemp2 As String
    Dim strArray() As String
    strArray = Split(Command)
    strReturn = vbNullString
    Select Case strArray(0)
        Case "ban", "b"
            If Access < 60 Then Exit Sub
            strArray = Split(Command, " ", 3)
            If UBound(strArray) < 1 Then Exit Sub
            strTemp = strArray(1)
            ReDim Preserve strArray(2)
            If InStr(strTemp, "*") = 0 Then
                If Not BanAccess(strTemp, Access, True) Then Exit Sub
                AddQ "/ban " & strTemp & " " & strArray(2), 1
                If IPBan Then AddQ "/squelch " & strTemp
            Else
                strTemp2 = strArray(2)
                strArray = WildcardChannel(strTemp)
                If strArray(0) = vbNullString Then Exit Sub
                For x = 0 To UBound(strArray)
                    If BanAccess(strArray(x), Access) Then
                        AddQ "/ban " & strArray(x) & " " & strArray(2)
                        If IPBan Then AddQ "/squelch " & strTemp
                    End If
                Next
            End If
            strReturn = "Spasm is cool."
        Case "kick", "k"
            If Access < 50 Then Exit Sub
            strArray = Split(Command, " ", 3)
            If UBound(strArray) < 1 Then Exit Sub
            strTemp = strArray(1)
            ReDim Preserve strArray(2)
            If InStr(strTemp, "*") = 0 Then
                If Not BanAccess(strTemp, Access, True) Then Exit Sub
                AddQ "/kick " & strTemp & " " & strArray(2), 1
            Else
                strTemp2 = strArray(2)
                strArray = WildcardChannel(strTemp)
                If strArray(0) = vbNullString Then Exit Sub
                For x = 0 To UBound(strArray)
                    If BanAccess(strArray(x), Access) Then
                        AddQ "/kick " & strArray(x) & " " & strArray(2)
                    End If
                Next
            End If
            strReturn = "Spasm is cool."
        Case "unban", "unip", "ub", "unipban"
            If Access < 50 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            strTemp = LCase(strArray(1))
            If InStr(Command, "*") = 0 Then
                AddQ "/unban " & strTemp, 1
                RemoveIPBan strTemp
            Else
                For x = 0 To (frmMain.lstBans.ListCount - 1)
                    If frmMain.lstBans.List(x) Like strTemp Then
                        AddQ "/unban " & frmMain.lstBans.List(x), 1
                        RemoveIPBan frmMain.lstBans.List(x)
                    End If
                Next
            End If
            strReturn = "Spasm is cool."
        Case "ip", "ipban"
            If Access < 60 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            strTemp = strArray(1)
            If InStr(strTemp, "*") = 0 Then
                If BanAccess(strTemp, Access, True) Then AddQ "/squelch " & strTemp
            Else
                strArray = WildcardChannel(strTemp)
                If strArray(0) = vbNullString Then Exit Sub
                For x = 0 To UBound(strArray)
                    If BanAccess(strArray(x), Access) Then AddQ "/squelch " & strArray(x)
                Next
            End If
            strReturn = "Spasm is cool."
        Case "trigger"
            If Access = 0 Then Exit Sub
            If UBound(strArray) = 0 Then
                strReturn = "The bot's current trigger is """ & Trigger & """"
            End If
        Case "whoami", "a", "access", "whois"
            If Access < 20 Then Exit Sub
            If UBound(strArray) = 1 Then
                strReturn = strArray(1) & " has " & GetAccess(LCase(strArray(1))) & " access."
            Else
                strReturn = "You have " & Access & " access."
            End If
        Case "ver", "version", "v"
            If Access = 0 Then Exit Sub
            strReturn = "Paradise Ops preview by Spasm"
        Case "say"
            If Access < 30 Then Exit Sub
            strTemp = Split(Command, " ", 2)(1)
            If Access < 100 Then
                strTemp = Username & " - " & strTemp
            End If
            AddQ strTemp
            strReturn = "Spasm is cool."
        Case "cq"
            If Access < 20 Then Exit Sub
            For x = 0 To UBound(BotVars)
                If BotVars(x).Connected Then frmMain.lstQueue(x).Clear
            Next
            strReturn = "Queue has been cleared."
        Case "scq"
            If Access < 20 Then Exit Sub
            For x = 0 To UBound(BotVars)
                If BotVars(x).Connected Then frmMain.lstQueue(x).Clear
            Next
        Case "add", "set"
            If Access < 70 Then Exit Sub
            If UBound(strArray) <> 2 Then Exit Sub
            If Val(strArray(2)) > Access Then strReturn = "You cannot add a user with more access than yourself.": Exit Sub
            strTemp = LCase(strArray(1))
            If Right(strTemp, Len(Realm)) = Realm And AssumeRealms Then strTemp = Left(strTemp, Len(strTemp) - Len(Realm))
            For x = 1 To UBound(Userlist)
                If Userlist(x).Username = strTemp Then
                    If Userlist(x).Access >= Access Then strReturn = "You cannot modify a user with equal or greater access than you.": Exit Sub
                    Userlist(x).Access = Val(strArray(2))
                    SaveLists
                    strReturn = strArray(1) & " has successfully been added with " & strArray(2) & " access."
                    Exit Sub
                End If
            Next
            x = UBound(Userlist) + 1
            ReDim Preserve Userlist(x)
            Userlist(x).Username = strTemp
            Userlist(x).Access = Val#(strArray(2))
            SaveLists
            strReturn = strArray(1) & " has successfully been added with " & strArray(2) & " access."
        Case "remove", "delete", "del"
            If Access < 70 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            strTemp = LCase(strArray(1))
            If Right(strTemp, Len(Realm)) = Realm And AssumeRealms Then strTemp = Left(strTemp, Len(strTemp) - Len(Realm))
            For x = 1 To UBound(Userlist)
                If Userlist(x).Username = strTemp Then
                    If Userlist(x).Access > Access Then strReturn = "You cannot remove a user with more access than you.": Exit Sub
                    Userlist(x).Access = 0
                    Userlist(x).Username = vbNullString
                    SaveLists
                    strReturn = "Successfully removed " & strArray(1) & " from the userlist."
                    Exit Sub
                End If
            Next
            strReturn = "Could not find " & strArray(1) & " in the userlist."
        Case "safelist", "safeadd"
            If Access < 70 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            strTemp = LCase(strArray(1))
            If Right(strTemp, Len(Realm)) = Realm And AssumeRealms Then strTemp = Left(strTemp, Len(strTemp) - Len(Realm))
            AddToArray strTemp, Safelist
            strReturn = "Successfully added " & strArray(1) & " to the safelist."
            SaveLists
        Case "saferemove", "safedel", "unsafelist", "safedelete", "saferem"
            If Access < 70 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            strTemp = LCase(strArray(1))
            If Right(strTemp, Len(Realm)) = Realm And AssumeRealms Then strTemp = Left(strTemp, Len(strTemp) - Len(Realm))
            RemoveFromArray strTemp, Safelist
            strReturn = "Removed " & strArray(1) & " from the safelist."
            SaveLists
        Case "phraseban", "pban", "padd", "phraseadd"
            If Access < 70 Then Exit Sub
            If UBound(strArray) > 0 Then
                AddToArray LCase(Split(Command, " ", 2)(1)), PhraseBans
                strReturn = "Phrasebanned added - """ & Split(Command, " ", 2)(1) & """"
                SaveLists
            End If
        Case "premove", "phasedel", "pdel", "prem"
            If Access < 70 Then Exit Sub
            If UBound(strArray) < 1 Then Exit Sub
            strTemp = LCase(Split(Command, " ", 2)(1))
            strReturn = "Could not find """ & strTemp & """ in the phrasebans."
            RemoveFromArray strTemp, PhraseBans
            strReturn = "Removed """ & strTemp & """ from the phrasebans."
            SaveLists
        Case "shitlist", "shitadd", "sa", "bl", "blacklist"
            If Access < 70 Then Exit Sub
            If UBound(strArray) < 1 Then Exit Sub
            strTemp = strArray(1)
            If Not BanAccess(strTemp, Access, True) Then Exit Sub
            strTemp = LCase(strTemp)
            If Right(strTemp, Len(Realm)) = Realm And AssumeRealms Then strTemp = Left(strTemp, Len(strTemp) - Len(Realm))
            If Shitlist(0).Username <> vbNullString Then
                x = UBound(Shitlist) + 1
                ReDim Preserve Shitlist(x)
                Shitlist(x).Username = strTemp
                If UBound(Split(Command, " ", 3)) = 2 Then
                    Shitlist(x).Reason = Split(Command, " ", 3)(2)
                Else
                    Shitlist(x).Reason = "Shitlisted"
                End If
                ParseCommand "/ban " & Shitlist(x).Username & " " & Shitlist(x).Reason, Username, strTemp2
            Else
                Shitlist(0).Username = strTemp
                If UBound(Split(Command, " ", 3)) = 2 Then
                    Shitlist(0).Reason = Split(Command, " ", 3)(2)
                Else
                    Shitlist(0).Reason = "Shitlisted"
                End If
                ParseCommand "/ban " & Shitlist(x).Username & " " & Shitlist(x).Reason, Username, strTemp2
            End If
            strReturn = "Added " & strTemp & " to the shitlist."
            SaveLists
        Case "shitdel", "shitrem", "sd", "sr", "unblacklist"
            If Access < 70 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            strTemp = LCase(strArray(1))
            If Right(strTemp, Len(Realm)) = Realm And AssumeRealms Then strTemp = Left(strTemp, Len(strTemp) - Len(Realm))
            strReturn = "Could not find " & strTemp & " in the shitlist."
            For x = 0 To UBound(Shitlist)
                If strTemp = Shitlist(x).Username Then
                    strReturn = "Removed " & strTemp & " from the shitlist."
                    ParseCommand "unban " & strTemp, Username, strTemp2
                    Shitlist(x).Username = vbNullString
                    Shitlist(x).Reason = vbNullString
                End If
            Next
            SaveLists
        Case "runs", "baal", "chaos"
            If (GetTickCount() - LastRunsCommand) < 45000 Then Exit Sub
            SetTimer frmMain.hWnd, 1337, 1000, AddressOf NoRunsTimer
            For x = 0 To UBound(BotVars)
                If BotVars(x).Connected Then
                    PBuffer.sendPacket &H65, x
                End If
            Next
            LastRunsCommand = GetTickCount()
        Case "addrunner"
            If Access < 50 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            For x = 0 To UBound(BotVars)
                If BotVars(x).FriendsCount < 25 And BotVars(x).Connected Then
                    strTemp = LCase(strArray(1))
                    If BotVars(x).Product = "WAR3" Then
                        If Right(strTemp, Len(Realm)) <> Realm Then strTemp = strTemp & Realm
                    End If
                    strReturn = "Attempting to add " & strTemp & " to " & BotVars(x).Username & "'s friends list..."
                    AddQ "/f a " & strTemp, 0, x
                    Exit Sub
                End If
            Next
            strReturn = "Could not find a bot with room on its friends list."
        Case "delrunner"
            If Access < 50 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            For x = 0 To UBound(BotVars)
                strTemp = LCase(strArray(1))
                If BotVars(x).Product = "WAR3" Then
                    If Right(strTemp, Len(Realm)) <> Realm Then strTemp = strTemp & Realm
                End If
                AddQ "/f r " & strTemp
            Next
        Case "timeban", "tb", "tban"
            If Access < 60 Then Exit Sub
            strArray = Split(Command, " ", 3)
            If UBound(strArray) <> 2 Then Exit Sub
            strTemp = strArray(1)
            If InStr(strTemp, "*") = 0 Then
                If Not BanAccess(strTemp, Access, True) Then Exit Sub
                AddQ "/ban " & strTemp & " Timebanned for " & strArray(2) & " minute(s)", 1
                frmMain.lstTimeBans.AddItem strTemp, 0
                frmMain.lstTimeBans.ItemData(0) = GetTickCount
                If IPBan Then AddQ "/squelch " & strTemp
            Else
                strTemp2 = strArray(2)
                strArray = WildcardChannel(strTemp)
                If strArray(0) = vbNullString Then Exit Sub
                For x = 0 To UBound(strArray)
                    If BanAccess(strArray(x), Access) Then
                        AddQ "/ban " & strArray(x) & " Timebanned for " & strArray(2) & " minute(s)"
                        frmMain.lstTimeBans.AddItem strTemp, 0
                        frmMain.lstTimeBans.ItemData(0) = GetTickCount
                        If IPBan Then AddQ "/squelch " & strTemp
                    End If
                Next
            End If
            strReturn = "Spasm is cool."
        Case "settrigger", "trigger"
            If Access < 80 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            strTemp = Left(strArray(1), 1)
            Trigger = strTemp
            WriteIni App.Path & "\config.ini", "Settings", "Trigger", Trigger
            strReturn = "The bot's new trigger is """ & Trigger & """"
        Case "load", "connect"
            If Access < 90 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            If IsNumeric(strArray(1)) Then
                LoadBot CInt(strArray(1))
            End If
            strReturn = "Attempting to load bot #" & strArray(1) & "."
        Case "unload", "disconnect"
            If Access < 90 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            If IsNumeric(strArray(1)) Then
                Closewinsock CInt(strArray(1))
            End If
            strReturn = "Attempting to unload bot #" & strArray(1) & "."
        Case "reload", "refresh"
            If Access < 90 Then Exit Sub
            LoadSettings
            strReturn = "Settings and other variables loaded."
        Case "op", "giveup", "giveops"
            If Access < 90 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            AddQ "/designate " & strArray(1)
            AddQ "/rejoin"
            strReturn = "Spasm is cool."
        Case "designate", "d"
            If Access < 90 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            AddQ "/designate " & strArray(1)
            strReturn = "Spasm is cool."
        Case "rejoin"
            If Access < 70 Then Exit Sub
            For x = 0 To UBound(BotVars)
                If BotVars(x).Connected Then
                    With PBuffer
                        .InsertDWORD 2
                        .InsertNTString "Paradise Ops"
                        .sendPacket &HC, x
                        .InsertDWORD 2
                        .InsertNTString Channel
                        .sendPacket &HC, x
                    End With
                End If
            Next
            strReturn = "Spasm is cool."
        Case "ddp"
            If Access < 90 Then Exit Sub
            For x = 0 To UBound(BotVars)
                If BotVars(x).ClanRank = 4 And BotVars(x).Connected Then
                    AddQ "/designate " & strArray(1), 1, x
                    If Shamans(0) <> vbNullString Then
                        For y = 0 To UBound(Shamans)
                            PBuffer.InsertDWORD (y + 1)
                            PBuffer.InsertNTString Shamans(y)
                            PBuffer.InsertBYTE &H2
                            PBuffer.sendPacket &H7A, x
                        Next
                    End If
                    SetTimer frmMain.hWnd, x + 1000, 3000, AddressOf DDPTimer
                    strReturn = "Spasm is cool."
                    Exit Sub
                End If
            Next
            strReturn = "One of the bots must be chieftan to use demote designate promote."
        Case "rank", "clanrank", "crank", "setrank"
            If Access < 90 Then Exit Sub
            If UBound(strArray) <> 2 Then Exit Sub
            If Not IsNumeric(strArray(2)) Then Exit Sub
            If Val(strArray(2)) > 3 Or Val(strArray(2)) < 1 Then Exit Sub
            If Right(LCase(strArray(1)), Len(War3Realm)) = War3Realm Then strArray(1) = Left(strArray(1), Len(strArray(1)) - Len(War3Realm))
            For x = 0 To UBound(BotVars)
                If BotVars(x).ClanRank > 2 And BotVars(x).Connected Then
                    If (BotVars(x).ClanRank = 4 And (IsShaman(strArray(1)) Or Val(strArray(2)) = 3)) Or (Val(strArray(2)) < 3) Then
                        PBuffer.InsertDWORD (x + 1)
                        PBuffer.InsertNTString strArray(1)
                        PBuffer.InsertBYTE strArray(2)
                        PBuffer.sendPacket &H7A, x
                        strReturn = "Attempting to change user's rank..."
                        Exit Sub
                    End If
                End If
            Next
            strReturn = "None of the bots are sufficient in rank to do that."
        Case "invite"
            If Access < 60 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            If Right(LCase(strArray(1)), Len(War3Realm)) = War3Realm Then strArray(1) = Left(strArray(1), Len(strArray(1)) - Len(War3Realm))
            For x = 0 To UBound(BotVars)
                If BotVars(x).ClanRank > 2 And BotVars(x).Connected Then
                    PBuffer.InsertDWORD (x + 1)
                    PBuffer.InsertNTString strArray(1)
                    strReturn = "Inviting user to the clan..."
                    Exit Sub
                End If
            Next
            strReturn = "None of the bots are sufficient in rank to do that."
        Case "clanremove", "cremove"
            If Access < 70 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            If Right(LCase(strArray(1)), Len(War3Realm)) = War3Realm Then strArray(1) = Left(strArray(1), Len(strArray(1)) - Len(War3Realm))
            For x = 0 To UBound(BotVars)
                If BotVars(x).ClanRank > 2 And BotVars(x).Connected Then
                    If (BotVars(x).ClanRank = 4 And IsShaman(strArray(1))) Or Not IsShaman(strArray(1)) Then
                        PBuffer.InsertDWORD (x + 1)
                        PBuffer.InsertNTString strArray(1)
                        strReturn = "Removing user from clan..."
                        Exit Sub
                    End If
                End If
            Next
            strReturn = "None of the bots are sufficient in rank to do that."
        Case "motd", "setmotd"
            If UBound(strArray) = 0 Then
                If Access < 50 Then Exit Sub
                For x = 0 To UBound(BotVars)
                    If BotVars(x).Product = "WAR3" And BotVars(x).Connected Then
                        PBuffer.InsertDWORD (x + 1)
                        PBuffer.sendPacket &H7C, x
                    End If
                    strReturn = "Spasm is cool."
                    Exit Sub
                Next
            Else
                If Access < 70 Then Exit Sub
                strArray = Split(Command, " ", 2)
                For x = 0 To UBound(BotVars)
                    If BotVars(x).ClanRank > 2 And BotVars(x).Connected Then
                        PBuffer.InsertDWORD (x + 1)
                        PBuffer.InsertNTString strArray(1)
                        strReturn = "Changed the message of the day."
                        Exit Sub
                    End If
                Next
                strReturn = "None of the bots are sufficient in rank to do that."
            End If
        Case "invites", "accept"
            If Access < 80 Then Exit Sub
            If UBound(strArray) = 0 Then
                If AcceptClanInvites Then
                    AcceptClanInvites = False
                    strReturn = "All bots are now rejecting clan invitations."
                Else
                    AcceptClanInvites = True
                    strReturn = "All bots are now accepting clan invitations."
                End If
            Else
                Select Case strArray(1)
                    Case "on", "1", "enable"
                        AcceptClanInvites = True
                        strReturn = "All bots are now accepting clan invitations."
                    Case "off", "0", "disable"
                        AcceptClanInvites = False
                        strReturn = "All bots are now rejecting clan invitations."
                End Select
            End If
        Case "annnounceruns"
            If Access < 80 Then Exit Sub
            If UBound(strArray) = 0 Then
                If AnnounceRuns Then
                    AnnounceRuns = False
                    WriteIni App.Path & "\config.ini", "Settings", "Announce Runs", "False"
                    strReturn = "No longer announcing runs."
                Else
                    AnnounceRuns = True
                    WriteIni App.Path & "\config.ini", "Settings", "Announce Runs", "True"
                    strReturn = "Runs will now be announced."
                End If
            Else
                Select Case strArray(1)
                    Case "on", "1", "enable"
                        AnnounceRuns = True
                        WriteIni App.Path & "\config.ini", "Settings", "Announce Runs", "True"
                        strReturn = "Runs will now be announced."
                    Case "off", "0", "disable"
                        AnnounceRuns = False
                        WriteIni App.Path & "\config.ini", "Settings", "Announce Runs", "False"
                        strReturn = "No longer announcing runs."
                End Select
            End If
        Case "idlekick"
            If Access < 90 Then Exit Sub
            If UBound(strArray) = 0 Then
                If IdleKick Then
                    IdleKick = False
                    WriteIni App.Path & "\config.ini", "Settings", "Idle Kick", "False"
                    strReturn = "Idle kick is off."
                Else
                    IdleKick = True
                    WriteIni App.Path & "\config.ini", "Settings", "Idle Kick", "True"
                    strReturn = "Kicking users idle for more than " & IdleKickTime & " seconds."
                End If
            Else
                If IsNumeric(strArray(1)) Then
                    If strArray(1) <> "0" And strArray(1) <> "1" Then
                        IdleKickTime = Val(strArray(1))
                        WriteIni App.Path & "\config.ini", "Settings", "Idle Kick Time", CStr(IdleKickTime)
                        strReturn = "Idle kick time set to " & IdleKickTime & " seconds."
                        Exit Sub
                    End If
                End If
                Select Case strArray(1)
                    Case "on", "1", "enable"
                        IdleKick = True
                        WriteIni App.Path & "\config.ini", "Settings", "Idle Kick", "True"
                        strReturn = "Kicking users idle for more than " & IdleKickTime & " seconds."
                    Case "off", "0", "disable"
                        IdleKick = False
                        WriteIni App.Path & "\config.ini", "Settings", "Idle Kick", "False"
                        strReturn = "Idle kick is off."
                End Select
            End If
        Case "idle", "idletime", "idlemessage"
            If Access < 80 Then Exit Sub
            Select Case UBound(strArray)
                Case 0
                    If Idle Then
                        Idle = False
                        WriteIni App.Path & "\config.ini", "Settings", "Idle", "False"
                        strReturn = "Idle messages have been turned off."
                    Else
                        Idle = True
                        WriteIni App.Path & "\config.ini", "Settings", "Idle", "True"
                        strReturn = "Idle messages have been turned on."
                    End If
                Case 1
                    If IsNumeric(strArray(1)) Then
                        If strArray(1) <> "0" And strArray(1) <> "1" Then
                            IdleTime = Val(strArray(1))
                            WriteIni App.Path & "\config.ini", "Settings", "Idle Time", CStr(IdleKickTime)
                            strReturn = "Idle time set to " & IdleTime & " seconds."
                            Exit Sub
                        End If
                    End If
                    Select Case strArray(1)
                        Case "on", "1", "enable"
                            Idle = True
                            WriteIni App.Path & "\config.ini", "Settings", "Idle", "True"
                            strReturn = "Idle messages have been turned on."
                        Case "off", "0", "disable"
                            Idle = False
                            WriteIni App.Path & "\config.ini", "Settings", "Idle", "False"
                            strReturn = "Idle messages have been turned off."
                        Case Else
                            IdleMessage = strArray(1)
                            WriteIni App.Path & "\config.ini", "Settings", "Idle Message", IdleMessage
                            strReturn = "Idle message has been set."
                    End Select
                Case Else
                    strArray = Split(Command, " ", 2)
                    IdleMessage = strArray(1)
                    WriteIni App.Path & "\config.ini", "Settings", "Idle Message", IdleMessage
                    strReturn = "Idle message turned on and set."
            End Select
        Case "bancount"
            If Access < 40 Then Exit Sub
            strReturn = "Current ban count: " & frmMain.lstBans.ListCount
        Case "assumerealms"
            If Access < 100 Then Exit Sub
            If UBound(strArray) = 0 Then
                If AssumeRealms Then
                    AssumeRealms = False
                    WriteIni App.Path & "\config.ini", "Settings", "Assume Realms", "False"
                    strReturn = "Bot will no longer assume STAR/SEXP/D2DV/D2XP/W2BN realms."
                Else
                    AssumeRealms = True
                    WriteIni App.Path & "\config.ini", "Settings", "Assume Realms", "True"
                    strReturn = "Bot will assume STAR/SEXP/D2DV/D2XP/W2BN realms."
                End If
            Else
                Select Case strArray(1)
                    Case "1", "on", "enable"
                        AssumeRealms = True
                        WriteIni App.Path & "\config.ini", "Settings", "Assume Realms", "True"
                        strReturn = "Bot will assume STAR/SEXP/D2DV/D2XP/W2BN realms."
                    Case "0", "off", "disable"
                        AssumeRealms = False
                        WriteIni App.Path & "\config.ini", "Settings", "Assume Realms", "False"
                        strReturn = "Bot will no longer assume STAR/SEXP/D2DV/D2XP/W2BN realms."
                End Select
            End If
        Case "ipbanning"
            If Access < 90 Then Exit Sub
            If UBound(strArray) = 0 Then
                If IPBan Then
                    IPBan = False
                    WriteIni App.Path & "\config.ini", "Settings", "IPBan", "False"
                    strReturn = "Bot will no longer IP ban banned users."
                Else
                    IPBan = True
                    WriteIni App.Path & "\config.ini", "Settings", "IPBan", "True"
                    strReturn = "Bot will IP ban banned users."
                End If
            Else
                Select Case strArray(1)
                    Case "1", "on", "enable"
                        IPBan = True
                        WriteIni App.Path & "\config.ini", "Settings", "IPBan", "True"
                        strReturn = "Bot will IP ban banned users."
                    Case "0", "off", "disable"
                        IPBan = False
                        WriteIni App.Path & "\config.ini", "Settings", "IPBan", "False"
                        strReturn = "Bot will no longer IP ban banned users."
                End Select
            End If
        Case "gui"
            If Access < 100 Then Exit Sub
            If UBound(strArray) = 0 Then
                If GUI Then
                    DisableGUI
                    strReturn = "GUI has been disabled."
                Else
                    EnableGUI
                    strReturn = "GUI has been enabled."
                End If
            Else
                Select Case strArray(1)
                    Case "1", "on", "enable"
                        EnableGUI
                        strReturn = "GUI has been enabled."
                    Case "0", "off", "disable"
                        DisableGUI
                        strReturn = "GUI has been disabled."
                End Select
            End If
        Case "server", "setserver"
            If Access < 100 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            Server = strArray(1)
            WriteIni App.Path & "\config.ini", "Global", "Server", Server
            strReturn = "Server has been set to " & Server
        Case "sethome"
            If Access < 100 Then Exit Sub
            If UBound(strArray) = 0 Then Exit Sub
            strArray = Split(Command, " ", 2)
            Home = strArray(1)
            WriteIni App.Path & "\config.ini", "Global", "Channel", Home
            strReturn = "Home channel has been set to " & Home & "."
        Case "email", "setemail"
            If Access < 100 Then Exit Sub
            If UBound(strArray) <> 1 Then Exit Sub
            Email = strArray(1)
            WriteIni App.Path & "\config.ini", "Global", "Email", Email
            strReturn = "Registration email has been set to " & Email & "."
    End Select
End Sub

Sub RemoveIPBan(ByVal Username As String)
    Dim x As Integer, y As Integer
    If AssumeRealms And Right(Username, Len(War3Realm)) <> War3Realm And Right(Username, Len(Realm)) <> Realm Then Username = Username & Realm
    For y = 0 To UBound(BotVars)
        For x = 0 To frmMain.lstIPBans(y).ListCount
            If frmMain.lstIPBans(y).List(x) = Username Then
                AddQ "/unsquelch " & Username, 1, y
            End If
        Next
    Next
End Sub

Function GetAccess(ByVal Username As String) As Integer
    If Username = "§" Then GetAccess = 32767: Exit Function
    Dim x As Integer
    Username = LCase(Username)
    If Right(Username, Len(Realm)) = Realm And AssumeRealms Then Username = Left(Username, Len(Username) - Len(Realm))
    For x = 0 To UBound(Userlist)
        If Userlist(x).Username = Username Then GetAccess = Userlist(x).Access: Exit Function
    Next
    GetAccess = 0
End Function

Public Function IsSafelisted(ByVal Username As String) As Boolean
    Dim x As Integer
    Username = LCase(Username)
    If Right(Username, Len(Realm)) = Realm And AssumeRealms Then Username = Left(Username, Len(Username) - Len(Realm))
    For x = 0 To UBound(Safelist)
        If Safelist(x) = Username Then IsSafelisted = True: Exit Function
    Next
    If GetAccess(Username) > 0 Then IsSafelisted = True: Exit Function
    IsSafelisted = False
End Function

Function WildcardChannel(ByVal Match As String) As String()
    Dim x As Integer
    Dim strArray() As String
    ReDim strArray(0)
    Match = LCase(Match)
    For x = 1 To frmMain.lstChannel.ListItems.Count
        If frmMain.lstChannel.ListItems(x).Key Like Match Then
            ReDim Preserve strArray(x - 1)
            strArray(x - 1) = frmMain.lstChannel.ListItems(x).Key
        End If
    Next
    WildcardChannel = strArray
End Function

Function BanAccess(ByVal Username As String, ByVal Access As Integer, Optional ByVal OverRideSafelist As Boolean = False) As Boolean
    Dim UserAccess As Integer
    UserAccess = GetAccess(Username)
    If UserAccess >= Access Then BanAccess = False: Exit Function
    If IsSafelisted(Username) And Not OverRideSafelist And Access < 80 Then BanAccess = False: Exit Function
    BanAccess = True
End Function

Public Function IsShitlisted(ByVal Username As String) As Integer
    Dim x As Integer
    Username = LCase(Username)
    If Right(Username, Len(Realm)) = Realm And AssumeRealms Then Username = Left(Username, Len(Username) - Len(Realm))
    For x = 0 To UBound(Shitlist)
        If InStr(Username, "*") = 0 Then
            If Shitlist(x).Username = Username Then IsShitlisted = x: Exit Function
        Else
            If Shitlist(x).Username Like Username Then IsShitlisted = x: Exit Function
        End If
    Next
    IsShitlisted = -1
End Function

