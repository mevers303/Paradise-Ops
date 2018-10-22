Attribute VB_Name = "modPackets"
Option Explicit

Public Sub ParseData(ByVal Data As String, ByVal Index As Integer)
    ' This Sub parses inbound packets recieved by Winsock and decides which packet
    ' is which and then directs it to the appropriate Sub to be further broken down
    ' and used by the bot
    Dim PacketID As Byte
    
    If Len(Data) = 2 Then Exit Sub
    
    PacketID = Asc(Mid(Data, 2, 1))
    Data = Mid(Data, 5)
    
    Debuffer.DebuffPacket Data    'Give the debuffer the packet to hold on to

    Select Case PacketID
        Case &H0: PBuffer.sendPacket &H0, Index
        Case &H25: Parse0x25 Index 'SID_PING
        Case &H3A: Parse0x3A Index 'SID_LOGONRESPONSE2
        Case &H3D: Parse0x3D Index 'SID_CREATEACCOUNT
        Case &H50: Parse0x50 Index 'SID_AUTH_INFO
        Case &H51: Parse0x51 Index 'SID_AUTH_CHECK
        Case &H52: Parse0x52 Index 'SID_AUTH_ACCOUNTCREATE
        Case &H53: Parse0x53 Index 'SID_AUTH_ACCOUNTLOGON
        Case &H54: Parse0x54 Index 'SID_AUTH_ACCOUNTLOGONPROOF
        Case &HA: Parse0x0A Index  'SID_ENTERCHAT
        Case &H59: Parse0x59 Index 'SID_SETEMAIL
        Case &HF: Parse0x0F Index  'SID_CHATEVENT
        Case &H19: Parse0x19 Index 'SID_MESSAGEBOX
        Case &H65: Parse0x65 Index 'SID_FRIENDSLIST
        Case &H75: Parse0x75 Index 'SID_CLANINFO
        Case &H72: Parse0x72 Index 'SID_CLANCREATIONINVITATION
        Case &H73: Parse0x73 Index 'SID_CLANDISBAND
        Case &H76: Parse0x76 Index 'SID_CLANQUITNOTIFY
        Case &H77: Parse0x77 Index 'SID_CLANINVITATION
        Case &H78: Parse0x78 Index 'SID_CLANREMOVEMEMBER
        Case &H79: Parse0x79 Index 'SID_CLANINVITATIONRESPONSE
        Case &H7A: Parse0x7A Index 'SID_CLANRANKCHANGE
        Case &H7C: Parse0x7C Index 'SID_CLANMOTD
        Case &H7D: Parse0x7D Index 'SID_CLANMEMBERLIST
        Case &H13: Parse0x13 Index 'SID_FLOODDETECTED
        Case &HB, &H4C, &H66, &H67, &H68:  'Ignore
        Case Else: ParseUnknown PacketID, Data
    End Select
    
    Debuffer.Clear                'Clear the debuffer when we're done
    
End Sub

Public Sub ParseBNLSData(ByVal Data As String, ByVal Socket As Integer)

    Dim PacketID As Byte
    Dim Index As Integer
    Dim Checksum As Long
    Dim lngDecoder As Long
    Dim HashLength As Long
    Dim EXEInfo As String
    Dim KeyHash As String
    Dim EXEVersion As Long
        
    Debuffer.Clear
    Debuffer.DebuffPacket Data
    
    Debuffer.RemoveWORD
    PacketID = Debuffer.DebuffBYTE
    
    Select Case Debuffer.DebuffDWORD
        Case &H1
        
            AddChat vbYellow, " - [BNLS] Version information received"
        
            With Debuffer
                EXEVersion = .DebuffDWORD
                Checksum = .DebuffDWORD
                EXEInfo = .DebuffNTString
                Index = CInt(.DebuffDWORD)
            End With
            
            'BNCSUtil requires that you initialize it before you start using it
            Call kd_init
            
            BotVars(Index).ClientToken = GetTickCount()
            
            lngDecoder = kd_create(BotVars(Index).CDKey, Len(BotVars(Index).CDKey))
            If lngDecoder = -1 Then AddChat vbRed, "[" & Index & "] CDKey could not be decoded": Closewinsock Index: Exit Sub
    
            HashLength = kd_calculateHash(lngDecoder, BotVars(Index).ClientToken, BotVars(Index).ServerToken)
            If HashLength = 0 Then AddChat vbRed, "[" & Index & "] CDKey could not be hashed": Closewinsock Index: Exit Sub
            KeyHash = String$(HashLength, vbNullChar)
    
            Call kd_getHash(lngDecoder, KeyHash)
            
            With PBuffer
                .InsertDWORD BotVars(Index).ClientToken
                .InsertDWORD EXEVersion
                .InsertDWORD Checksum
                .InsertDWORD &H1
                .InsertDWORD &H0
                .InsertDWORD Len(BotVars(Index).CDKey)
                .InsertDWORD kd_product(lngDecoder)
                .InsertDWORD kd_val1(lngDecoder)
                .InsertDWORD &H0
                .InsertNonNTString KeyHash
                .InsertNTString EXEInfo
                .InsertNTString "§[" & Owner & "]"
                .sendPacket &H51, Index
            End With
            
            Call kd_free(lngDecoder)
            
        Case &H0
            Index = CInt(Debuffer.DebuffDWORD)
            AddChat vbRed, "[" & Index & "] Failed to get game versions from BNLS."
    End Select
    
    frmMain.sckBNLS(Socket).Close
    AddChat vbRed, " - [BNLS] Disconnected"
    
End Sub

Private Sub Parse0x7D(ByVal Index As Integer)
    '(DWORD)      Cookie
    '(BYTE)       Number of Members
    '
    'For each member:
    '(STRING)     Username
    '(BYTE)       Rank
    '(BYTE)       Online Status
    '(STRING)     Location
    
    Dim x As Integer
    Dim y As Integer
    Dim Username As String
    Dim Rank As Byte
    ReDim Shamans(0)
    
    With Debuffer
    
        .RemoveDWORD
        
        For x = 0 To .DebuffBYTE
            Username = .DebuffNTString
            Rank = .DebuffBYTE
            If Rank = &H3 Then
                ReDim Preserve Shamans(y)
                Shamans(y) = Username
                y = y + 1
            End If
            .RemoveBYTE
            .RemoveNTString
        Next
    
    End With
    
End Sub

Private Sub Parse0x7C(ByVal Index As Integer)
    '(DWORD)      Cookie
    '(DWORD)      Unknown (0)
    '(STRING)     MOTD
    
    Debuffer.RemoveDWORD
    Debuffer.RemoveDWORD
    
    AddQ "The MOTD is """ & Debuffer.DebuffNTString & """"
    
End Sub

Private Sub Parse0x7A(ByVal Index As Integer)
    '(DWORD)      Cookie
    '(BYTE)       Status
    
    Debuffer.RemoveBYTE
    
    Select Case Debuffer.DebuffBYTE
        Case &H0
            AddQ "Successfully changed user's rank."
        Case &H1
            AddQ "Failed to change user's rank."
        Case &H2
            AddQ "Failed to change user's rank because they have not been in the clan for one week yet."
        Case &H7
            AddQ "Failed to change user's rank because this bot is not shaman or chieftan.", 0, Index
        Case &H8
            AddQ "Failed to change user's rank because they are a shaman or chief."
    End Select
        
End Sub

Private Sub Parse0x79(ByVal Index As Integer)
    '(DWORD)      Cookie
    '(DWORD)      Clan tag
    '(STRING)     Clan name
    '(STRING)     Inviter
    
    Dim Cookie As Long
    Dim Tag As String * 4
    Dim ClanName As String
    Dim Username As String
    
    With Debuffer
        Cookie = .DebuffDWORD
        Tag = StrReverse(.DebuffRaw(4))
        ClanName = .DebuffNTString
        Username = .DebuffNTString
    End With

    If AcceptClanInvites Then
        With PBuffer
            .InsertDWORD Cookie
            .InsertNonNTString StrReverse(Tag)
            .InsertNTString Username
            .InsertBYTE &H6
            .sendPacket &H79, Index
        End With
        AddQ "I have been invited to join " & ClanName & " (" & Tag & ") by " & Username & ".  I have automatically accepted the invite.", 0, Index
    Else
        With PBuffer
            .InsertDWORD Cookie
            .InsertNonNTString StrReverse(Tag)
            .InsertNTString Username
            .InsertBYTE &H4
            .sendPacket &H79, Index
        End With
        AddQ "I have been invited to join " & ClanName & " (" & Tag & ") by " & Username & ".  I have automatically declined the invite.", 0, Index
    End If

End Sub

Private Sub Parse0x78(ByVal Index As Integer)
    '(DWORD)      Cookie
    '(BYTE)       Status
    
    Debuffer.RemoveDWORD
    
    Select Case Debuffer.DebuffBYTE
        Case &H0
            AddQ "Removed user from clan."
        Case &H7
            AddQ "This bot is not authorized to remove that clan member.", 0, Index
        Case &H8
            AddQ "That user is not in the clan."
    End Select
    
End Sub

Private Sub Parse0x77(ByVal Index As Integer)
    '(DWORD)      Cookie
    '(BYTE)       Result
    
    Debuffer.RemoveDWORD
    
    Select Case Debuffer.DebuffBYTE
        Case &H0
            AddQ "User accepted clan invitation."
        Case &H4
            AddQ "User declined clan invitation."
        Case &H5
            AddQ "Failed to send clan invitation."
        Case &H9
            AddQ "Invitation failed because the clan is full."
    End Select
    
End Sub

Private Sub Parse0x76(ByVal Index As Integer)
    '(BYTE)       Status
    
    If Debuffer.DebuffBYTE = &H1 Then
        AddQ "This bot has successfully left its clan.", 0, Index
    Else
        AddQ "Unknown response from Battle.net when trying to leave the clan.", 0, Index
    End If
    
End Sub

Private Sub Parse0x72(ByVal Index As Integer)
    '(DWORD)      Cookie
    '(DWORD)      Clan Tag
    '(STRING)     Clan Name
    '(STRING)     Inviter's username
    '(BYTE)       Number of users being invited
    '(STRING[])   List of users being invited
    
    Dim Cookie As Long
    Dim Tag As String * 4
    Dim ClanName As String
    Dim Username As String
    
    With Debuffer
        Cookie = .DebuffDWORD
        Tag = StrReverse(.DebuffRaw(4))
        ClanName = .DebuffNTString
        Username = .DebuffNTString
    End With

    If AcceptClanCreationInvites Then
        With PBuffer
            .InsertDWORD Cookie
            .InsertNonNTString StrReverse(Tag)
            .InsertNTString Username
            .InsertBYTE &H6
            .sendPacket &H72, Index
        End With
        AddQ "I have been invited to create " & ClanName & " (" & Tag & ") by " & Username & ".  I have automatically accepted the invite.", 0, Index
    Else
        With PBuffer
            .InsertDWORD Cookie
            .InsertNonNTString StrReverse(Tag)
            .InsertNTString Username
            .InsertBYTE &H4
            .sendPacket &H72, Index
        End With
        AddQ "I have been invited to create " & ClanName & " (" & Tag & ") by " & Username & ".  I have automatically declined the invite.", 0, Index
    End If
    
End Sub

Private Sub Parse0x73(ByVal Index As Integer)
    '(DWORD)      Cookie
    '(BYTE)       Result
    
    Debuffer.RemoveDWORD
    
    Select Case Debuffer.DebuffBYTE
        Case &H0
            AddQ "Clan was successfully disbanded.", 0, Index
        Case &H2
            AddQ "Clan could not be disbanded because it is not one week old yet.", 0, Index
        Case &H7
            AddQ "Clan could not be disbanded because this bot is not the chieftan.", 0, Index
    End Select
End Sub

Private Sub Parse0x75(ByVal Index As Integer)
    '(BYTE)       Unknown (0)
    '(DWORD)      Clan tag
    '(BYTE)       Rank
    
    With Debuffer
        .RemoveBYTE
        .RemoveDWORD
        BotVars(Index).ClanRank = .DebuffBYTE
    End With
    
End Sub

Public Sub Parse0x65(ByVal Index As Integer)
    '(BYTE)       Number of Entries
    '
    'For each entry:
    '(STRING)     Account
    '(BYTE)       Status
    '(BYTE)       Location
    '(DWORD)      ProductID
    '(STRING)     Location name
    
    Dim x As Byte
    Dim Username As String
    Dim Game As String
    Dim Location As Byte
    BotVars(Index).FriendsCount = Debuffer.DebuffBYTE
    
    For x = 1 To BotVars(Index).FriendsCount
        Username = Debuffer.DebuffNTString
        Debuffer.RemoveBYTE
        Location = Debuffer.DebuffBYTE
        Debuffer.RemoveDWORD
        Game = Debuffer.DebuffNTString
        If Location And 5 Then
            If InStr(LCase(Game), "baal") > 0 Or InStr(LCase(Game), "chaos") > 0 Then
                AddQ "-+[ *" & Username & " is in the game " & Game & ". ]+-  Paradise Ops by Spasm"
                FoundRun = True
            End If
        End If
    Next
            
End Sub

Private Sub Parse0x13(ByVal Index As Integer)
    '[blank]
    
    AddChat vbRed, "The bot has been disconnected for flooding!  Please tell me this happened so I can make a few adjustments."
    Closewinsock Index
End Sub

Public Sub Parse0x19(ByVal Index As Integer)
    '(DWORD)      Style
    '(STRING)     Text
    '(STRING)     Caption
    
    Debuffer.RemoveDWORD
    AddChat vbRed, "[" & Index & "] " & Debuffer.DebuffNTString
End Sub

Public Sub Parse0x25(ByVal Index As Integer)
    '(DWORD)      Ping Value
    With PBuffer
        .InsertDWORD Debuffer.DebuffDWORD
        .sendPacket &H25, Index
    End With
End Sub

Public Sub Parse0x3A(ByVal Index As Integer)
    '(DWORD)      Result
    '(STRING)     Reason
    
    Select Case Debuffer.DebuffDWORD
        Case &H0:
            AddChat vbYellow, "[" & Index & "] Login Successful"
            With PBuffer
                .InsertNonNTString "tenb"
                .sendPacket &H14, Index 'SID_UDPPINGRESPONSE
            End With
            
            Send0xABC Index
            
       Case &H1: Send0x3D Index
       Case &H2: AddChat vbRed, "[" & Index & "] Invalid Password": BotVars(Index).Connected = True: Closewinsock Index
    End Select
End Sub

Public Sub Parse0x3D(ByVal Index As Integer)
    '(DWORD)      Status
    '(STRING)     Account name suggestion
    
    Select Case Debuffer.DebuffDWORD
       Case &H0: AddChat vbYellow, "[" & Index & "] Account creation successful": Send0x3A Index
       Case &H2: AddChat vbRed, "[" & Index & "] Username contained invalid characters": BotVars(Index).Connected = True: Closewinsock Index
       Case &H3: AddChat vbRed, "[" & Index & "] Userame contained a naughty word": BotVars(Index).Connected = True: Closewinsock Index
       Case &H4: AddChat vbRed, "[" & Index & "] Username already exists": BotVars(Index).Connected = True: Closewinsock Index
       Case &H6: AddChat vbRed, "[" & Index & "] Username did not contain enough alphanumeric characters": BotVars(Index).Connected = True: Closewinsock Index
    End Select
End Sub

Public Sub Parse0x50(ByVal Index As Integer)
    '(DWORD)      Logon Type (NOT NEEDED)
    '(DWORD)      Server Token
    '(DWORD)      UDPValue** (NOT NEEDED)
    '(FILETIME)   MPQ filetime (NOT NEEDED)
    '(STRING)     IX86ver filename (MPQ Number is contained here)
    '(STRING)     ValueString (HashCommand)
    
    Dim MPQNumber As Byte
    Dim HashCommand As String
    Dim MPQName As String
    Dim Filetime As String
    
    ' These values are used in 0x50 for hashing
    Debuffer.RemoveDWORD
    BotVars(Index).ServerToken = Debuffer.DebuffDWORD
    Debuffer.RemoveDWORD
    Debuffer.RemoveFILETIME
    MPQName = Debuffer.DebuffNTString
    HashCommand = Debuffer.DebuffNTString
    MPQNumber = Val#(Left$(Mid$(MPQName, InStr(1, MPQName, "-IX86-") + 6), InStr(1, LCase$(Mid$(MPQName, InStr(1, MPQName, "-IX86-") + 6)), ".mpq") - 1))
    AddChat vbYellow, "MPQ Name: " & MPQName
    
    Select Case BotVars(Index).Product
        Case "WAR3", "D2DV"
            If WAR3BNLS Then
                SendBNLS0x1A Filetime, MPQName, HashCommand, Index, 0
                Exit Sub
            End If
        Case "STAR", "SEXP", "W2BN"
            If STARBNLS Then
                SendBNLS0x1A Filetime, MPQName, HashCommand, Index, 1
                Exit Sub
            End If
    End Select
    
    Send0x51 HashCommand, MPQNumber, Index
End Sub

Public Sub Parse0x51(ByVal Index As Integer)
    '(DWORD)      Result
    '(STRING)     Additional Information
    
    Select Case Debuffer.DebuffDWORD
        Case &H0:
            If BotVars(Index).lngNLS <> 0 Then nls_free (BotVars(Index).lngNLS)
            AddChat vbYellow, "[" & Index & "] Version accepted"
            If BotVars(Index).Product = "WAR3" Then
                Send0x53 Index
            Else
                Send0x3A Index
            End If
            Exit Sub
        Case &H100: AddChat vbRed, "[" & Index & "] Game version must be upgraded with the MPQ specified in " & Debuffer.DebuffNTString
        Case &H101: AddChat vbRed, "[" & Index & "] The game version is invalid."
        Case &H102: AddChat vbRed, "[" & Index & "] Game version must be downgraded with the MPQ specified " & Debuffer.DebuffNTString
        Case &H200: AddChat vbRed, "[" & Index & "] CD key is invalid"
        Case &H201: AddChat vbRed, "[" & Index & "] CD key in use by " & Debuffer.DebuffNTString: SetTimer frmMain.hWnd, CLng(Index + 2000), 300000, AddressOf Reconnect: AddChat vbYellow, "[" & Index & "] Bot will try to reconnect again in 5 minutes."
        Case &H202: AddChat vbRed, "[" & Index & "] CD key is banned"
        Case &H203: AddChat vbRed, "[" & Index & "] Your CD key is for the wrong product"
        Case Else: AddChat vbRed, "[" & Index & "] Unknown error in 0x51": AddChat vbRed, "Additional information: " & Debuffer.DebuffNTString
    End Select
    BotVars(Index).Connected = True
    Closewinsock Index
End Sub

Public Sub Parse0x52(ByVal Index As Integer)
    ' (DWORD)      Status
    
    Select Case Debuffer.DebuffDWORD
      Case &H0:
        AddChat vbYellow, "[" & Index & "] Account creation successful"
        Send0x53 Index
      Case &H6: AddChat vbRed, "[" & Index & "] Name already exists": BotVars(Index).Connected = True: Closewinsock Index
      Case &H7: AddChat vbRed, "[" & Index & "] Name is blank/not long enough": BotVars(Index).Connected = True: Closewinsock Index
      Case &H8: AddChat vbRed, "[" & Index & "] Name contains invalid characters": BotVars(Index).Connected = True: Closewinsock Index
      Case &H9: AddChat vbRed, "[" & Index & "] Name contains banned words": BotVars(Index).Connected = True: Closewinsock Index
      Case &HA: AddChat vbRed, "[" & Index & "] Name needs more alphanumeric characters": BotVars(Index).Connected = True: Closewinsock Index
      Case &HB: AddChat vbRed, "[" & Index & "] Name cannot have adjacent puncuation": BotVars(Index).Connected = True: Closewinsock Index
      Case &HC: AddChat vbRed, "[" & Index & "] Name has too much puncuation": BotVars(Index).Connected = True: Closewinsock Index
      Case Else: AddChat vbRed, "[" & Index & "] Failed": Closewinsock Index
    End Select
End Sub

Public Sub Parse0x53(ByVal Index As Integer)
    ' (DWORD)      Status
    ' (BYTE[32])   Salt (s)
    ' (BYTE[32])   Server Key (B)
    
    Dim Status2 As Byte
    Dim Salt As String
    Dim ServerKey As String

    Status2 = Debuffer.DebuffDWORD
    Salt = Debuffer.DebuffRaw(32)
    ServerKey = Debuffer.DebuffRaw(32)
    
    Select Case Status2
        Case 0: Send0x54 Salt, ServerKey, Index
        Case 1:
            AddChat vbRed, "[" & Index & "] Account doesn't exist"
            AddChat vbYellow, "Creating account..."
            Send0x52 Index
        Case 5: AddChat vbRed, "[" & Index & "] Account must be upgraded": Closewinsock Index
        Case Else: AddChat vbRed, "[" & Index & "] Account creation failed": Closewinsock Index
    End Select
End Sub

Public Sub Parse0x54(ByVal Index As Integer)
    ' (DWORD)      Status
    ' (BYTE[20])   Server Password Proof (M2)
    ' (STRING)     Additional information
    
    Select Case Debuffer.DebuffDWORD
        Case &H0:
            Send0xABC Index
        Case &H2: AddChat vbRed, "[" & Index & "] Incorrect password": BotVars(Index).Connected = True: Closewinsock Index
        Case &HE:
            Send0x59 Index
            Send0xABC Index
            
      Case Else: AddChat vbRed, "[" & Index & "] Failed": BotVars(Index).Connected = True: Closewinsock Index
    End Select
End Sub

Public Sub Parse0x0A(ByVal Index As Integer)
    '(STRING)     Unique Username
    '(STRING)     Statstring
    '(STRING)     Account name
    Dim Username As String
    Username = Debuffer.DebuffNTString
    AddChat vbYellow, "[" & Index & "] Logged in as: ", &HFFFF00, Username
    BotVars(Index).Username = Username
    BotVars(Index).Connected = True
    If Index = 0 Then
        Status = "Online"
        If GUI Then frmMain.lblStatus.Caption = "Status: " & Status
        Connected = True
    End If
End Sub

Public Sub Parse0x59(ByVal Index As Integer)
    '[blank]
    Send0x59 Index
End Sub

Public Sub Parse0x0F(ByVal Index As Integer)
    '(DWORD)      Event ID
    '(DWORD)      User's Flags
    '(DWORD)      Ping
    '(DWORD)      IP Address (Defunct)
    '(DWORD)      Account number (Defunct)
    '(DWORD)      Registration Authority (Defunct)
    '(STRING)     Username
    '(STRING)     Text
    
    Dim EventID As Long
    Dim Flags As Long
    Dim Ping As Long
    Dim Username As String
    Dim Text As String
        
    With Debuffer
        EventID = .DebuffDWORD
        Flags = .DebuffDWORD
        Ping = .DebuffDWORD
        .RemoveDWORD
        .RemoveDWORD
        .RemoveDWORD
        Username = .DebuffNTString
        Text = .DebuffNTString
    End With
    
    If Username = vbNullString Then Exit Sub
        
    If BotVars(Index).Product = "D2DV" Then Username = Mid(Username, InStr(Username, "*") + 1)
    
    Select Case EventID
        Case &H1:  'EID_SHOWUSER
            UserInChannel Flags, Ping, Username, Index
        Case &H2:  'EID_JOIN
            UserJoin Flags, Ping, Username, Index
        Case &H3:  'EID_LEAVE
            UserLeave Username, Index
        Case &H5:  'EID_TALK
            UserTalk Username, Text, Flags, Index
        Case &H7:  'EID_CHANNEL
            JoinedChannel Text, Index
        Case &H17: 'EID_EMOTE
            UserEmote Username, Text, Index
        Case &H6:  'EID_BROADCAST
            ServerBroadcast Text, Index
        Case &H12: 'EID_INFO
            Info Text, Index
        Case &H4:  'EID_WHISPER
            UserWhisper Username, Text, Index
        Case &H9:  'EID_USERFLAGS
            FlagsUpdate Flags, Ping, Username, Index
        Case &HA:  'EID_WHISPERSENT
            WhisperSent Username, Text, Index
        Case &HD:  'EID_CHANNELFULL
            ChannelFull Text, Index
        Case &HE:  'EID_CHANNELDOESNOTEXIST
            ChannelDoesNotExist Text, Index
        Case &HF:  'EID_CHANNELRESTRICTED
            ChannelRestricted Text, Index
        Case &H13: 'EID_ERROR
            BNETError Text, Index
    End Select
        
End Sub


Public Sub ParseUnknown(PacketID As Byte, Data As String)
    ' This Sub shows us packets that the bot does not recognize in an easier to read
    ' hex form

    AddChat vbRed, "Unknown Packet: 0x" & Hex$(PacketID)
    AddChat vbYellow, PBuffer.DebugOutput(Data)
End Sub



'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'//////////////////////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



Public Sub SendBNLS0x1A(ByVal Filetime As String, ByVal MPQName As String, ByVal HashCommand As String, ByVal Index As Integer, ByVal Socket As Integer)

    Dim ProductID As Long
    
    AddChat vbYellow, " - [BNLS] Requesting version information"
    
    Select Case BotVars(Index).Product
        Case "STAR"
            ProductID = &H1
        Case "SEXP"
            ProductID = &H2
        Case "W2BN"
            ProductID = &H3
        Case "D2DV"
            ProductID = &H4
        Case "WAR3"
            ProductID = &H7
        Case Else
            AddChat vbRed, "[" & Index & "] You specified an invalid product.  Please check your configuration."
    End Select
    
    With PBuffer
        .InsertDWORD ProductID
        .InsertDWORD 0
        .InsertDWORD CLng(Index)
        .InsertNonNTString Filetime
        .InsertNTString MPQName
        .InsertNTString HashCommand
        .sendBNLSPacket &H1A, Socket
    End With
    
End Sub


Public Sub Send0xABC(ByVal Index As Integer)
    '(STRING)     *Username
    '(STRING)     **Statstring.
    
    '(DWORD)      Product ID
    
    '(DWORD)      Flags
    '(STRING)     Channel
    
    With PBuffer
        .InsertNTString BotVars(Index).Username
        .InsertBYTE &H0
        .sendPacket &HA, Index
        
        .InsertNonNTString StrReverse(BotVars(Index).Product)
        .sendPacket &HB, Index

        .InsertDWORD &H2
        .InsertNTString Home
        .sendPacket &HC, Index
    End With
End Sub

Public Sub Send0x3A(ByVal Index As Integer)
    '(DWORD)      Client Token
    '(DWORD)      Server Token
    '(DWORD[5])   Password Hash
    '(STRING)     Username
    
    AddChat vbYellow, "[" & Index & "] Logging in..."
    With PBuffer
        .InsertDWORD BotVars(Index).ClientToken
        .InsertDWORD BotVars(Index).ServerToken
        .InsertNonNTString doubleHashPassword(BotVars(Index).Password, BotVars(Index).ClientToken, BotVars(Index).ServerToken)
        .InsertNTString BotVars(Index).Username
        .sendPacket &H3A, Index
    End With
End Sub

Public Sub Send0x3D(ByVal Index As Integer)
    '(DWORD[5])   Password hash
    '(STRING)     Username
    
    With PBuffer
        .InsertNonNTString hashPassword(BotVars(Index).Password)
        .InsertNTString BotVars(Index).Username
        .sendPacket &H3D, Index
    End With
End Sub

Public Sub Send0x50(ByVal Index As Integer)
    '(DWORD)      Protocol ID (0)
    '(DWORD)      Platform ID (This is always 68XI for PCs)
    '(DWORD)      Product ID (This is STAR, SEXP, WAR3, etc, reversed)
    '(DWORD)      Version Byte (AKA the Verbyte)
    '(DWORD)      Product language UNUSED
    '(DWORD)      Local IP for NAT compatibility* UNUSED
    '(DWORD)      Time zone bias* UNUSED
    '(DWORD)      Locale ID* UNUSED
    '(DWORD)      Language ID* UNUSED
    '(STRING)     Country abreviation
    '(STRING)     Country

    
    Select Case BotVars(Index).Product
        Case "STAR", "SEXP", "W2BN", "D2DV", "WAR3":
        Case Else:
            AddChat vbRed, "[" & Index & "] Invalid product"
            Closewinsock Index
            Exit Sub
    End Select
    
    With PBuffer
        .InsertDWORD &H0
        .InsertNonNTString "68XI"
        .InsertNonNTString StrReverse(BotVars(Index).Product)
        .InsertDWORD Val#(GetVerByte(BotVars(Index).Product))
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertNTString "USA"
        .InsertNTString "United States"
        .sendPacket &H50, Index
    End With
End Sub

Public Sub Send0x51(HashCommand As String, MPQNumber As Byte, Index As Integer)
    '(DWORD)      Client Token
    '(DWORD)      EXE Version
    '(DWORD)      EXE Hash
    '(DWORD)      Number of keys in this packet
    '(BOOLEAN)    Using Spawn (32-bit)

    'For Each Key:
    '(DWORD)      Key Length
    '(DWORD)      CD key's product value
    '(DWORD)      CD key's public value
    '(DWORD)      Unknown (0)
    '(DWORD[5])   Hashed Key Data

    '(STRING)     Exe Information
    '(STRING)     CD Key owner name
    
    
    
    Dim Checksum As Long
    Dim CRevision As Long
    Dim lngDecoder As Long
    Dim HashLength As Long
    Dim EXEInfo As String
    Dim KeyHash As String
    Dim EXEVersion As Long
    Dim Gamefile(2) As String
    
    Select Case BotVars(Index).Product
        Case "STAR", "SEXP":
            Gamefile(0) = App.Path & "\Hashes\STAR\Starcraft.exe"
            Gamefile(1) = App.Path & "\Hashes\STAR\Storm.dll"
            Gamefile(2) = App.Path & "\Hashes\STAR\Battle.snp"
       
        Case "W2BN":
            Gamefile(0) = App.Path & "\Hashes\W2BN\Warcraft II BNE.exe"
            Gamefile(1) = App.Path & "\Hashes\W2BN\Storm.dll"
            Gamefile(2) = App.Path & "\Hashes\W2BN\Battle.snp"
          
        Case "D2DV":
            Gamefile(0) = App.Path & "\Hashes\D2DV\Game.exe"
            Gamefile(1) = App.Path & "\Hashes\D2DV\Bnclient.dll"
            Gamefile(2) = App.Path & "\Hashes\D2DV\D2Client.dll"
       
        Case "WAR3":
            Gamefile(0) = App.Path & "\Hashes\WAR3\War3.exe"
            Gamefile(1) = App.Path & "\Hashes\WAR3\Storm.dll"
            Gamefile(2) = App.Path & "\Hashes\WAR3\Game.dll"
    End Select
    
    'BNCSUtil requires that you initialize it before you start using it
    Call kd_init
    
    AddChat vbYellow, "[" & Index & "] Checking version..."
    
    CRevision = checkRevision(HashCommand, Gamefile(0), Gamefile(1), Gamefile(2), CLng(MPQNumber), Checksum)
    If CRevision = 0 Then
        AddChat vbRed, "[" & Index & "] Hashes did not pass checkrevision"
        Closewinsock Index
        Exit Sub
    End If
    
    EXEVersion = getExeInfo(Gamefile(0), EXEInfo)
    EXEInfo = PBuffer.KillNull(EXEInfo)
    BotVars(Index).ClientToken = GetTickCount()
    
    lngDecoder = kd_create(BotVars(Index).CDKey, Len(BotVars(Index).CDKey))
    If lngDecoder = -1 Then AddChat vbRed, "[" & Index & "] CDKey could not be decoded": Closewinsock Index: Exit Sub
    
    HashLength = kd_calculateHash(lngDecoder, BotVars(Index).ClientToken, BotVars(Index).ServerToken)
    If HashLength = 0 Then AddChat vbRed, "[" & Index & "] CDKey could not be hashed": Closewinsock Index: Exit Sub
    KeyHash = String$(HashLength, vbNullChar)
    
    Call kd_getHash(lngDecoder, KeyHash)
    
    With PBuffer
        .InsertDWORD BotVars(Index).ClientToken
        .InsertDWORD EXEVersion
        .InsertDWORD Checksum
        .InsertDWORD &H1
        .InsertDWORD &H0
        .InsertDWORD Len(BotVars(Index).CDKey)
        .InsertDWORD kd_product(lngDecoder)
        .InsertDWORD kd_val1(lngDecoder)
        .InsertDWORD &H0
        .InsertNonNTString KeyHash
        .InsertNTString EXEInfo
        .InsertNTString "§" & Owner
        .sendPacket &H51, Index
    End With
    
    Call kd_free(lngDecoder)
End Sub

Public Sub Send0x52(ByVal Index As Integer)
    ' (BYTE[32])   Salt (s)
    ' (BYTE[32])   Verifier (v)
    ' (STRING)     Username
    
    Dim Buffer As String, buflen As Long
    
    buflen = 65 + Len(BotVars(Index).Username)
    Buffer = String$(buflen, vbNullChar)
    If (nls_account_create(BotVars(Index).lngNLS, Buffer, buflen) = 0) Then
        Closewinsock Index
        AddChat vbRed, "[" & Index & "] Failed to make account creation packet."
        Exit Sub
    End If
    
    With PBuffer
        .InsertNonNTString Buffer
        .sendPacket &H52, Index
    End With
End Sub

Public Sub Send0x53(ByVal Index As Integer)
    '(BYTE[32])   Client Key ('A')
    '(STRING)     Username
    
    Dim strReturn As String * 32
    
    If BotVars(Index).lngNLS = 0 Then BotVars(Index).lngNLS = nls_init(BotVars(Index).Username, BotVars(Index).Password)
    If BotVars(Index).lngNLS = 0 Then AddChat vbRed, "[" & Index & "] Failed to initialize NLS"
    Call nls_get_A(BotVars(Index).lngNLS, strReturn)
    
    With PBuffer
        .InsertNonNTString strReturn
        .InsertNTString BotVars(Index).Username
        .sendPacket &H53, Index
    End With
End Sub

Public Sub Send0x54(Salt As String, ServerKey As String, Index As Integer)
    ' (BYTE[20])   Client Password Proof (M1)
    
    Dim strReturn As String * 20
    
    Call nls_get_M1(BotVars(Index).lngNLS, strReturn, ServerKey, Salt)
    
    With PBuffer
        .InsertNonNTString strReturn
        .sendPacket &H54, Index
    End With
End Sub


Public Sub Send0x59(ByVal Index As Integer)
    With PBuffer
        .InsertNTString Email
        .sendPacket &H59, Index 'SID_SETEMAIL
    End With
    AddChat vbYellow, "[" & Index & "] Set registration email"
End Sub
