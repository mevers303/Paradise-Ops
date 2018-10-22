Attribute VB_Name = "modParsePackets"
Option Explicit

Public Sub ParseData(data As String)
    ' This Sub parses inbound packets recieved by Winsock and decides which packet
    ' is which and then directs it to the appropriate Sub to be further broken down
    ' and used by the bot
    Dim PacketID As Integer
    
    PacketID = Asc(Mid(data, 2, 1))
    data = Mid(data, 5)
    
    Debuffer.DebuffPacket data    'Give the debuffer the packet to hold on to

    Select Case PacketID
        Case &H25: Parse0x25 data 'SID_PING
        Case &H3A: Parse0x3A data 'SID_LOGONRESPONSE2
        Case &H3D: Parse0x3D data 'SID_CREATEACCOUNT
        Case &H50: Parse0x50 data 'SID_AUTH_INFO
        Case &H51: Parse0x51 data 'SID_AUTH_CHECK
        Case &H52: Parse0x52 data 'SID_AUTH_ACCOUNTCREATE
        Case &H53: Parse0x53 data 'SID_AUTH_ACCOUNTLOGON
        Case &H54: Parse0x54 data 'SID_AUTH_ACCOUNTLOGONPROOF
        Case Else: ParseUnknown PacketID, data
    End Select
    
    Debuffer.Clear                'Clear the debuffer when we're done
    
End Sub

Public Sub Parse0x25(data As String)
    '(DWORD)      Ping Value
    
    ' By simply sending the DWORD value in this packet back to Battle.Net, you can
    ' avoid your connection with the server timing out
    Dim Ping As Long
    Ping = Debuffer.DebuffDWORD
    
    With PBuffer
        .InsertDWORD Ping
        .SendPacket &H25
    End With
End Sub

Public Sub Parse0x3A(data As String)
    '(DWORD)      Result
    '(STRING)     Reason
    
    ParseUnknown data
    Select Case Debuffer.DebuffDWORD
        Case &H0:
            AddChat vbBlack, "[BNET] Login Successful"
            With PBuffer
                .InsertNonNTString "tenb"
                .SendPacket &H14 'SID_UDPPINGRESPONSE
            End With
            Send0x0A
            Send0x0C
            
       Case &H1: Send0x3D
       Case &H2: AddChat vbBlack, "[BNET] Invalid password": CloseWinsock
    End Select
End Sub

Public Sub Parse0x3D(data As String)
    '(DWORD)      Status
    '(STRING)     Account name suggestion
    
    Select Case Debuffer.DebuffDWORD
       Case &H0: AddChat vbBlack, "[BNET] Account Created! Logging in": Send0x3A
       Case &H2: AddChat vbBlack, "[BNET] Name contained invalid characters": CloseWinsock
       Case &H3: AddChat vbBlack, "[BNET] Name contained a banned word": CloseWinsock
       Case &H4: AddChat vbBlack, "[BNET] Account already exists": CloseWinsock
       Case &H6: AddChat vbBlack, "[BNET] Name did not contain enough alphanumeric characters": CloseWinsock
    End Select
End Sub

Public Sub Parse0x50(data As String)
    '(DWORD)      Logon Type (NOT NEEDED)
    '(DWORD)      Server Token
    '(DWORD)      UDPValue** (NOT NEEDED)
    '(FILETIME)   MPQ filetime (NOT NEEDED)
    '(STRING)     IX86ver filename (MPQ Number is contained here)
    '(STRING)     ValueString (HashCommand)
    
    Dim MPQNumber As Byte
    Dim HashCommand As String
    Dim MPQName As String
    
    ' These values are used in 0x50 for hashing
    Debuffer.RemoveDWORD
    ServerToken = Debuffer.DebuffDWORD
    Debuffer.RemoveDWORD
    Debuffer.RemoveFILETIME
    MPQName = Debuffer.DebuffNTString
    HashCommand = Debuffer.DebuffNTString
    MPQNumber = CInt(Mid$(MPQName, 10, 1))
    
    Send0x51 HashCommand, MPQNumber
End Sub

Public Sub Parse0x51(data As String)
    '(DWORD)      Result
    '(STRING)     Additional Information
    
    Select Case Debuffer.DebuffDWORD
        Case &H0:
            If lngNLS <> 0 Then nls_free (lngNLS)
            AddChat vbBlack, "[BNET] Client version accepted!"
            If Product = "WAR3" Then
                Send0x53
            Else
                Send0x3A
            End If
            Exit Sub
        Case &H100: AddChat vbBlack, "[BNET] Game version must be upgraded with the MPQ specified in " & Debuffer.DebuffNTString
        Case &H101: AddChat vbBlack, "[BNET] The game version is invalid."
        Case &H102: AddChat vbBlack, "[BNET] Game version must be downgraded with the MPQ specified " & Debuffer.DebuffNTString
        Case &H200: AddChat vbBlack, "[BNET] CD-key is invalid"
        Case &H201: AddChat vbBlack, "[BNET] CD-key in use by " & Debuffer.DebuffNTString
        Case &H202: AddChat vbBlack, "[BNET] CD-key is disabled"
        Case &H203: AddChat vbBlack, "[BNET] Your CD-key is for the wrong product"
        Case Else: AddChat vbBlack, "[BNET] Unknown error occurred in 0x51"
    End Select
    CloseWinsock
End Sub

Public Sub Parse0x52(data As String)
    ' (DWORD)      Status
    
    Select Case Debuffer.DebuffDWORD
      Case &H0:
        AddChat vbBlack, "[BNET] Successfully created account"
        Send0x53
      Case &H6: AddChat vbBlack, "[BNET] Name already exists": CloseWinsock
      Case &H7: AddChat vbBlack, "[BNET] Name is blank/not long enough": CloseWinsock
      Case &H8: AddChat vbBlack, "[BNET] Name contains invalid characters": CloseWinsock
      Case &H9: AddChat vbBlack, "[BNET] Name contains banned words": CloseWinsock
      Case &HA: AddChat vbBlack, "[BNET] Name needs more alphanumeric characters": CloseWinsock
      Case &HB: AddChat vbBlack, "[BNET] Name cannot have adjacent puncuation": CloseWinsock
      Case &HC: AddChat vbBlack, "[BNET] Name has too much puncuation": CloseWinsock
      Case Else: AddChat vbBlack, "[BNET] Failed": CloseWinsock
    End Select
End Sub

Public Sub Parse0x53(data As String)
    ' (DWORD)      Status
    ' (BYTE[32])   Salt (s)
    ' (BYTE[32])   Server Key (B)
    
    Dim Status As Byte
    Dim Salt As String
    Dim ServerKey As String

    Status = Debuffer.DebuffDWORD
    Salt = Debuffer.DebuffRaw(32)
    ServerKey = Debuffer.DebuffRaw(32)
    
    Select Case Status
        Case 0: Send0x54 Salt, ServerKey
        Case 1:
            AddChat vbRed, "[BNET] Account doesn't exist"
            Send0x52
        Case 5: AddChat vbRed, "[BNET] Account must be upgraded": CloseWinsock
        Case Else: AddChat vbRed, "[BNET] Failed": CloseWinsock
    End Select
End Sub

Public Sub Parse0x54(data As String)
    ' (DWORD)      Status
    ' (BYTE[20])   Server Password Proof (M2)
    ' (STRING)     Additional information
    
    AddChat vbBlack, "Parse0x54"
    Select Case Debuffer.DebuffDWORD
        Case &H0:
            AddChat vbBlack, "[BNET] Successfully logged in!"
            Send0x0A
            Send0x0C
            
        Case &H2: AddChat vbBlack, "[BNET] Invalid password"
        
        Case &HE:
            Send0x59
            Send0x0A
            Send0x0C
            
      Case Else: AddChat vbBlack, "[BNET] Failed": CloseWinsock
    End Select
End Sub

Public Sub ParseUnknown(PacketID As Byte, data As String)
    ' This Sub shows us packets that the bot does not recognize in an easier to read
    ' hex form

    AddChat vbBlack, "Unidentified Packet: 0x" & Hex$(PacketID)
    AddChat vbBlack, PBuffer.DebugOutput(data)
End Sub
