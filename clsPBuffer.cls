VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacketBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================
'Alpha and Omega's PacketBuffer Class
'Coded by Fleet-
'=================================================
Option Explicit
Private Buffer As String

Public Sub AddBuffer(ByVal Data As String)
  Buffer = Buffer & Data
End Sub

Public Function InsertDWORDArray(Data() As Long)
    Dim i As Integer
    For i = LBound(Data) To UBound(Data) Step 1
        Buffer = Buffer & MakeDWORD(Data(i))
    Next i
End Function

Public Function InsertDWORD(ByVal Data As Long)
    Buffer = Buffer & MakeDWORD(Data)
End Function

Public Function InsertData(ByVal Data As String)
    Buffer = Buffer & Data
End Function

Public Function InsertBYTE(ByVal Data As Integer)
    Buffer = Buffer & Chr$(Data)
End Function

Public Sub InsertBytes(ByVal Data As String)
    Dim i As Long
    Dim enqueueer As String

    For i = 1 To Len(Data) Step 3
        enqueueer = enqueueer & Chr$(Val#("&h0" & Mid$(Data, i, 2)))
    Next i
    Buffer = Buffer & enqueueer
End Sub

Public Function InsertNTString(ByVal Data As String)
    Buffer = Buffer & Data & vbNullChar
End Function

Public Function InsertNonNTString(ByVal Data As String)
    Buffer = Buffer & Data
End Function

Private Function MakeDWORD(ByVal Value As Long) As String
    Dim Result As String * 4
    CopyMemory ByVal Result, Value, 4
    MakeDWORD = Result
End Function

Private Function MakeWORD(ByVal Value As Integer) As String
    Dim Result As String * 2
    CopyMemory ByVal Result, Value, 2
    MakeWORD = Result
End Function

Public Sub sendPacket(ByVal PacketID As Byte, ByVal Index As Integer)
       
    Buffer = Chr$(&HFF) & Chr$(PacketID) & MakeWORD(Len(Buffer) + 4) & Buffer
    frmMain.sckBNET(Index).SendData Buffer
    ClearBuffer
    
End Sub

Public Sub sendBNLSPacket(ByVal PacketID As Byte, ByVal Socket As Integer)
     
    Do While frmMain.sckBNLS(Socket).State <> sckConnected
        DoEvents
    Loop
    
    Buffer = MakeWORD(Len(Buffer) + 3) & Chr$(PacketID) & Buffer
    frmMain.sckBNLS(Socket).SendData Buffer
    ClearBuffer
    
End Sub

Public Function InsertWORD(ByVal Data As Integer)
    Buffer = Buffer & MakeWORD(Data)
End Function
Public Sub ClearBuffer()
    Buffer = vbNullString
End Sub

Public Function MakeBuffer(ByVal sData As String)
  Buffer = sData
End Function

Public Function GetWORD(ByVal Data As String) As Long
Dim lReturn As Long
    Call CopyMemory(lReturn, ByVal Data, 2)
    GetWORD = lReturn
End Function

Public Function KillNull(ByVal Text As String) As String
    Dim i As Integer
    i = InStr(1, Text, Chr(0))
    If i = 0 Then
        KillNull = Text
        Exit Function
    End If
    KillNull = Left(Text, i - 1)
End Function

Public Function DebugOutput(ByVal sIn As String) As String
     Dim X1 As Long, Y1 As Long
     Dim iLen As Long, iPos As Long
     Dim sB As String, ST As String
     Dim sOut As String
     Dim Offset As Long, sOffset As String

     iLen = Len(sIn)
     If iLen = 0 Then Exit Function
     sOut = ""
     Offset = 0
     For X1 = 0 To ((iLen - 1) \ 16)
         sOffset = Right$("0000" & Hex(Offset), 4)
         sB = String(48, " ")
         ST = "................"
         For Y1 = 1 To 16
             iPos = 16 * X1 + Y1
             If iPos > iLen Then Exit For
             Mid(sB, 3 * (Y1 - 1) + 1, 2) = Right("00" & Hex(Asc(Mid(sIn, iPos, 1))), 2) & " "
             Select Case Asc(Mid(sIn, iPos, 1))
             Case 0, 9, 10, 13: Mid(ST, Y1, 1) = "."
             Case Else
                 Mid(ST, Y1, 1) = Mid(sIn, iPos, 1)
             End Select
         Next Y1
         If Len(sOut) > 0 Then sOut = sOut & vbCrLf
         sOut = sOut & sOffset & ":  "
         sOut = sOut & sB & "  " & ST
         Offset = Offset + 16
     Next X1
     DebugOutput = sOut
 End Function

