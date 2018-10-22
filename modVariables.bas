Attribute VB_Name = "modVariables"
' Another module I always create is one to store all my public variables that need
' to be shared among forms and modules
Option Explicit

Public PBuffer As New clsPacketBuffer
Public Debuffer As New clsPacketDebuffer

Public Type Bot
    Username As String
    CDKey As String
    Password As String
    Product As String * 4
    ClientToken As Long
    ServerToken As Long
    lngNLS As Long
    Connected As Boolean
    LastTick As Long
    SentBytes As Long
    TimerEnabled As Boolean
    HasOps As Boolean
    strBuffer As String
    FriendsCount As Byte
    ClanRank As Byte
End Type

Public Type UserlistUser
    Username As String
    Access As Integer
End Type

Public Type SquelchedUser
    Username As String
    Bot As Integer
End Type

Public Type ShitlistedUser
    Username As String
    Reason As String
End Type

Public BotVars() As Bot
Public Channel As String
Public Email As String
Public Home As String
Public Server As String
Public Owner As String
Public MaxQueue As Integer
Public Connected As Boolean
Public GUI As Boolean
Public HasOps As Boolean
Public IPBan As Boolean
Public IPBans() As SquelchedUser
Public Userlist() As UserlistUser
Public AssumeRealms As Boolean
Public Realm As String
Public War3Realm As String
Public Safelist() As String
Public Trigger As String * 1
Public Idle As Boolean
Public IdleTime As Integer
Public IdleMessage As String
Public IdleKick As Boolean
Public IdleKickTime As Integer
Public Shitlist() As ShitlistedUser
Public PhraseBans() As String
Public AnnounceRuns As Boolean
Public FoundRun As Boolean
Public Status As String
Public STARBNLS As Boolean
Public WAR3BNLS As Boolean
Public STARBNLSServer As String
Public WAR3BNLSServer As String
Public AcceptClanInvites As Boolean
Public AcceptClanCreationInvites As Boolean
Public Shamans() As String

' This is a Windows API function which is used primarily in the packet buffer
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' This Windows API function will tell you how long the application has been running
Public Declare Function GetTickCount Lib "kernel32" () As Long
