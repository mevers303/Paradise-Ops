VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paradise Ops Configuration"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Miscelaneous Settings"
      Height          =   1815
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtIdleTime 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Text            =   "500"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkIdle 
         Caption         =   "Idle Messages"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Send idle messages."
         Top             =   360
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moderation Settings"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtIdleKickTime 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Text            =   "500"
         ToolTipText     =   "Kick users who are idle for more than this many seconds."
         Top             =   2040
         Width           =   495
      End
      Begin VB.CheckBox chkIdleKick 
         Caption         =   "Idle Kick"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Kick users that are idle and not safelisted?"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtTrigger 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "."
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox chkIPBan 
         Caption         =   "IP Banning"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Bot will automatically IP ban users when they are banned."
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkAssumeRealms 
         Caption         =   "Assum Realms"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Bot will automatically assume Starcraft realms."
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtMaxQueue 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "10"
         ToolTipText     =   "Maximum number of pending outgoing messages.  This is on a per bot basis."
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Idle Kick Time"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Trigger"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Maximum Queue"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

