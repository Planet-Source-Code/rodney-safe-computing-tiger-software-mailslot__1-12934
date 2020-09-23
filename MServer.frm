VERSION 5.00
Begin VB.Form MServerFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server - Mailslot"
   ClientHeight    =   1695
   ClientLeft      =   75
   ClientTop       =   1590
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   3720
   Begin VB.Timer Timer 
      Interval        =   250
      Left            =   1980
      Top             =   840
   End
   Begin VB.TextBox MsgInTxt 
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   1260
      Width           =   3375
   End
   Begin VB.CommandButton MsgOutSendCmd 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   2700
      TabIndex        =   1
      Top             =   420
      Width           =   855
   End
   Begin VB.TextBox MsgOutTxt 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   2355
   End
   Begin VB.Label Label2 
      Caption         =   "Message in"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message out"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   1335
   End
End
Attribute VB_Name = "MServerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=====================
'Rodney Safe Computing
'Created by Rodney Godfried
'18-08-1999
'=====================


Private Sub Form_Unload(Cancel As Integer)
    Call Quit(Cancel)
End Sub

Private Sub MsgOutSendCmd_Click()
    '
    ' Send message from user interface
    '
    If ClientSlot <> -1 Then
        Call MailSlotWrite(ClientSlot, MsgOutTxt.Text)
    End If
End Sub

Private Sub Timer_Timer()
    Dim Text As String

    '
    ' Check for inbound messages
    '
    Text = MailSlotRead(ServerSlot)
    If Len(Text) Then
        ' check for a client login message
        If Left$(Text, Len(LoginCommand)) = LoginCommand Then
            ClientSlotName = Trim(Mid$(Text, Len(LoginCommand) + 1))
            '
            ' Create local representation of client mailslot.
            ' This will be used to send messages to client's
            ' local mailslot.
            '
            ClientSlot = CreateFileNoSecurity(ClientSlotName, _
                GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, _
                FILE_ATTRIBUTE_NORMAL, 0)
            If ServerSlot = -1 Then
                MsgBox "Failed to open server mailslot"
            End If
        Else
            ' display new message
            MsgInTxt = Text
        End If
    End If
End Sub

