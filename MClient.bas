Attribute VB_Name = "MClientMod"
Option Explicit
'=====================
'Rodney Safe Computing
'Created by Rodney Godfried
'18-08-1999
'=====================

Global ClientSlot As Long
' ClientSlotName is the name of the local mailslot.
' It must begin with "\\" and be followed by the machine
' name or "." for local, and must be followed by "\mailslot\".
' After that, you can make up any valid directory tree and
' filename.
Global Const ClientSlotName As String = "\\.\mailslot\RodneyGodfried\mclient"
' ClientSlotNameForServer is the name the server uses to
' address the local mailslot.
' Over a network, the name would have the local machine name
' instead of the dot (".").
Global Const ClientSlotNameForServer As String = "\\.\mailslot\RodneyGodfried\mclient"

Global ServerSlot As Long
Global Const ServerSlotName As String = "\\.\mailslot\RodneyGodfried\mserver"

' LoginCommand is a phrase that the server recognizes as a command
' for a client that is logging in (registering its return mailslot).
Global Const LoginCommand As String = "Mailslot-Login:"
Sub Main()
    '
    ' Create local (client) mailslot
    '
    ClientSlot = CreateMailslotNoSecurity(ClientSlotName, 0, 0, 0)
    If ClientSlot = -1 Then
        MsgBox "Failed to open client mailslot"
    End If

    '
    ' Create local representation of server mailslot.
    ' This will be used to send messages to server's
    ' local mailslot.
    '
    ServerSlot = CreateFileNoSecurity(ServerSlotName, GENERIC_WRITE, _
        FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If ServerSlot = -1 Then
        MsgBox "Failed to open server mailslot"
    End If

    '
    ' Register return mailslot with server.
    ' This will allow the server to send info back to the client.
    '
    Call MailSlotWrite(ServerSlot, LoginCommand & ClientSlotNameForServer)

    '
    ' Expose the user interface
    '
    MClientFrm.Show
End Sub

Sub Quit(Cancel As Integer)
    '
    ' Shut down the application
    '
    Dim Result As Long
    Static Recur As Boolean

    If Recur Then Exit Sub
    Recur = True

    Result = CloseHandle(ClientSlot)
    Result = CloseHandle(ServerSlot)
    End

    Recur = False
End Sub

