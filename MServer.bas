Attribute VB_Name = "MServerMod"
Option Explicit

' ServerSlot is the "handle" of the server mailslot.
Global ServerSlot As Long
' ServerSlotName is the pathname of the mailslot file.
' It must begin with "\\" and be followed by the machine
' name or "." for local, and must be followed by "\mailslot\".
' After that, you can make up any valid directory tree and
' filename.  The client must know this name in order to
' start a conversation with the server.
Global Const ServerSlotName As String = "\\.\mailslot\RodneyGodfried\mserver"

' ClientSlot is the "handle" of an open client.
' It is used to talk to the client.
Global ClientSlot As Long
' ClientSlotName is the name of the client's mailslot.
Global ClientSlotName As String

' LoginCommand is a phrase that the server recognizes as a command
' for a client that is logging in (registering its return mailslot).
Global Const LoginCommand As String = "Mailslot-Login:"

Sub Main()
    '
    ' Set up server (local) mailslot.
    '
    ServerSlot = CreateMailslotNoSecurity(ServerSlotName, 0, 0, 0)
    If ServerSlot = -1 Then
        MsgBox "Failed to open mailslot"
    End If

    '
    ' Declare NO client yet.
    '
    ClientSlot = -1

    '
    ' Display user interface
    '
    MServerFrm.Show
End Sub


Sub Quit(Cancel As Integer)

    '
    ' Shutdown the application
    '
    Dim Result As Long
    Static Recur As Boolean

    '
    ' Check if client is still logged in
    '
    If ClientSlot <> -1 Then
        If WriteFileSimple(ClientSlot, "Ping", 5, Result, 0) Then
            Cancel = True
            MsgBox "Client logged in." & Chr$(13) & Chr$(10) _
                & "Can't quit now.", vbExclamation
            Exit Sub
        End If
    End If

    '
    ' Prevent recursion of this routine
    '
    If Recur Then Exit Sub
    Recur = True

    '
    ' Close mailslots.
    ' This is very important.  If left open, you will
    ' fail to create them until the OS is rebooted.
    '
    Result = CloseHandle(ServerSlot)
    Result = CloseHandle(ClientSlot)
    End

    Recur = False
End Sub

