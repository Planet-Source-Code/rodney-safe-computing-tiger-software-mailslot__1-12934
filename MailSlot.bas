Attribute VB_Name = "MailslotMod"
Option Explicit
'=====================
'Rodney Safe Computing
'Created by Rodney Godfried
'18-08-1999
'=====================

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function CreateMailslotNoSecurity Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, ByVal Zero As Long) As Long
Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Declare Function SetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, ByVal lReadTimeout As Long) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function ReadFileSimple Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal Zero As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function WriteFileSimple Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal Zero As Long) As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CreateFileNoSecurity Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal Zero As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const MAILSLOT_NO_MESSAGE = (-1)
Public Const MAILSLOT_WAIT_FOREVER = (-1)
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const OPEN_EXISTING = 3
Function MailSlotRead(SlotHandle As Long) As String
    '
    ' Read a specified mailslot.
    ' Return the text of the next message (if any).
    '
    Dim Result As Boolean ' result of system calls
    Dim MessageNew As String ' text of new message
    Dim MessageCount As Long ' count of messages waiting in slot
    Dim MessageLength As Long ' len of next messgae in queue
    Dim ReadTimeout As Long ' how long to wait for message
    Dim BytesRead As Long ' actual count of bytes read from slot

    '
    ' Check if any messages are waiting in mailslot
    '
    Result = GetMailslotInfo(SlotHandle, 0, MessageLength, MessageCount, 0)
    If Result = 0 Then
        MsgBox "Bad return from GetMailslotInfo"
        Exit Function
    End If
    If MessageLength = MAILSLOT_NO_MESSAGE Then
        ' no message in mailslot queue
        MailSlotRead = ""
        Exit Function
    End If

    '
    ' Retrieve next message from mailslot
    '
    MessageNew = String$(MessageLength + 1, " ")
    Result = ReadFileSimple(SlotHandle, MessageNew, MessageLength, _
        BytesRead, 0)
    If Result = 0 Then
        'MsgBox "failed to read message"
        MailSlotRead = ""
        Exit Function
    End If
    If BytesRead <> MessageLength Then
        'MsgBox "Did not read correct message length"
        MailSlotRead = ""
        Exit Function
    End If

    '
    ' Return message
    '
    MailSlotRead = MessageNew
End Function
Sub MailSlotWrite(SlotHandle As Long, Text As String)
    '
    ' Write message to server's mailslot.
    '
    Dim Result As Boolean ' result of system calls
    Dim TextLen As Long ' length of message to send
    Dim BytesWritten As Long ' actual count of bytes sent

    '
    ' Send message over mailslot
    '
    TextLen = Len(Text) + 1
    Result = WriteFileSimple(SlotHandle, Text, TextLen, BytesWritten, 0)
    If Result = 0 Then
        MsgBox "failed to write message"
        Exit Sub
    End If
    If BytesWritten <> TextLen Then
        MsgBox "wrote wrong number of bytes"
        Exit Sub
    End If
End Sub
