<div align="center">

## Mailslot


</div>

### Description

Mailslots

Mailslots have been around for a long time in Microsoft operating systems. I have heard, but not confirmed, that utilities have existed in MS-DOS to operate on Mailslots.

When you send a message via a mailslot, a single packet of data called a &#8220;datagram&#8221; is transmitted over the network. The delivery of the message is not guaranteed and it is usually not possible to know with any certainty that the message arrived at the destination. So for every important message sent over a mailslot, you must have logic to obtain a reply of some kind to indicate that the message transmitted successfully. If a mailslot message travels only inside a single machine, the Win32 library routines will indicate the successful transfer of message.

You can locate any mailslot server on the network

This code is well commented.

Have Fun.

Rodney Godfried
 
### More Info
 
The mailslot server must been started prior to start the client.

The client can be anywhere on a computer connected to a network, and will be able to find the mailslot server whitout knowing the ip adrees of the computer where the mailslot server is installed.

You may use this code as tutorial and in your program. But do not forget to metion my name. Rodney Godfried.


<span>             |<span>
---                |---
**Submitted On**   |2000-11-19 13:48:02
**By**             |[Rodney Safe Computing \(Tiger Software\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rodney-safe-computing-tiger-software.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1183011192000\.zip](https://github.com/Planet-Source-Code/rodney-safe-computing-tiger-software-mailslot__1-12934/archive/master.zip)

### API Declarations

```
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
```





