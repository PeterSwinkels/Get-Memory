Attribute VB_Name = "GetMemoryModule"
'This module contains this program's main interface.
Option Explicit

'The Microsoft Windows API constants, functions and structures used by this program:
Private Const ERROR_IO_PENDING As Long = 997
Private Const ERROR_SUCCESS As Long = 0
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const MEM_COMMIT As Long = &H1000&
Private Const MEM_PRIVATE As Long = &H20000
Private Const PAGE_GUARD As Long = &H100&
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&
Private Const PROCESS_VM_READ As Long = &H10&
Private Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_DISABLED As Long = &H0&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2&
Private Const TOKEN_ALL_ACCESS As Long = &HFF&

Private Type LUID
   LowPart As Long
   HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type

Private Type MEMORY_BASIC_INFORMATION
   BaseAddress As Long
   AllocationBase As Long
   AllocationProtect As Long
   RegionSize As Long
   State As Long
   Protect As Long
   lType As Long
End Type

Private Type SYSTEM_INFO
   dwOemID As Long
   dwPageSize As Long
   lpMinimumApplicationAddress As Long
   lpMaximumApplicationAddress As Long
   dwActiveProcessorMask As Long
   dwNumberOrfProcessors As Long
   dwProcessorType As Long
   dwAllocationGranularity As Long
   wProcessorLevel As Integer
   wProcessorRevision As Integer
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges(1) As LUID_AND_ATTRIBUTES
End Type

Private Declare Sub GetSystemInfo Lib "Kernel32.dll" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function AdjustTokenPrivileges Lib "Advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetCurrentProcess Lib "Kernel32.dll" () As Long
Private Declare Function LookupPrivilegeValueA Lib "Advapi32.dll" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "Advapi32.dll" (ByVal ProcessH As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function ReadProcessMemory Lib "Kernel32.dll" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, lpNumberOfBytesRead As Long) As Long
Private Declare Function VirtualQueryEx Lib "Kernel32.dll" (ByVal hProcess As Long, ByVal lpAddress As Long, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long

'The constants and events used by this program:
Private Const MAX_STRING As Long = 65535   'The maximum number of characters used for a string buffer.
Private Const NO_PROCESS As Long = 0       'Indicates that no process is being viewed.


'This procedure checks whether an error has occurred during the most recent Windows API call.
Private Function CheckForError(ReturnValue As Long, Optional Ignored As Long = ERROR_SUCCESS) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long

   ErrorCode = Err.LastDllError
   Err.Clear
   
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = Ignored) Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, ByVal CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
   
      Description = Description & "API Error code: " & CStr(ErrorCode) & vbCr
      Description = Description & "Return value: " & CStr(ReturnValue)
      
      MsgBox Description, vbExclamation
   End If
   
   CheckForError = ReturnValue
End Function


'This procedure writes the memory contents of the selected process to a file.
Private Sub GetMemory(ProcessId As Long)
Dim Buffer As String
Dim BytesRead As Long
Dim FileH As Long
Dim Offset As Long
Dim ProcessH As Long
Dim MemoryBasicInformation As MEMORY_BASIC_INFORMATION
Dim ReturnValue As Long
Dim SystemInformation As SYSTEM_INFO

   If Not ProcessId = NO_PROCESS Then
      Screen.MousePointer = vbHourglass: DoEvents
      SetPrivilege SE_DEBUG_NAME, SE_PRIVILEGE_ENABLED
      
      GetSystemInfo SystemInformation
      ProcessH = CheckForError(OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION, CLng(False), ProcessId))
      If Not ProcessH = NO_PROCESS Then
         FileH = FreeFile()
         Open CStr(ProcessId) & ".dat" For Output As FileH
            Offset = SystemInformation.lpMinimumApplicationAddress
            Do While Offset <= SystemInformation.lpMaximumApplicationAddress
               ReturnValue = CheckForError(VirtualQueryEx(ProcessH, Offset, MemoryBasicInformation, Len(MemoryBasicInformation)))
               If ReturnValue = 0 Then Exit Do
               
               If Not (MemoryBasicInformation.Protect And PAGE_GUARD) = PAGE_GUARD Then
                  If MemoryBasicInformation.lType = MEM_PRIVATE Then
                     If MemoryBasicInformation.State = MEM_COMMIT Then
                        Buffer = String$(MemoryBasicInformation.RegionSize, vbNullChar)
                        ReturnValue = CheckForError(ReadProcessMemory(ProcessH, Offset, Buffer, Len(Buffer), BytesRead))
                        If Not ReturnValue = 0 Then Print #FileH, Left$(Buffer, BytesRead);
                     End If
                  End If
               End If
               Offset = MemoryBasicInformation.BaseAddress + MemoryBasicInformation.RegionSize
            Loop
         Close FileH
         CheckForError CloseHandle(ProcessH)
      End If
      
      SetPrivilege SE_DEBUG_NAME, SE_PRIVILEGE_DISABLED
      Screen.MousePointer = vbDefault
   End If
End Sub


'This procedure is executed when this program starts.
Private Sub Main()
On Error GoTo ErrorTrap
Dim Process As String
Dim ProcessId As Long

   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   Process = InputBox$("Path or process id (prefixed with ""*""):")
   If Not Process = Empty Then
      If Left$(Process, 1) = "*" Then
         ProcessId = CLng(Val(Mid$(Process, 2)))
      Else
         ProcessId = Shell(Process)
      End If
   
      If Not ProcessId = NO_PROCESS Then GetMemory ProcessId
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   MsgBox "Error: " & CStr(Err.Number) & vbCr & Err.Description, vbExclamation
   Resume EndRoutine
End Sub


'This procedure disables/enables the specified privilege for the current process.
Private Sub SetPrivilege(PrivilegeName As String, Status As Long)
Dim Length As Long
Dim NewPrivileges As TOKEN_PRIVILEGES
Dim PreviousPrivileges As TOKEN_PRIVILEGES
Dim PrivilegeId As LUID
Dim ReturnValue As Long
Dim TokenH As Long

   ReturnValue = CheckForError(OpenProcessToken(GetCurrentProcess(), TOKEN_ALL_ACCESS, TokenH))
   If Not ReturnValue = 0 Then
      
      ReturnValue = CheckForError(LookupPrivilegeValueA(vbNullString, PrivilegeName, PrivilegeId), Ignored:=ERROR_IO_PENDING)
      If Not ReturnValue = 0 Then
         NewPrivileges.Privileges(0).pLuid = PrivilegeId
         NewPrivileges.PrivilegeCount = CLng(1)
         NewPrivileges.Privileges(0).Attributes = Status
         
         CheckForError AdjustTokenPrivileges(TokenH, CLng(False), NewPrivileges, Len(NewPrivileges), PreviousPrivileges, Length)
      End If
      CheckForError CloseHandle(TokenH)
   End If
End Sub

