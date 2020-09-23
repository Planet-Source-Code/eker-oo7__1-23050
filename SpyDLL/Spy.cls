VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Spy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************** DECLARES REGISTER EVENT HANDLER / SUBCLASS *****************
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
   (ByVal pv As Long)

Private Declare Function GetProp Lib "user32" _
    Alias "GetPropA" _
   (ByVal hwnd As Long, _
    ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32" _
    Alias "SetPropA" _
    (ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal hData As Long) As Long
    
Private Declare Function RemoveProp Lib "user32" _
    Alias "RemovePropA" _
    (ByVal hwnd As Long, _
    ByVal lpString As String) As Long
    
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
        
Private Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" _
   (ByVal hwnd As Long, _
    ByVal uFlags As SHCN_ItemFlags, _
    ByVal dwEventID As Long, _
    ByVal uMsg As Long, _
    ByVal cItems As Long, _
    lpps As PIDLSTRUCT) As Long

Private Declare Function SHChangeNotifyDeregister Lib "shell32" _
    Alias "#4" _
   (ByVal hNotify As Long) As Boolean
   
   
   
Private Enum SHCN_ItemFlags
   SHCNF_IDLIST = &H0         'LPITEMIDLIST
   SHCNF_PATHA = &H1          'path name
   SHCNF_PRINTERA = &H2       'printer friendly name
   SHCNF_DWORD = &H3          'DWORD
   SHCNF_PATHW = &H5          'path name
   SHCNF_PRINTERW = &H6       'printer friendly name
   SHCNF_TYPE = &HFF
   SHCNF_FLUSH = &H1000
   SHCNF_FLUSHNOWAIT = &H2000
   #If UNICODE Then
       SHCNF_PATH = SHCNF_PATHW
       SHCNF_PRINTER = SHCNF_PRINTERW
   #Else
        SHCNF_PATH = SHCNF_PATHA
        SHCNF_PRINTER = SHCNF_PRINTERA
   #End If
End Enum

Private Type PIDLSTRUCT
   pidl As Long
   bWatchSubFolders As Long
End Type
   
'*********************** DECLARES OS VERSION INFO ********************
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (LpVersionInformation As OSVERSIONINFO) As Long



Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string For PSS usage
End Type

Private Enum VERSION
    WIN95 = &H0
    WIN98 = &H1
    NT3 = &H3
    NT4 = &H4
    NT2000 = &H5
End Enum


'*********************** DECLARES WRITE TO TEXT FILE ********************
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long


Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

'*************************** DECLARES NT LOG ****************************
Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias _
        "RegisterEventSourceA" (ByVal lpUNCServerName As String, _
        ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "advapi32.dll" ( _
        ByVal hEventLog As Long) As Long
Private Declare Function ReportEvent Lib "advapi32.dll" Alias _
      "ReportEventA" ( _
        ByVal hEventLog As Long, ByVal wType As Integer, _
        ByVal wCategory As Integer, ByVal dwEventID As Long, _
        ByVal lpUserSid As Any, ByVal wNumStrings As Integer, _
        ByVal dwDataSize As Long, plpStrings As Long, _
        lpRawData As Any) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        hpvDest As Any, hpvSource As Any, _
        ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" ( _
         ByVal wFlags As Long, _
         ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" ( _
         ByVal hMem As Long) As Long
            
      Private Const EVENTLOG_SUCCESS = 0
      Private Const EVENTLOG_ERROR_TYPE = 1
      Private Const EVENTLOG_WARNING_TYPE = 2
      Private Const EVENTLOG_INFORMATION_TYPE = 4
      Private Const EVENTLOG_AUDIT_SUCCESS = 8
      Private Const EVENTLOG_AUDIT_FAILURE = 10
      Private Const CUSTOM_ERROR = 50001
'******************** GET ERRORS *********************
Private Declare Function GetLastError Lib "kernel32" () As Long

'******************** CONSTANTS **********************
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const OPEN_ALWAYS = 4
Private Const CREATE_ALWAYS = 2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Const FILE_END = 2
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_SHNOTIFY = &H401
Private Const GWL_WNDPROC As Long = (-4)
Private Const OLDWNDPROC As String = "OldWndProc"
Private lpPrevWndProc As Long
Private OS As Long
Private LAST_ERROR As String
Private lbWatchSubFolders As Boolean

'************************* Make Constanst Available ***************
Public Property Get SHNOTIFY() As Long
    SHNOTIFY = WM_SHNOTIFY
End Property

Public Property Get NCDESTROY() As Long
    NCDESTROY = WM_NCDESTROY
End Property
Public Property Get dllLastError() As String
    dllLastError = LAST_ERROR
End Property
Public Property Let WatchSubFolders(ByVal Value As Boolean)
    lbWatchSubFolders = Value
End Property
'********************** Exposed Functions *****************************************
Public Function EmploySpy(ByVal hwnd As Long, ByVal EventType As Long, ByVal FolderPath As String, ByVal AddressOfFunction As Long, ByRef pidl As Long, ByRef hSHNotify As Long) As Boolean
On Error GoTo Err_Employ
   If Hook(hwnd, AddressOfFunction) Then
           If SHNotify_Register(hwnd, FolderPath, EventType, pidl, hSHNotify) Then
                EmploySpy = True
           Else
               If GetLastError <> 0 Then
                    Err.Raise GetLastError
               Else
                     Err.Raise CUSTOM_ERROR, "EmploySpy", "EmploySpy Failed For Unknown reasons"
               End If
           End If
      
   Else
        If GetLastError <> 0 Then
               Err.Raise GetLastError
        Else
             Err.Raise CUSTOM_ERROR, "EmploySpy", "EmploySpy Failed For Unknown reasons"
        End If
   End If
Exit Function
Err_Employ:
     EmploySpy = False
     LogEvent "Error in [EmploySpy]:#" & Err.Number & "-" & Err.Description, EVENTLOG_ERROR_TYPE, Err.Number
     Err.Clear
End Function

Public Function DestroySpy(ByRef hSHNotify As Long, ByRef pidl As Long, ByVal hwnd As Long) As Boolean
On Error GoTo Err_Destroy
    If SHNotify_Unregister(hSHNotify, pidl) Then
        If UNHook(hwnd) Then
             DestroySpy = True
        Else
            If GetLastError <> 0 Then
                Err.Raise GetLastError
            Else
                Err.Raise CUSTOM_ERROR, "DestroySpy", "EmploySpy Failed For Unknown reasons"
            End If
        End If
    Else
        If GetLastError <> 0 Then
            Err.Raise GetLastError
        Else
            Err.Raise CUSTOM_ERROR, "DestroySpy", "EmploySpy Failed For Unknown reasons"
        End If
    End If

Exit Function
Err_Destroy:
     DestroySpy = False
     LogEvent "Error in [DestroySpy]:#" & Err.Number & "-" & Err.Description, EVENTLOG_ERROR_TYPE, Err.Number
     Err.Clear
End Function
 
 Public Function NotifyWindows(hwnd, uMsg, wParam, lParam) As Long
 On Error GoTo Err_Notify
        NotifyWindows = CallWindowProc(GetProp(hwnd, OLDWNDPROC), hwnd, uMsg, wParam, lParam)
        NotifyWindows = NotifyWindows
Exit Function
Err_Notify:
     NotifyWindows = 0
     LogEvent "Error in [NotifyWindows]:#" & Err.Number & "-" & Err.Description, EVENTLOG_ERROR_TYPE, Err.Number
     Err.Clear
End Function
'******************** General Usage *************************
Private Sub LogEvent(ByVal Msg As String, ByVal LOG_TYPE As Integer, ByVal ERROR_NUMBER As Long)
    If OS <= VERSION.WIN98 Then
        LogTextEvent Msg
    Else
        LogNTEvent Msg, LOG_TYPE, ERROR_NUMBER
    End If
    LAST_ERROR = ERROR_NUMBER & "-" & Msg
End Sub

 Private Function LogNTEvent(ByVal sString As String, ByVal iLogType As Integer, _
        ByVal iEventID As Long)
        Dim bRC As Boolean
        Dim iNumStrings As Integer
        Dim hEventLog As Long
        Dim hMsgs As Long
        Dim cbStringSize As Long
        hEventLog = RegisterEventSource("", App.Title)
        cbStringSize = Len(sString) + 1
        hMsgs = GlobalAlloc(&H40, cbStringSize)
        CopyMemory ByVal hMsgs, ByVal sString, cbStringSize
        iNumStrings = 1
        If ReportEvent(hEventLog, _
           iLogType, 0, _
           iEventID, 0&, _
           iNumStrings, cbStringSize, _
           hMsgs, hMsgs) = 0 Then
           'GetLastError()
        End If
        Call GlobalFree(hMsgs)
        DeregisterEventSource (hEventLog)
 End Function

Private Function LogTextEvent(ByVal sString As String)
Dim filehandle As Long
Dim filepath As String * 255
Dim filename As String
Dim writesuccess As Long
Dim written As Long
Dim security As SECURITY_ATTRIBUTES
Dim length As Long

Dim Msg() As Byte
Msg() = Now() & sString & "|"
With security
    .nLength = Len(security)
    .lpSecurityDescriptor = 0
    .bInheritHandle = 0
End With

filename = App.Title & ".log"
filepath = App.Path & "\" & filename
filehandle = CreateFile(filepath, GENERIC_WRITE, 0, security, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
SetFilePointer filehandle, 0, 0, FILE_END
If filehandle <> INVALID_HANDLE_VALUE Then
  writesuccess = WriteFile(filehandle, Msg(0), UBound(Msg), written, 0)
  FlushFileBuffers (filehandle)
Else
  'Write Failed
End If
CloseHandle (filehandle)
End Function
'******************** SUBCLASS(Hook)/UNSUBCLASS(UnHook)*********************
Private Function Hook(hwnd As Long, ByVal AddressOfFunction As Long) As Boolean
    If GetProp(hwnd, OLDWNDPROC) = 0 Then
         lpPrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOfFunction)
         If lpPrevWndProc Then
            Hook = CBool(SetProp(hwnd, OLDWNDPROC, lpPrevWndProc))
         Else
            Hook = False
            Call UNHook(hwnd)
         End If
      End If
End Function

Private Function UNHook(hwnd As Long) As Boolean
   Dim lpTemp As Long
   lpTemp = GetProp(hwnd, OLDWNDPROC)
         If lpTemp Then
            If RemoveProp(hwnd, OLDWNDPROC) Then
                UNHook = CBool(SetWindowLong(hwnd, GWL_WNDPROC, lpTemp))
            End If
        End If
End Function

'********************* Register/UnRegister Event Handeler******************
Private Function SHNotify_Register(hwnd As Long, ByVal FolderPath As String, ByVal EventType As Long, ByRef pidl As Long, ByRef hSHNotify As Long) As Boolean
   Dim ps As PIDLSTRUCT
   If (hSHNotify = 0) Then
         If OS <= VERSION.WIN98 Then
            pidl = SHSimpleIDListFromPath(FolderPath)
         Else
            pidl = SHSimpleIDListFromPath(StrConv(FolderPath, vbUnicode))
         End If
         ps.pidl = pidl
         If lbWatchSubFolders Then
            ps.bWatchSubFolders = True
         Else
            ps.bWatchSubFolders = False
         End If
    
         hSHNotify = SHChangeNotifyRegister(hwnd, _
                                              SHCNF_TYPE Or SHCNF_IDLIST, _
                                              EventType, _
                                              WM_SHNOTIFY, _
                                              1, _
                                              ps)
        SHNotify_Register = CBool(hSHNotify)
  End If
End Function

Private Function SHNotify_Unregister(ByRef hSHNotify As Long, ByRef pidl As Long) As Boolean
    If hSHNotify Then
       If SHChangeNotifyDeregister(hSHNotify) Then
         hSHNotify = 0
         Call CoTaskMemFree(pidl)
         pidl = 0
         SHNotify_Unregister = True
      End If
    End If
End Function

Private Sub Class_Initialize()
Dim OSINFO As OSVERSIONINFO
OSINFO.dwOSVersionInfoSize = 148
OSINFO.szCSDVersion = Space$(128)
GetVersionEx OSINFO
          
          With OSINFO
            Select Case .dwPlatformId
               Case 1
                  If .dwMinorVersion = 0 Then
                     OS = VERSION.WIN95
                  ElseIf .dwMinorVersion = 10 Then
                     OS = VERSION.WIN98
                  End If
               Case 2
                  If .dwMajorVersion = 3 Then
                     OS = VERSION.NT3
                  ElseIf .dwMajorVersion = 4 Then
                     OS = VERSION.NT4
                  ElseIf .dwMajorVersion = 5 Then
                     OS = VERSION.NT2000
                  End If
               Case Else
                 'Error
            End Select
          End With
LogEvent "Object Created Succesfully", EVENTLOG_INFORMATION_TYPE, 0
Debug.Print App.EXEName
End Sub
