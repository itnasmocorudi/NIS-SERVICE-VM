Attribute VB_Name = "Security"
Option Explicit
Public Declare Function CreateToolHelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Long
Private Declare Function SendARP Lib "iphlpapi.dll" (ByVal DestIp As Long, ByVal ScrIP As Long, pMacAddr As Long, PhyAddrLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILE_TIME, lpLastAccessTime As FILE_TIME, lpLastWriteTime As FILE_TIME) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OF_STRUCT, ByVal wStyle As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
'Declare Function WNetGetUser& Lib "Mpr" Alias "WNetGetUserA" (lpName As Any, ByVal lpUserName$, lpnLength&)
'Declare Function NetWkstaGetInfo& Lib "Netapi32" (strServer As Any, ByVal lLevel&, pbBuffer As Any)
'Declare Function NetWkstaUserGetInfo& Lib "Netapi32" (Reserved As Any, ByVal lLevel&, pbBuffer As Any)
'Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal size&)
'Declare Sub lstrcpyW Lib "kernel32" (dest As Any, ByVal src As Any)
'Declare Function NetApiBufferFree& Lib "Netapi32" (ByVal buffer&)

Public Const MAX_PATH As Integer = 260
Public Const PROCESS_VM_READ = &H10
Public Const PROCESS_QUERY_INFORMATION = &H400
Private Const OFS_MAXPATHNAME = 128
Private OriName As String, FileDesc As String

'Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const No_ERROR = 0

Public Type PROCESS_MEMORY_COUNTERS
   cb As Long
   PageFaultCount As Long
   PeakWorkingSetSize As Long
   WorkingSetSize As Long
   QuotaPeakPagedPoolUsage As Long
   QuotaPagedPoolUsage As Long
   QuotaPeakNonPagedPoolUsage As Long
   QuotaNonPagedPoolUsage As Long
   PagefileUsage As Long
   PeakPagefileUsage As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type PROCESSENTRY32
       dwSize As Long
       cntUsage As Long
       th32ProcessID As Long
       th32DefaultHeapID As Long
       th32ModuleID As Long
       cntThreads As Long
       th32ParentProcessID As Long
       pcPriClassBase As Long
       dwFlags As Long
       szExeFile As String * MAX_PATH
End Type

Private Type VS_FIXEDFILEINFO
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type

Public Type FILE_INFORMATION
    sFileDescription As String
    cFilename As String
    cDirectory As String
    cFullFilePath As String
    cFileType As String
    sOriginalFileName As String
End Type

Private Type SYSTEM_TIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type FILE_TIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type OF_STRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Type WKSTA_INFO_101
    wki101_platform_id As Long
    wki101_computername As Long
    wki101_langroup As Long
    wki101_ver_major As Long
    wki101_ver_minor As Long
    wki101_lanroot As Long
End Type

Type WKSTA_USER_INFO_1
    wkui1_username As Long
    wkui1_logon_domain As Long
    wkui1_logon_server As Long
    wkui1_oth_domains As Long
End Type

Public SQLConSecu   As String
Private DBConSecu   As ADODB.Connection
Public IpAdd                            As String       'IP Address
Public MacAdd                           As String       'Mac Address
Public CompName                         As String       'Computer Name
Public langroup                         As String       'Work Group
Public logondomain                      As String       'User Domain
Public UserName                         As String       'PC Username
Public CompGroup                        As String       'PC WorkGroup
Public CompDomain                       As String       'User Domain

Public Function Get_Connection_Secuter() As Boolean
 On Error GoTo vb_error
 Set DBConSecu = New ADODB.Connection
 DBConSecu.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & gUser & ";password=" & gPwd & ";Initial Catalog=Secuter ;server=" & gServerDb
 DBConSecu.CursorLocation = adUseClient
 DBConSecu.Open
 SQLConSecu = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & gUser & ";password=" & gPwd & ";Initial Catalog=Secuter ;server=" & gServerDb
 Get_Connection_Secuter = True
' If DBConSecu.State = adStateOpen Then DBConSecu.Close
 Exit Function
 
vb_error:
    Get_Connection_Secuter = False
    End
End Function

Public Function Check_Security()
      Dim hSnapshot As Long, lNext As Long, PID As Long, szExename As String, uProcess As PROCESSENTRY32, lOpen As Long, _
      lName As Long, sBuff As String, sExePath As String, SYSTIME As SYSTEMTIME, PROMEM As PROCESS_MEMORY_COUNTERS, _
      lProMem As Long, sUsage As String, lProTime As Long, Threads As Integer, i As Integer, ft As FILETIME, sTime As Long, _
      tExit As FILETIME, FileInformation As FILE_INFORMATION, tUser As FILETIME, tKernel As FILETIME, z() As String, Mac(5) As String, _
      Bts As Integer, sSQL As String
      
1     On Error GoTo Msg_Error
2     Call Get_McD
3     z() = Split(MacAdd, ", ")
4     On Error GoTo Errz
5     For i = 0 To 5
6       Mac(i) = z(i)
7     Next i
Errz:
8     Bts = i - 1
9     If Get_Connection_Secuter() = False Then
10       MsgBox "Koneksi Error! Silahkan Hubungi IT!"
11       End
12    Else
13       If DBConSecu.State = 0 Then DBConSecu.Open
14       If DBConSecu.Execute("Select CheckAktif from Profil")!CheckAktif = 0 Then
15          If DBConSecu.State = adStateOpen Then DBConSecu.Close
16          GoTo User_Lolos
17       End If
18       sSQL = "Select count(macaddress) as Jumlah from CompHO where MacAddress like '%" & Mac(0) & "%'"
19       For i = 1 To Bts
20          sSQL = sSQL & " Or MacAddress like '%" & Mac(i) & "%'"
21       Next i
22       If DBConSecu.Execute(sSQL)!JUMLAH > 0 Then
23          If DBConSecu.State = adStateOpen Then DBConSecu.Close
24          GoTo User_Lolos
25       End If
26       If DBConSecu.State = adStateOpen Then DBConSecu.Close
27    End If
28    hSnapshot = CreateToolHelpSnapshot(2&, 0&)
29    If hSnapshot = 0 Then Exit Function
30    uProcess.dwSize = Len(uProcess)
31    lNext = ProcessFirst(hSnapshot, uProcess)
32    Do While lNext
33        i = InStr(1, uProcess.szExeFile, Chr(0))
34        szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
35        PID = uProcess.th32ProcessID
36        Threads = uProcess.cntThreads
37        lOpen = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
38        sBuff = Space(MAX_PATH)
39        lName = GetModuleFileNameExA(lOpen, 0, sBuff, MAX_PATH)
40        sExePath = Left(sBuff, lName)
41        If sExePath = vbNullString Then sExePath = "System"
42        lProMem = GetProcessMemoryInfo(lOpen, PROMEM, Len(PROMEM))
43        Call GetFileInformation(sExePath, FileInformation, True)
44        lProTime = GetProcessTimes(lOpen, ft, tExit, tKernel, tUser)
45        FileTimeToLocalFileTime ft, ft
46        FileTimeToSystemTime ft, SYSTIME
47        If InStr(UCase(OriName), "BASESVR") > 0 And InStr(UCase(FileDesc), "MICROSOFT (R) BASESVR") > 0 Then
48              GoTo Lanjutkan
49        End If
50        lNext = ProcessNext(hSnapshot, uProcess)
51    Loop
52    GoTo Missing_Exe
Msg_Error:
53        Call CloseHandle(hSnapshot)
54        Call CloseHandle(lOpen)
55        MsgBox Erl & " - " & err.Description
56        End
Lanjutkan:
57        Call CloseHandle(hSnapshot)
58        Call CloseHandle(lOpen)
59        Exit Function
Missing_Exe:
60        Call CloseHandle(hSnapshot)
61        Call CloseHandle(lOpen)
62        MsgBox "Ada komponen yang kurang, Proses tidak dapat dilanjutkan!" & vbCrLf & "Silahkan Hubungi IT", vbCritical, "Error Notification"
63        End
User_Lolos:
End Function

Public Function GetFileInformation(ByVal fileFullPath As String, ByRef FileInformation As FILE_INFORMATION, Optional ByVal showMsgBox As Boolean = False) As Boolean
    Dim lDummy As Long, lsize As Long, rc As Long, lVerbufferLen As Long, sBuffer() As Byte, udtVerBuffer As VS_FIXEDFILEINFO, _
    lBufferLen As Long, bytebuffer(255) As Byte, Lang_Charset_String As String, HexNumber As Long, i As Integer, buffer As String, lVerPointer As Long, strVersionInfo(7) As String
    'Dim hFile As Integer, FileStruct As OF_STRUCT, CreationTime As FILE_TIME, LastAccessTime As FILE_TIME, LastWriteTime As FILE_TIME, LocalFileTime As SYSTEM_TIME, MessageString As String, strTemp As String
    On Error GoTo e_HandleFileInformationError
        With FileInformation
            lsize = GetFileVersionInfoSize(fileFullPath, lDummy)
            If lsize >= 1 Then
                ReDim sBuffer(lsize)
                rc = GetFileVersionInfo(fileFullPath, 0&, lsize, sBuffer(0))
                rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
                MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
            End If
            buffer = String(255, 0)
            lBufferLen = GetFileVersionInfoSize(fileFullPath, lDummy)
            If lBufferLen >= 1 Then
                ReDim sBuffer(lBufferLen)
                rc = GetFileVersionInfo(fileFullPath, 0&, lBufferLen, sBuffer(0))
                If rc <> 0 Then
                    rc = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
                    If rc <> 0 Then
                        MoveMemory bytebuffer(0), lVerPointer, lBufferLen
                        HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + _
                        bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
                        Lang_Charset_String = Hex(HexNumber)
                        Do While Len(Lang_Charset_String) < 8
                            Lang_Charset_String = "0" & Lang_Charset_String
                        Loop
            
                        With FileInformation
                            .sFileDescription = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "FileDescription", lVerPointer, lBufferLen, sBuffer)
                            .sOriginalFileName = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "OriginalFileName", lVerPointer, lBufferLen, sBuffer)
                        End With
                    End If
                End If
            End If
        End With
    OriName = FileInformation.sOriginalFileName
    FileDesc = FileInformation.sFileDescription
    GetFileInformation = True
    Exit Function

e_HandleFileInformationError:
    GetFileInformation = False
    Exit Function
End Function

Private Function GetStringValue(ByRef searchString As String, ByVal lVerPointer As Long, ByVal lBufferLen As Long, ByRef sBuffer() As Byte) As String
    Dim buffer As String, strTemp As String, rc As Long
    GetStringValue = ""
    buffer = String(255, 0)
    rc = VerQueryValue(sBuffer(0), searchString, lVerPointer, lBufferLen)
    If rc <> 0 Then
        lstrcpy buffer, lVerPointer
        GetStringValue = Mid$(buffer, 1, InStr(buffer, Chr(0)) - 1)
    End If
End Function

Public Function Get_McD()
Dim computer As String
Dim wmi As Variant
Dim query As Variant
Dim Mac As Variant
Dim ip As Variant
Dim mac_ids As String
Dim ip_ids As String
    
    On Error GoTo Get_McD_Error
    computer = "."
    Set wmi = GetObject("winmgmts:" & _
        "{impersonationLevel=impersonate}!\\" & _
        computer & "\root\cimv2")
    Set query = wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration where IPEnabled = true")
    
    mac_ids = ""
    ip_ids = ""
    
    For Each Mac In query
        mac_ids = mac_ids & ", " & Mac.macaddress
    Next Mac
    If Len(mac_ids) > 0 Then mac_ids = Mid$(mac_ids, 3)
    
    For Each ip In query
      ip_ids = ip_ids & ", " & ip.IPAddress(0)
    Next
    If Len(ip_ids) > 0 Then ip_ids = Mid$(ip_ids, 3)
    
    MacAdd = mac_ids
    IpAdd = ip_ids
    
    Set wmi = Nothing
    Set query = Nothing
    Exit Function
    
Get_McD_Error:
End Function

Public Function GetRemoteMACAddress(ByVal sRemoteIP As String, sRemoteMacAddress As String) As Boolean
    Dim dwRemoteIp As Long
    Dim pMacAddr As Long
    Dim bpMacAddr() As Byte
    Dim PhyAddrLen As Long
    Dim cnt As Long
    Dim tmp As String
    
    dwRemoteIp = inet_addr(sRemoteIP)
    If dwRemoteIp <> 0 Then
        PhyAddrLen = 6
        If SendARP(dwRemoteIp, 0&, pMacAddr, PhyAddrLen) = No_ERROR Then
            If pMacAddr <> 0 And PhyAddrLen <> 0 Then
                ReDim bpMacAddr(0 To PhyAddrLen - 1)
                CopyMemory bpMacAddr(0), pMacAddr, ByVal PhyAddrLen
                For cnt = 0 To PhyAddrLen - 1
                    If bpMacAddr(cnt) = 0 Then
                        tmp = tmp & "00:"
                    Else
                        tmp = tmp & Hex$(bpMacAddr(cnt)) & ":"
                    End If
                Next
                If Len(tmp) > 0 Then
                    sRemoteMacAddress = Left$(tmp, Len(tmp) - 1)
                    GetRemoteMACAddress = True
                End If
                Exit Function
            Else
                GetRemoteMACAddress = False
            End If
        Else
            GetRemoteMACAddress = False
        End If
    Else
        GetRemoteMACAddress = False
    End If
End Function
