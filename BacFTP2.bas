Attribute VB_Name = "modPub"

Option Explicit

Public Const BSLASH = "\"
Public Const FSLASH = "/"
Public Const LYELLOW = &HCDFAFF
Public Const SPCR = 120, UNIT = 300

Public Const MAX_PATH = 260
Public Const ERROR_NO_MORE_FILES = 18
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public Type SYSTEMTIME
  wYear             As Integer
  wMonth            As Integer
  wDayOfWeek        As Integer
  wDay              As Integer
  wHour             As Integer
  wMinute           As Integer
  wSecond           As Integer
  wMilliseconds     As Long
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

' strFTP()
'0 = Server Address (ftp.myserver.com)
'1 = User Name
'2 = Password
'3 = Index File Name
Public strFTP$(3)

'INI Files
Public inifile1$    'Command strings
Public inifile2$    'Transfer types - extensions
Public inifile3$    'Transfer types - filenames
Public txtfile1$    'Commands help file
Public txtfile3$    'Set FTP server help file
Public txtfile2$    'Transfer type help file

'Transfer types
Public transExtens As clsFileInfo
Public transFNames As clsFileInfo

'Keep track of Form positions
Public frmCmdsLeft!, frmCmdsTop!

'Misc flag
Public doNotExecute As Boolean

Public Declare Function FileTimeToLocalFileTime _
    Lib "kernel32" (lpFileTime As FILETIME, _
        lpLocalFileTime As FILETIME) As Long

Public Declare Function FileTimeToSystemTime _
    Lib "kernel32" ( _
        lpFileTime As FILETIME, _
        lpSystemTime As SYSTEMTIME) As Long

Public Declare Function FindFirstFile _
    Lib "kernel32" Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindNextFile _
    Lib "kernel32" Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long
        
Public Declare Function GetTempPath _
    Lib "kernel32" Alias "GetTempPathA" ( _
        ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
        
Public Function WinTmpDir$()
   Dim n&, s$
   
   s = Space$(MAX_PATH)
   n = GetTempPath(MAX_PATH, s)
   WinTmpDir = Left$(s, n)
End Function
        
Public Function GetFileDate(ByVal fpn$) As Date
    Dim hFile&, d As Date
    Dim WFD As WIN32_FIND_DATA
    
    hFile = FindFirstFile(fpn, WFD)
    If hFile > 0 Then
        d = ConvertFileDate(WFD.ftLastWriteTime)
        FindClose hFile
    Else
        d = 0
    End If
    
    GetFileDate = d
End Function
        
Public Function ConvertFileDate(f As FILETIME) As Date
    Dim ST As SYSTEMTIME
    Dim g As FILETIME

    FileTimeToLocalFileTime f, g
    FileTimeToSystemTime g, ST
   
    ConvertFileDate = _
        DateSerial(ST.wYear, ST.wMonth, ST.wDay) + _
        TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
End Function

Public Function TrimNull$(s$)
    Dim pos%
    pos = InStr(s, Chr$(0))
    
    Select Case pos
        Case 0: TrimNull = s
        Case 1: TrimNull = ""
        Case Else: TrimNull = Left$(s, pos - 1)
    End Select
End Function

Public Function fixSlash$(ByVal s$, s1$)
    Dim j%, s2$
    
    s2 = IIf(s1 = FSLASH, BSLASH, FSLASH)
    
    For j = 1 To Len(s)
        If Mid$(s, j, 1) = s1 Then Mid$(s, j, 1) = s2
    Next
    
    fixSlash = s
End Function
