Attribute VB_Name = "modFileCopy"
'========================================================
'                  FTP BACKUP UTILITY
'                   in Visual Basic©
'           by: Rick Meyer    Date: August 2001
'========================================================
' File Name:   BacFTP3.bas
' Object Name: modFileCopy
' Description: Module declaring dummy API's
'               reworked as file copiers
'========================================================
Option Explicit

Public Const scUserAgent = "vb wininet"
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

Public Const INTERNET_INVALID_PORT_NUMBER = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1

Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2

Public Const INTERNET_FLAG_PASSIVE = &H8000000
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const ERROR_INTERNET_EXTENDED_ERROR = 12003

' Initialize use of the Win32 Internet functions
Public Function InternetOpen( _
        ByVal sAgent As String, _
        ByVal lAccessType As Long, _
        ByVal sProxyName As String, _
        ByVal sProxyBypass As String, _
        ByVal lFlags As Long) As Long
    
    InternetOpen = 1
End Function

' Opens a session for a given site.
Public Function InternetConnect( _
        ByVal hInternetSession As Long, _
        ByVal sServerName As String, _
        ByVal nServerPort As Integer, _
        ByVal sUsername As String, _
        ByVal sPassword As String, _
        ByVal lService As Long, _
        ByVal lFlags As Long, _
        ByVal lContext As Long) As Long
    
    InternetConnect = 1
End Function

Public Function InternetGetLastResponseInfo( _
        lpdwError As Long, _
        ByVal lpszBuffer As String, _
        lpdwBufferLength As Long) As Boolean
    
    InternetGetLastResponseInfo = False
End Function
    
Public Function FtpFindFirstFile( _
        ByVal hFtpSession As Long, _
        ByVal lpszSearchFile As String, _
        lpFindFileData As WIN32_FIND_DATA, _
        ByVal dwFlags As Long, _
        ByVal dwContent As Long) As Long
    
    Dim lng&
    lng = FindFirstFile( _
        fixSlash(lpszSearchFile, FSLASH), lpFindFileData)
    
    If lng < 0 Then lng = 0
    FtpFindFirstFile = lng
End Function

Public Function InternetFindNextFile( _
        ByVal hFind As Long, _
        lpvFindData As WIN32_FIND_DATA) As Long
    
    InternetFindNextFile = FindNextFile(hFind, _
                            lpvFindData)
End Function
        
Public Function FtpGetFile( _
        ByVal hFtpSession As Long, _
        ByVal lpszRemoteFile As String, _
        ByVal lpszNewFile As String, _
        ByVal fFailIfExists As Boolean, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Boolean

    fileMove fixSlash(lpszRemoteFile, FSLASH), lpszNewFile, dwFlags
    FtpGetFile = True
End Function

Public Function FtpPutFile( _
        ByVal hFtpSession As Long, _
        ByVal lpszLocalFile As String, _
        ByVal lpszRemoteFile As String, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Boolean
        
    fileMove lpszLocalFile, fixSlash(lpszRemoteFile, FSLASH), dwFlags
    FtpPutFile = True
End Function

Private Sub fileMove(src$, dst$, flg&)
    Dim f%, g%, t$, b As Byte
    
    f = FreeFile
    
    If flg = FTP_TRANSFER_TYPE_ASCII Then
        Open src For Input As f
        g = FreeFile
        Open dst For Output As g
        
        Do While Not EOF(f)
            t = Input(1, f)
            Print #g, t;
        Loop
    Else
        Open src For Binary As f
        g = FreeFile
        Open dst For Binary As g
        
        Do While Not EOF(f)
            Get f, , b
            Put g, , b
        Loop
    End If
    
    Close f, g
End Sub

Public Function InternetCloseHandle%(ByVal hInet&)
    InternetCloseHandle = 0
End Function

Public Function SetFTPflag() As Boolean
    SetFTPflag = False
End Function

