Attribute VB_Name = "modAPI"

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
Public Declare Function InternetOpen _
    Lib "wininet.dll" Alias "InternetOpenA" ( _
        ByVal sAgent As String, _
        ByVal lAccessType As Long, _
        ByVal sProxyName As String, _
        ByVal sProxyBypass As String, _
        ByVal lFlags As Long) As Long

' Opens a session for a given site.
Public Declare Function InternetConnect _
    Lib "wininet.dll" Alias "InternetConnectA" ( _
        ByVal hInternetSession As Long, _
        ByVal sServerName As String, _
        ByVal nServerPort As Integer, _
        ByVal sUsername As String, _
        ByVal sPassword As String, _
        ByVal lService As Long, _
        ByVal lFlags As Long, _
        ByVal lContext As Long) As Long

Public Declare Function InternetGetLastResponseInfo _
    Lib "wininet.dll" _
    Alias "InternetGetLastResponseInfoA" ( _
        lpdwError As Long, _
        ByVal lpszBuffer As String, _
        lpdwBufferLength As Long) As Boolean
    
Public Declare Function FtpFindFirstFile _
    Lib "wininet.dll" Alias "FtpFindFirstFileA" ( _
        ByVal hFtpSession As Long, _
        ByVal lpszSearchFile As String, _
        lpFindFileData As WIN32_FIND_DATA, _
        ByVal dwFlags As Long, _
        ByVal dwContent As Long) As Long

Public Declare Function InternetFindNextFile _
    Lib "wininet.dll" Alias "InternetFindNextFileA" ( _
        ByVal hFind As Long, _
        lpvFindData As WIN32_FIND_DATA) As Long

Public Declare Function FtpGetFile _
    Lib "wininet.dll" Alias "FtpGetFileA" ( _
        ByVal hFtpSession As Long, _
        ByVal lpszRemoteFile As String, _
        ByVal lpszNewFile As String, _
        ByVal fFailIfExists As Boolean, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Boolean

Public Declare Function FtpPutFile _
    Lib "wininet.dll" Alias "FtpPutFileA" ( _
        ByVal hFtpSession As Long, _
        ByVal lpszLocalFile As String, _
        ByVal lpszRemoteFile As String, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Boolean

'Closes a single Internet handle
' or a subtree of Internet handles.
Public Declare Function InternetCloseHandle _
    Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Function SetFTPflag() As Boolean
    SetFTPflag = True
End Function

