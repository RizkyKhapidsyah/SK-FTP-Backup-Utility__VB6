VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "FTP Backup"
   ClientHeight    =   2655
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   4815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "BacFTP.frx":0000
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Menu mnuChgOp 
      Caption         =   "&Backup"
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
   End
   Begin VB.Menu mnuEditComms 
      Caption         =   "&Commands"
   End
   Begin VB.Menu mnuEditTypes 
      Caption         =   "&Types"
   End
   Begin VB.Menu mnuAutoMode 
      Caption         =   "&Automatic"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum opType
    UPLD
    DNLD
End Enum

Private Enum autoType
    MANUAL
    FULL
End Enum

Dim MODE As opType
Dim AUTO As autoType

'FileInfo defined in Class module
Dim localFiles As clsFileInfo
Dim localIndex As clsFileInfo
Dim remoteFiles As clsFileInfo
Dim remoteIndex As clsFileInfo

Dim localDir$, remoteDir$
Dim hOpen&, hConnection&

Dim done As Boolean
Dim errorcondition As Boolean
Dim saveLocalIndexFlag As Boolean

Dim FTPflag As Boolean
Dim helpFlag As Boolean
Dim HelpLoaded As Boolean

'INI file (other INI files defined in modPub)
Dim inifile0$
Dim txtfile0$

Private Sub generalFTPopen()
    If frmCmds.List1.ListCount = 0 Then
        MsgBox "No commands entered"
        Exit Sub
    End If
    
    DoEvents
    hOpen = InternetOpen(scUserAgent, _
        INTERNET_OPEN_TYPE_DIRECT, vbNullString, _
        vbNullString, 0)
    DoEvents
        
    If hOpen = 0 Then
        errMsg "InternetOpen"
        Exit Sub
    End If

    openhConnection
    DoEvents
    If hConnection = 0 Then Exit Sub
    
uc1: cycleCommands
    If Not done And AUTO = FULL Then GoTo uc1
    
    Me.SetFocus
End Sub

Private Sub cycleCommands()
    Static cntr%
    Dim j%, k%, l%, tmp$, tmp1$

ex1: cntr = cntr + 1
    If done Then cntr = frmCmds.List1.ListCount + 2

    If cntr > frmCmds.List1.ListCount Then
        IClose hConnection
        IClose hOpen
        Label1.Caption = "All Done!"
        done = True
        cntr = 0
        Command1.Caption = "Start"
        Command2.Caption = "Exit"
        Command2.SetFocus
        Exit Sub
    End If

    j = cntr - 1
    doNotExecute = True
    frmCmds.List1.Selected(j) = True
    doNotExecute = False
    tmp = frmCmds.List1.List(j)
    k = InStr(tmp, " to ")
    l = Len(tmp)
    
    If k > 1 Then
        tmp1 = Left$(tmp, k - 1)
        localDir = fileDir(tmp1)
        localDir = addSlash(localDir, BSLASH)

        k = k + 3
        If k < l Then
            remoteDir = Right$(tmp, l - k)
            remoteDir = addSlash(remoteDir, FSLASH)
            
            List1.Clear
            Label1.Caption = ""
            errorcondition = False
            If MODE = UPLD Then
                executeUploadCmd tmp1
            Else
                executeDnloadCmd tmp1
            End If
            
            If AUTO = MANUAL Then
                If errorcondition Then GoTo ex1
                Command1.SetFocus
            Else
                GoTo ex1
            End If
        End If
    End If
End Sub

'========================================================
'            Upload Specific Routines
'========================================================
'Determine which local files need to be backed up.
'  The localFileSpec can have normal file search
'  wildcards and all matching files are checked.
Private Sub executeUploadCmd(localFileSpec$)
    'Must do this first since we need to verify
    '  with all files.
    If getLocalIndex Then
        getLocalFiles "*.*"
        verifyLocalIndex
    End If
    
    'Now just get the local matching files
    getLocalFiles FileName(localFileSpec)
    
    If localFiles.Count = 0 Then
        InformNothingToDo "None matching " & _
                        localFileSpec
        Exit Sub
    End If
    
    Dim j%, k%, tmp$
    'Now fill the List Box with the files
    For j = 0 To localFiles.Count - 1
        List1.AddItem localFiles.FileName(j)
    Next

    'Now look in the remote directory for the file
    '  index. This index contains last modification
    '  dates of the files previously backed up. From
    '  these dates determine if the requested
    '  localFileSpec(s) actually need to be backed up.
    'All matching localFileSpec(s) will be placed in
    '  the listbox but only the ones that need to be
    '  backed up are selected for upload.
    If getRemoteIndex = False Then
        'No index was found
        selectAllFiles
        GoTo xup1
    End If
    
    'We have the remoteIndex class which lists the
    '  actual remote files. Perhaps the actual remote
    '  file has been deleted somehow. So now we get
    '  all the files in the remote directory and check
    '  them against the index.
    getRemoteFiles "*.*"
    
    j = 0
    Do Until j >= remoteIndex.Count
        tmp = remoteIndex.FileName(j)
        If remoteFiles.Exists(tmp, k) Then
            j = j + 1
        Else
            'File not found so remove
            '  from the index.
            remoteIndex.RemoveItem tmp
        End If
    Loop

    'Now loop through the listbox of local files and
    '  check if a remote file of the same name is in
    '  the index. If not then select it in the listbox
    '  and add it to the remoteIndex. If it is then check
    '  if the remote file is out of date. If it is then
    '  select it and update the remoteIndex.newDate.
    Dim ptr1%, ptr2%, ptr3%, rptr%
    saveLocalIndexFlag = False
    
    For j = 0 To List1.ListCount - 1
        tmp = List1.List(j)
        localFiles.Exists tmp, ptr1
        'See if it exists in the remote index
        If remoteIndex.Exists(tmp, rptr) Then
            'See if it exists in the local index
            If localIndex.Exists(tmp, ptr2) Then
                If localFiles.OldDate(ptr1) > _
                  localIndex.OldDate(ptr2) Then
                    List1.Selected(j) = True
                    'Update the index for upload
                    remoteIndex.NewDate(rptr) = _
                        localFiles.OldDate(ptr1)
                    localIndex.OldDate(ptr2) = _
                        localFiles.OldDate(ptr1)
                    localIndex.NewDate(ptr2) = _
                        localFiles.OldDate(ptr1)
                    saveLocalIndexFlag = True
                End If
            'No entry in the local directory
            ElseIf localFiles.OldDate(ptr1) > _
              remoteIndex.OldDate(rptr) Then
                List1.Selected(j) = True
                'Update the index for upload
                remoteIndex.NewDate(rptr) = _
                    localFiles.OldDate(ptr1)
            End If
        Else
            List1.Selected(j) = True
            
            'Add a new index for upload
            remoteIndex.AddItem tmp, , _
                    localFiles.OldDate(ptr1)
        End If
    Next
    
    'Now check if no items are selected
xup1:
    Dim fileCntr%
    For j = 0 To List1.ListCount - 1
        If List1.Selected(j) Then
            fileCntr = fileCntr + 1
        End If
    Next
    
    If fileCntr = 0 Then
        InformNothingToDo
    Else
        tmp = " file"
        If fileCntr > 1 Then tmp = tmp & "s"
        Label1.Caption = Trim$(Str$(fileCntr)) _
                & tmp & " to backup"
    End If
End Sub

Private Sub continueUpload()
    Dim j%, k%, s$, t$, b As Boolean
    
    Dim mt As New clsMyTime
    
    For j = 0 To List1.ListCount - 1
        If List1.Selected(j) Then
            s = List1.List(j)
            
            DoEvents
            b = FtpPutFile(hConnection, localDir & s, _
                remoteDir & s, setTransferType(s), 0)
            DoEvents
            
            If b Then
                t = " *upload OK"
            Else
                t = " *FAILED " & lastErr
                
                'The upload of this file has failed
                '  remove it from the remote index
                remoteIndex.Exists List1.List(j), k
                remoteIndex.RemoveItem k
            End If
            
            List1.List(j) = List1.List(j) & t
            List1.Selected(j) = False
            Label1.Caption = mt.getMyTime
            DoEvents
        End If
    Next
        
    uploadIndex
    If saveLocalIndexFlag Then saveLocalIndex
    Label1.Caption = mt.getMyTime
    Set mt = Nothing
End Sub

Private Sub uploadIndex()
    Dim f%, j%, tmpFile$, b As Boolean
    
    f = FreeFile
    tmpFile = WinTmpDir & strFTP(3)
    
    'First write the remoteIndex to a temp file
    Open tmpFile For Output As f
    For j = 0 To remoteIndex.Count - 1
        Write #f, remoteIndex.FileName(j), _
                  remoteIndex.NewDate(j)
    Next
    Close f
    
    'Now upload the index
    DoEvents
    b = FtpPutFile(hConnection, tmpFile, _
        remoteDir & strFTP(3), _
        FTP_TRANSFER_TYPE_ASCII, 0)
    DoEvents
    
    If b = False Then errMsg "FtpPutFile " & tmpFile
    If Dir$(tmpFile) <> "" Then Kill tmpFile
End Sub


'========================================================
'          Download Specific Routines
'========================================================
Private Sub executeDnloadCmd(localFileSpec$)
    'First get the remote index
    If getRemoteIndex = False Then
        InformNothingToDo "No Remote Index"
        Exit Sub
    End If
    
    'Now get the Remote files that match the filespec
    getRemoteFiles FileName(localFileSpec)
    
    'Make sure the remote files are in the remote index
    '  If not then remove it from the download list
    Dim j%, k%, tmp$
    Do Until j >= remoteFiles.Count
        tmp = remoteFiles.FileName(j)
        If remoteIndex.Exists(tmp, k) = False Then
            remoteFiles.RemoveItem tmp
        Else
            remoteFiles.OldDate(j) = remoteIndex.OldDate(k)
            j = j + 1
        End If
    Loop
    
    'Now make sure there are some files to get
    If remoteFiles.Count = 0 Then
        InformNothingToDo _
            "No Matching Files in Remote Index"
        Exit Sub
    End If
    
    'Now fill the List Box with the files
    For j = 0 To remoteFiles.Count - 1
        List1.AddItem remoteFiles.FileName(j)
    Next
    
    'Get all files in the local directory
    getLocalFiles "*.*"

    'Now look in the local directory for the file
    '  index. This index contains last modification
    '  dates and last download dates of the
    '  files previously restored. From these dates
    '  determine if the requested remote files actually
    '  need to be restored.
    'All matching remoteFileSpec(s) have been placed in
    '  the listbox but only the ones that need to be
    '  restored are selected for download.
    'Get the local index
    getLocalIndex
    
    'The local index file will be rewritten so take this
    '  opportunity to verify that all files in the index
    '  are actually in the directory. If not then remove
    '  from index.
    verifyLocalIndex
    
    'In the localIndex
    'Filename, Download (filedate), remoteIndexDate
    
    'Now loop through the listbox of remote files and
    '  check if a file of the same name is in the local
    '  index. If not then select it in the listbox and
    '  add it to the remoteIndex. If it is then check if
    '  the local file is out of date.
    '  If it is then select it and update the
    '  localIndex.newDate.
    Dim lFptr%, lIptr%, rIptr%
    For j = 0 To List1.ListCount - 1
        tmp = List1.List(j)
        localFiles.Exists tmp, lFptr   'set pointers
        remoteIndex.Exists tmp, rIptr
        
        If localIndex.Exists(tmp, lIptr) Then
            'Outdated with repect to remote index
            If remoteIndex.OldDate(rIptr) > _
              localIndex.NewDate(lIptr) Then
                'Make sure the actual local file is not
                '  the most current
                If localFiles.OldDate(lFptr) <= _
                  localIndex.OldDate(lIptr) Then
                    List1.Selected(j) = True
                    'Update the local index for saving
                    localIndex.NewDate(lIptr) = _
                        remoteIndex.OldDate(rIptr)
                End If
            End If
        Else
            If localFiles.OldDate(lFptr) < _
              remoteIndex.OldDate(rIptr) Then
                List1.Selected(j) = True
            
                'Add a new file in the local index
                localIndex.AddItem tmp, , _
                    remoteIndex.OldDate(rIptr)
            End If
        End If
    Next
    
    'Now check if no items are selected
    Dim fileCntr%
    For j = 0 To List1.ListCount - 1
        If List1.Selected(j) Then
            fileCntr = fileCntr + 1
        End If
    Next
        
    If fileCntr = 0 Then
        InformNothingToDo
    Else
        tmp = " file"
        If fileCntr > 1 Then tmp = tmp & "s"
        Label1.Caption = Trim$(Str$(fileCntr)) _
                & tmp & " to restore"
    End If
End Sub

Private Sub continueDownload()
    Dim j%, ptr%, s$, t$, b As Boolean
    
    Dim mt As New clsMyTime
    
    For j = 0 To List1.ListCount - 1
        If List1.Selected(j) Then
            s = List1.List(j)
            localIndex.Exists s, ptr
            
            DoEvents
            b = FtpGetFile(hConnection, remoteDir & s, _
                localDir & s, False, INTERNET_FLAG_RELOAD, _
                setTransferType(s), 0)
            DoEvents
            
            If b Then
                t = " *download OK"
                localIndex.OldDate(ptr) = _
                    GetFileDate(localDir & s)
            Else
                t = " *FAILED " & lastErr
                
                'The Download of this file has failed
                '  remove it from the local index
                localIndex.RemoveItem ptr
            End If
            
            List1.List(j) = List1.List(j) & t
            List1.Selected(j) = False
            Label1.Caption = mt.getMyTime
            DoEvents
        End If
    Next
        
    saveLocalIndex
    Label1.Caption = mt.getMyTime
    Set mt = Nothing
End Sub

Private Sub saveLocalIndex()
    saveLocalIndexFlag = False
    If localIndex.Count = 0 Then Exit Sub
    Dim f%, j%, indexFile$
    
    f = FreeFile
    indexFile = localDir & strFTP(3)
    
    Open indexFile For Output As f
    For j = 0 To localIndex.Count - 1
        Write #f, localIndex.FileName(j), _
            localIndex.OldDate(j), localIndex.NewDate(j)
    Next
    Close f
End Sub

'========================================================
'        Routines Common to Upload & Download
'========================================================
'Get the remoteIndex of previously uploaded files.
'Populate the remoteIndex collection.
'Returns False if error encountered
Private Function getRemoteIndex() As Boolean
    Dim indexFile$, tmpFile$, b As Boolean
    Dim fData As WIN32_FIND_DATA
    
    getRemoteIndex = False
    
    fData.cFileName = String$(MAX_PATH, 0)
    indexFile = remoteDir & strFTP(3)
    tmpFile = WinTmpDir & strFTP(3)
    
    'See if the index file exists
    Dim hFind&
    DoEvents
    hFind = FtpFindFirstFile(hConnection, _
                    indexFile, fData, 0, 0)
    DoEvents
    
    'Either no file or an FTP problem
    If hFind = 0 Then Exit Function
    IClose hFind
    
    'Download the index file to a local temp file
    DoEvents
    b = FtpGetFile(hConnection, indexFile, _
        tmpFile, False, INTERNET_FLAG_RELOAD, _
        FTP_TRANSFER_TYPE_ASCII, 0)
    DoEvents
    
    'An FTP problem
    If b = False Then
        errMsg "FtpGetFile " & indexFile
        Exit Function
    End If
    
    'The Index file has been downloaded to a local
    '  temp file so now we open it and get all the
    '  file names and dates in it
    Dim f%, tmp$
    Dim dat As Date
    Set remoteIndex = New clsFileInfo
    
    f = FreeFile
    On Error GoTo up3
    Open tmpFile For Input As f
    
    On Error GoTo up2
    Do Until EOF(f)
        Input #f, tmp, dat
        remoteIndex.AddItem tmp, dat, dat
        DoEvents
    Loop
    
up1: On Error GoTo 0
    Close

    If Dir$(tmpFile) <> "" Then Kill tmpFile
    DoEvents
    getRemoteIndex = (remoteIndex.Count > 0)
    Exit Function
    
    'Error reading file
up2: Resume up1
    
    'Error opening file
up3: Resume up4
up4: On Error GoTo 0
End Function

'Populate the remoteFiles class with the filenames
'  of all the files actually in the remote directory.
Private Sub getRemoteFiles(spec$)
    Dim hFind&, tmp$
    Dim bRet As Boolean
    Dim fData As WIN32_FIND_DATA
    Set remoteFiles = New clsFileInfo
    
    fData.cFileName = String$(MAX_PATH, 0)
    DoEvents
    hFind = FtpFindFirstFile(hConnection, _
            remoteDir & spec, fData, 0, 0)
    DoEvents
    
    If hFind Then
        Do
            tmp = TrimNull(fData.cFileName)
            'Don't add the index file
            If tmp <> strFTP(3) Then
            'Don't add 0 length files
            If fData.nFileSizeLow Then
            'Don't add directories
            If (fData.dwFileAttributes And _
                vbDirectory) = 0 Then
                
                  remoteFiles.AddItem tmp, _
                    ConvertFileDate(fData.ftLastWriteTime)
                    
            End If
            End If
            End If
        
            fData.cFileName = String$(MAX_PATH, 0)
            DoEvents
            bRet = InternetFindNextFile(hFind, fData)
            DoEvents
        
            If Not bRet Then Exit Do
        Loop
        
        DoEvents
        IClose hFind
    End If
End Sub

'Get the local index of previously restored files.
'Populate the localIndex collection.
'Returns False if error encountered
Private Function getLocalIndex() As Boolean
    Dim indexFile$
    
    getLocalIndex = False
    indexFile = localDir & strFTP(3)
    
    'Open the index and get all the
    '  file names and dates in it
    'If the file does not exists the error handling
    '  will return True for this function
    Dim f%, tmp$
    Dim dat1 As Date, dat2 As Date
    Set localIndex = New clsFileInfo
    
    f = FreeFile
    On Error GoTo gl3
    Open indexFile For Input As f
    
    On Error GoTo gl2
    Do Until EOF(f)
        Input #f, tmp, dat1, dat2
        localIndex.AddItem tmp, dat1, dat2
        DoEvents
    Loop
    
gl1: On Error GoTo 0
    Close

    DoEvents
    getLocalIndex = (localIndex.Count > 0)
    Exit Function
    
    'Error reading file
gl2: Resume gl1
    
    'Error opening file
gl3: Resume gl4
gl4: On Error GoTo 0
End Function

'Populate the collection with matching localFileSpecs
' Also get the respective file dates
Private Sub getLocalFiles(ByVal spec$)
    Dim tmp$, hFile&
    Dim WFD As WIN32_FIND_DATA
    
    Set localFiles = Nothing
    Set localFiles = New clsFileInfo
    tmp = localDir & spec
    
    hFile = FindFirstFile(tmp, WFD)
    If hFile > 0 Then
        addLocalFile WFD
        
        While FindNextFile(hFile, WFD)
            addLocalFile WFD
        Wend
        FindClose hFile
    End If
End Sub

Private Sub addLocalFile(w As WIN32_FIND_DATA)
    Dim fn$, d As Date
    
    'Don't want directories
    If (w.dwFileAttributes And _
      FILE_ATTRIBUTE_DIRECTORY) = 0 Then
        fn = TrimNull(w.cFileName)
        
        'Don't add the index file
        If fn <> strFTP(3) Then
            
            'We don't want any zero length files
            If w.nFileSizeLow > 0 Then
                d = ConvertFileDate(w.ftLastWriteTime)
                localFiles.AddItem fn, d
            End If
        End If
    End If
    
    DoEvents
End Sub

'Check the local index against the actual local files
'  If no file exists then remove it from the index
Private Sub verifyLocalIndex()
    Dim j%, k%, tmp$
    
    Do Until j >= localIndex.Count
        tmp = localIndex.FileName(j)
        
        If localFiles.Exists(tmp, k) Then
            j = j + 1
        Else
            localIndex.RemoveItem tmp
        End If
    Loop
End Sub

'Select all files
Private Sub selectAllFiles()
    Dim j%
    
    Set remoteIndex = New clsFileInfo
        
    For j = 0 To List1.ListCount - 1
        List1.Selected(j) = True
        
        'And put all the files in the index
        remoteIndex.AddItem localFiles.FileName(j), , _
                            localFiles.OldDate(j)
    Next
End Sub

'Show nothing to do
Private Sub InformNothingToDo(Optional msg$ = "")
    If msg = "" Then msg = "FTP " & mnuChgOp.Caption
    
    Label1.Caption = "No files to " & _
                        mnuChgOp.Caption
    
    If AUTO = MANUAL Then
        MsgBox Label1.Caption, , msg
        Label1.Caption = ""
        errorcondition = True
    End If
End Sub


'========================================================
'            Menu Handling Routines
'========================================================
Private Sub mnuAutoMode_Click()
    'Toggle between 0 and 1
    If Not done Then
        done = True
        cycleCommands
    End If
    AUTO = 1 - AUTO
    showModes
    Command1.SetFocus
End Sub

Private Sub mnuChgOp_Click()
    'Toggle between 0 and 1
    If Not done Then
        done = True
        cycleCommands
    End If
    MODE = 1 - MODE
    showModes
    Command1.SetFocus
End Sub

Private Sub showModes()
    mnuAutoMode.Caption = IIf(AUTO = MANUAL, _
                        "M&anual", "&Automatic")
                        
    If MODE = DNLD Then
        frmMain.Caption = "FTP Restore to Local"
        mnuChgOp.Caption = "Restore"
    Else
        frmMain.Caption = "FTP Backup to Remote"
        mnuChgOp.Caption = "Backup"
    End If
    
    List1.Clear
    Label1.Caption = ""
End Sub

Private Sub mnuEditComms_Click()
    frmCmds.Show
End Sub

Private Sub mnuHelp_Click()
    Dim f%, otherFlag As Boolean
    
    otherFlag = helpFlag
    helpFlag = Not helpFlag
    
    Text1.Visible = helpFlag
    Command1.Visible = otherFlag
    Command2.Visible = otherFlag
    List1.Visible = otherFlag
    Label1.Visible = otherFlag
    mnuAutoMode.Enabled = otherFlag
    mnuChgOp.Enabled = otherFlag
    mnuEditComms.Enabled = otherFlag
    mnuEditTypes.Enabled = otherFlag
    mnuSettings.Enabled = otherFlag
            
    If helpFlag Then
        If Not HelpLoaded Then
            On Error GoTo hp1
            f = FreeFile
            Open txtfile0 For Input As f
            Text1.Text = Input(LOF(f), f)
            Close f
            HelpLoaded = True
        End If
    Else
        Command1.SetFocus
    End If
    
    GoTo hp2
    
hp1: Text1.Text = "Help File Not Found:" _
                & vbCrLf & vbCrLf & txtfile0
    Resume hp2
hp2: On Error GoTo 0
End Sub

Private Sub mnuSettings_Click()
    frmSet.Show vbModal, Me
End Sub


Private Sub mnuEditTypes_Click()
    frmType.Show vbModal, Me
End Sub


'========================================================
'          Command Button Routines
'========================================================
'After the Listbox is populated with each batch of
'  files, click the "Upload" button to upload
Private Sub Command1_Click()
    If Command1.Caption = "Start" Then
        If (strFTP(0) = "" And FTPflag) _
                Or strFTP(3) = "" Then
            MsgBox "Server Settings Not Made", , _
                "CAN NOT PROCEED"
            Exit Sub
        End If
        
        done = False
        Label1.Caption = ""
        Command1.Caption = IIf(MODE = UPLD, _
                            "Upload", "Download")
        Command2.Caption = "Cancel"
        
        frmCmds.Show
        frmCmds.FormInitList
        generalFTPopen
    Else
        If MODE = UPLD Then
            continueUpload
        Else
            continueDownload
        End If
    
        If AUTO = MANUAL Then
            Label1.Caption = "Command Complete"
            cycleCommands
        End If
    End If
End Sub

Private Sub Command2_Click()
    If Command2.Caption = "Exit" Then
        Unload Me
    Else
        Label1.Caption = ""
        cycleCommands
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape: mnuHelp_Click: KeyAscii = 0
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                        Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1: mnuHelp_Click: KeyCode = 0
    End Select
End Sub


'========================================================
'         General String Handling Routines
'========================================================
Private Function fileDir$(ByVal s$)
    Dim j%, l%
    
    l = Len(s)
    
    'Get the file directory of a full path spec
    '  by working back until you hit a slash
    '  return includes that slash
    For j = l To 1 Step -1
        Select Case Mid$(s, j, 1)
            Case BSLASH, FSLASH
                Exit For
        End Select
    Next
    
    fileDir = Left$(s, j)
End Function

Private Function FileName$(ByVal s$)
    Dim j%, l%
    
    l = Len(s)
    
    'Get the filename of a full path spec by working
    '  back until you hit a slash
    For j = l To 1 Step -1
        Select Case Mid$(s, j, 1)
            Case BSLASH, FSLASH
                Exit For
        End Select
    Next
    
    FileName = Right$(s, l - j)
End Function

Private Function fileExt$(ByVal s$)
    Dim j%, l%
    
    l = Len(s)
    fileExt = ""
    
    'Get the extension of a file name by working
    '  back until you hit a period or a slash
    For j = l To 2 Step -1
        Select Case Mid$(s, j, 1)
            Case BSLASH, FSLASH
                Exit For
            Case "."
                fileExt = Right$(LCase$(s), l - j)
                Exit For
        End Select
    Next
End Function

Private Function addSlash$(ByVal s$, ByVal sl$)
    If Right$(s, 1) <> sl Then s = s & sl
    addSlash = s
End Function


'========================================================
'               FTP Routines
'========================================================
Private Sub openhConnection()
    If hConnection Then Exit Sub
    
    hConnection = InternetConnect(hOpen, _
        strFTP(0), INTERNET_INVALID_PORT_NUMBER, _
        strFTP(1), strFTP(2), INTERNET_SERVICE_FTP, _
        INTERNET_FLAG_PASSIVE, 0)
        
    If hConnection = 0 Then errMsg "InternetConnect"
End Sub

Private Sub errMsg(s$)
    MsgBox s & vbCrLf & lastErr
End Sub

Private Function lastErr$()
    Dim code&, l&, s$
    
    If Err.LastDllError = _
            ERROR_INTERNET_EXTENDED_ERROR Then
        InternetGetLastResponseInfo code, vbNullString, l
        s = String(l + 1, 0)
        InternetGetLastResponseInfo code, s, l
        lastErr = "Err: " & code & " " & TrimNull(s)
    Else
        lastErr = "Unspecified FTP Error"
    End If
End Function


Private Function setTransferType&(fName$)
    Dim j&, k%
    
    j = FTP_TRANSFER_TYPE_BINARY
    
    'Check the extension
    If transExtens.Exists(fileExt(fName), k) Then
        j = FTP_TRANSFER_TYPE_ASCII
    End If
    
    'Check the fileName
    If transFNames.Exists(fName, k) Then
        j = IIf(j = FTP_TRANSFER_TYPE_ASCII, _
                    FTP_TRANSFER_TYPE_BINARY, _
                    FTP_TRANSFER_TYPE_ASCII)
    End If
    
    setTransferType = j
End Function

'========================================================
'         General Enter & Exit Routines
'========================================================
Private Sub Form_Activate()
    helpFlag = False
    Command1.SetFocus
End Sub

Private Sub Form_Load()
    Dim f%, frmMainLeft, frmMainTop
    Dim ht!, lt!, tp!, wd!
    
    'Set the INI file names
    inifile0 = App.Path & BSLASH & App.EXEName
    inifile1 = inifile0 & "1.ini" 'commands
    inifile2 = inifile0 & "2.ini" 'type ext
    inifile3 = inifile0 & "3.ini" 'type file
    txtfile0 = inifile0 & "0.txt" 'genl help
    txtfile1 = inifile0 & "2.txt" 'commands help
    txtfile2 = inifile0 & "4.txt" 'type help
    txtfile3 = inifile0 & "3.txt" 'set help
    inifile0 = inifile0 & "0.ini" 'genl ini
    
'========================================================
    'Initial settings
    f = FreeFile
    On Error GoTo errf0
    'The "0" INI file
    Open inifile0 For Input As f
    Input #f, strFTP(0), strFTP(1), strFTP(2), strFTP(3)
    Input #f, frmMainLeft, frmMainTop, MODE, AUTO
    Input #f, frmCmdsLeft, frmCmdsTop

    GoTo ovr0
errf0: 'Initial settings if no INI file
    strFTP(3) = "bacindex"
    frmMainLeft = 930: frmMainTop = 855
    frmCmdsLeft = 5955: frmCmdsTop = 1035
    Resume ovr0
ovr0: On Error GoTo 0
    Close
    
    Move frmMainLeft, frmMainTop, 5000, 3300
    List1.Move SPCR, SPCR, 3000, ScaleHeight - SPCR
    
    ht = 375
    tp = ht
    lt = 3000 + SPCR * 2
    wd = ScaleWidth - lt - SPCR
    Label1.Move lt, tp, wd, ht
    tp = tp + ht * 2
    Command1.Move lt, tp, wd, ht
    tp = tp + ht * 2
    Command2.Move lt, tp, wd, ht
    
    With Command3
        .Move Width, SPCR, 1, 1
        .Default = True
        .TabStop = False
    End With
    
    With Text1
        .BackColor = LYELLOW
        .Move SPCR, SPCR, ScaleWidth - SPCR * 2, _
                    ScaleHeight - SPCR * 2
        .Visible = False
    End With
    
    showModes
    Command1.Caption = "Start"
    Command2.Caption = "Exit"
    
    'Function in modAPI specifying FTP or copy
    FTPflag = SetFTPflag
    
    'Note the .Hide will first Load Form
    ' Command strings initialized in frmCmds
    frmCmds.Hide
    ' Transfer types initialized in frmType
    frmType.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim f%
    
    If Not done Then
        done = True
        cycleCommands
    End If
    
    IClose hConnection
    IClose hOpen
    
    f = FreeFile
    Open inifile0 For Output As f
    'The "0" INI file
    Write #f, strFTP(0), strFTP(1), strFTP(2), strFTP(3)
    Write #f, Left, Top, MODE, AUTO
    Write #f, frmCmdsLeft, frmCmdsTop
    Close f
    
    Unload frmCmds
    Unload frmType
    Unload frmSet
    
    Set frmCmds = Nothing
    Set frmType = Nothing
    Set frmSet = Nothing
    
    Set localFiles = Nothing
    Set localIndex = Nothing
    Set remoteFiles = Nothing
    Set remoteIndex = Nothing
End Sub

Private Sub IClose(n&)
    If n Then
        InternetCloseHandle n
        n = 0
    End If
End Sub
