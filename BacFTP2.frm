VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCmds 
   Caption         =   "Command Strings"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   330
      Left            =   4560
      TabIndex        =   8
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&AddNew"
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "&Delete"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuClear 
      Caption         =   "&Clear"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuBrowse 
      Caption         =   "&Browse"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmCmds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim saveFlag As Boolean
Dim helpFlag As Boolean
Dim HelpLoaded As Boolean
Dim cmmands As clsFileInfo

Private Sub Form_Activate()
    FormInitList
    saveFlag = False
    helpFlag = False
    HelpLoaded = False
    Text1.SetFocus
End Sub

Public Sub FormInitList()
    Dim j%
    
    List1.Clear
    For j = 0 To cmmands.Count - 1
        List1.AddItem cmmands.FileName(j)
    Next
End Sub

Private Sub Form_Load()
    Dim ht!, lt!, tp!, wd!, wd1!
    Move frmCmdsLeft, frmCmdsTop, 5300, 3800
    
    wd = (ScaleWidth - SPCR * 3) / 2
    With Label1
        .Move SPCR, SPCR, wd, UNIT
        .Alignment = vbCenter
        .Caption = "Source"
    End With
    
    lt = wd + SPCR * 2
    With Label2
        .Move lt, SPCR, wd, UNIT
        .Alignment = vbCenter
        .Caption = "Destination"
    End With
    
    ht = 375
    tp = UNIT + SPCR
    Text1.Move SPCR, tp, wd, ht
    Text2.Move lt, tp, wd, ht
    
    tp = tp + ht + SPCR
    ht = ScaleHeight - UNIT - SPCR
    With Command1
        .Move SPCR, ht, wd, UNIT
        .Caption = "OK"
    End With
    
    With Command2
        .Move lt, ht, wd, UNIT
        .Caption = "Cancel"
    End With
    
    With Command3
        .Move Width, SPCR, 1, 1
        .Default = True
        .TabStop = False
    End With
    
    wd1 = ScaleWidth - SPCR * 2
    ht = ht - tp - SPCR
    List1.Move SPCR, tp, wd1, ht
    
    With Text3
        .BackColor = LYELLOW
        .Move SPCR, SPCR, wd1, ScaleHeight - SPCR * 2
        .Visible = False
    End With
    
    Caption = "Command Strings"
    loadCommands
End Sub

Private Sub loadCommands()
    Dim f%, tmp$
    
    f = FreeFile
    Set cmmands = New clsFileInfo
    
    On Error GoTo errf1
    Open inifile1 For Input As f
    
    Do Until EOF(f)
        Line Input #f, tmp
        cmmands.AddItem tmp
    Loop

    GoTo ovr2
errf1: Resume ovr2
ovr2: On Error GoTo 0
    Close
End Sub

Private Sub mnuAdd_Click()
    Dim j%, tmp$
    
    If fixTexts() = 0 Then Exit Sub
    
    tmp = Text1.Text & " to " & Text2.Text
    
    For j = 0 To List1.ListCount - 1
        If List1.List(j) = tmp Then
            List1.Selected(j) = True
            List1_Click
            Exit For
        End If
    Next
    
    If j >= List1.ListCount Then
        List1.AddItem tmp
        saveFlag = True
    End If
    
    Text1.SetFocus
End Sub

Private Sub mnuEdit_Click()
    Dim j%, tmp$
    
    If fixTexts() = 0 Then Exit Sub
    
    tmp = Text1.Text & " to " & Text2.Text
    
    For j = 0 To List1.ListCount - 1
        If List1.Selected(j) Then
            List1.List(j) = tmp
            saveFlag = True
            Exit For
        End If
    Next
    
    Text1.SetFocus
End Sub

Private Sub mnuDelete_Click()
    Dim j%
    
    For j = 0 To List1.ListCount - 1
        If List1.Selected(j) Then
            If MsgBox("Confirm delete please.", vbYesNo, _
                List1.List(j)) = vbNo Then Exit For
            List1.RemoveItem j
            mnuClear_Click
            saveFlag = True
            Exit For
        End If
    Next
End Sub

Private Sub mnuClear_Click()
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub mnuHelp_Click()
    Dim f%, otherFlag As Boolean
    
    otherFlag = helpFlag
    helpFlag = Not helpFlag
    
    Command1.Visible = otherFlag
    Command2.Visible = otherFlag
    Text1.Visible = otherFlag
    Text2.Visible = otherFlag
    Text3.Visible = helpFlag
    List1.Visible = otherFlag
    mnuAdd.Enabled = otherFlag
    mnuEdit.Enabled = otherFlag
    mnuDelete.Enabled = otherFlag
    mnuClear.Enabled = otherFlag
    mnuBrowse.Enabled = otherFlag
            
    If helpFlag Then
        If Not HelpLoaded Then
            On Error GoTo hp1
            f = FreeFile
            Open txtfile1 For Input As f
            Text3.Text = Input(LOF(f), f)
            Close f
            HelpLoaded = True
        End If
    Else
        Text1.SetFocus
    End If
    
    GoTo hp2
    
hp1: Text3.Text = "Help File Not Found:" _
                & vbCrLf & vbCrLf & txtfile1
    Resume hp2
hp2: On Error GoTo 0
End Sub

Private Sub mnuBrowse_Click()
    Static lastdir$, lastname$
    
    If lastdir = "" Then lastdir = "c:\"
    On Error GoTo errd1
    
    With CommonDialog1
        .CancelError = True
        .InitDir = lastdir
        .Filter = "All Files|*.*"
        .ShowOpen
        lastdir = .FileName
        Text1.Text = .FileName
    End With
    
    GoTo errd2

errd1: Resume errd2
errd2: On Error GoTo 0
    Text1.SetFocus
End Sub

Private Sub List1_Click()
    If doNotExecute Then Exit Sub
    Dim j%, k%, l%, tmp$
    
    For j = 0 To List1.ListCount - 1
        If List1.Selected(j) Then
            mnuClear_Click
            tmp = List1.List(j)
            k = InStr(tmp, " to ")
            l = Len(tmp)
            If k > 1 Then Text1.Text = Left$(tmp, k - 1)
            k = k + 3
            If k < l Then Text2.Text = Right$(tmp, l - k)
            Exit For
        End If
    Next
End Sub

Private Function fixTexts%()
    Dim j%, tmp1$, tmp2$
    
    tmp2 = fixText(Text2.Text, BSLASH)
    'If Len(tmp2) = 0 Then tmp2 = FSLASH
    Text2.Text = tmp2
    
    tmp1 = fixText(Text1.Text, FSLASH)
    fixTexts = Len(tmp1)
    Text1.Text = tmp1
    
    If Len(tmp1) = 0 Then
        MsgBox "No source command entered"
        Text1.SetFocus
    End If
End Function

Private Function fixText$(s1$, s2$)
    fixText = LCase$(Trim$(fixSlash(s1, s2$)))
End Function

Private Sub Text1_GotFocus()
    mnuBrowse.Visible = True
End Sub

Private Sub Text2_GotFocus()
    mnuBrowse.Visible = False
End Sub

Private Sub List1_GotFocus()
    mnuBrowse.Visible = False
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
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

'The OK Button
Private Sub Command1_Click()
    Unload Me
End Sub

'The Cancel Button
Private Sub Command2_Click()
    saveFlag = False
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCmdsLeft = Left
    frmCmdsTop = Top
    
    If saveFlag Then saveCommands
End Sub

Private Sub saveCommands()
    Dim f%, j%
    
    f = FreeFile
    Set cmmands = Nothing
    Set cmmands = New clsFileInfo
    
    Open inifile1 For Output As f
    For j = 0 To List1.ListCount - 1
        cmmands.AddItem List1.List(j)
        Print #f, List1.List(j)
    Next
    
    Close f
    saveFlag = False
End Sub

Private Sub Form_Terminate()
    Set cmmands = Nothing
End Sub

