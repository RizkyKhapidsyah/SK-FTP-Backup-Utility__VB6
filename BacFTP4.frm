VERSION 5.00
Begin VB.Form frmType 
   Caption         =   "Transfer Types"
   ClientHeight    =   3390
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   3390
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Specific Files"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Extensions"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum focusFlags
    NONE
    TXT1
    TXT2
End Enum

Dim focusFlag As focusFlags
Dim saveFlag1 As Boolean
Dim saveFlag2 As Boolean
Dim HelpLoaded As Boolean
Dim helpFlag As Boolean

Private Sub Form_Activate()
    Dim j%
    
    List1.Clear
    For j = 0 To transExtens.Count - 1
        List1.AddItem transExtens.FileName(j)
    Next
    List2.Clear
    For j = 0 To transFNames.Count - 1
        List2.AddItem transFNames.FileName(j)
    Next
    
    helpFlag = False
    saveFlag1 = False
    saveFlag2 = False
    HelpLoaded = False
    Move frmMain.Left + 1000, frmMain.Top + 1000
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Dim lt!, ht!, tp1!, tp2!, wd1!, wd2!
    
    Move 0, 0, 4800, 3800
    
    wd1 = 1200
    tp1 = UNIT - 40
    lt = wd1 + SPCR * 2
    wd2 = ScaleWidth - lt - SPCR
    With Label1
        .Move SPCR, SPCR, wd1, tp1
        .Caption = "Extensions"
        .Alignment = vbCenter
    End With
    With Label2
        .Move lt, SPCR, wd2, tp1
        .Caption = "Specific Files"
        .Alignment = vbCenter
    End With
    
    tp1 = tp1 + SPCR
    With Text1
        .Move SPCR, tp1, wd1, UNIT
        .Text = ""
    End With
    With Text2
        .Move lt, tp1, wd2, UNIT
        .Text = ""
    End With
    
    tp1 = tp1 + SPCR + UNIT
    tp2 = ScaleHeight - SPCR - UNIT
    ht = tp2 - tp1 - SPCR
    With List1
        .Move SPCR, tp1, wd1, ht
        .TabStop = False
    End With
    With List2
        .Move lt, tp1, wd2, ht
        .TabStop = False
    End With
    
    wd1 = (ScaleWidth - SPCR * 4) / 3
    With Command1
        .Move SPCR, tp2, wd1, UNIT
        .Caption = "OK"
    End With
    lt = wd1 + SPCR * 2
    With Command2
        .Move lt, tp2, wd1, UNIT
        .Caption = "Help"
    End With
    lt = lt + wd1 + SPCR
    With Command3
        .Move lt, tp2, wd1, UNIT
        .Caption = "Cancel"
    End With
    With Command4
        .Move Width, SPCR, 1, 1
        .Default = True
        .TabStop = False
    End With
    
    With Text3
        .BackColor = LYELLOW
        .Move SPCR, SPCR, ScaleWidth - SPCR * 2, _
                tp2 - SPCR * 2
        .Visible = False
    End With
    
    loadTypes
End Sub

Private Sub loadTypes()
    Dim f1%, f2%, tmp$
    
    'The type extensions
    Set transExtens = New clsFileInfo

    f1 = FreeFile
    On Error GoTo errf1
    Open inifile2 For Input As f1
    Do Until EOF(f1)
        Line Input #f1, tmp
        transExtens.AddItem tmp
    Loop
    
    GoTo ovr1
errf1: Resume ovr1
ovr1: On Error GoTo errf2

    'The type names
    Set transFNames = New clsFileInfo

    f2 = FreeFile
    Open inifile3 For Input As f2
    Do Until EOF(f2)
        Line Input #f2, tmp
        transFNames.AddItem tmp
    Loop
    
    GoTo ovr2
errf2: Resume ovr2
ovr2: On Error GoTo 0
    Close
End Sub

'The OK Button
Private Sub Command1_Click()
    Unload Me
End Sub

'The Help Button
Private Sub Command2_Click()
    Dim f%, otherFlag As Boolean
    
    otherFlag = helpFlag
    helpFlag = Not helpFlag
    
    Text1.Visible = otherFlag
    Text2.Visible = otherFlag
    Text3.Visible = helpFlag
    List1.Visible = otherFlag
    List2.Visible = otherFlag
    Command1.Visible = otherFlag
    Command3.Visible = otherFlag
            
    If helpFlag Then
        If Not HelpLoaded Then
            On Error GoTo hp1
            f = FreeFile
            Open txtfile2 For Input As f
            Text3.Text = Input(LOF(f), f)
            Close f
            HelpLoaded = True
        End If
    Else
        Text1.SetFocus
    End If
    
    GoTo hp2
    
hp1: Text3.Text = "Help File Not Found:" _
                & vbCrLf & vbCrLf & txtfile2
    Resume hp2
hp2: On Error GoTo 0
End Sub

'The Cancel Button
Private Sub Command3_Click()
    saveFlag1 = False
    saveFlag2 = False
    
    Unload Me
End Sub

Private Sub Command4_Click()
    Select Case focusFlag
        Case TXT1: addItems Text1, List1
        Case TXT2: addItems Text2, List2
    End Select
End Sub

Private Sub List1_Click()
    List1.RemoveItem List1.ListIndex
    saveFlag1 = True
    Text1.SetFocus
End Sub

Private Sub List2_Click()
    List2.RemoveItem List2.ListIndex
    saveFlag2 = True
    Text2.SetFocus
End Sub

Private Sub Text1_GotFocus()
    focusFlag = TXT1
End Sub

Private Sub Text1_LostFocus()
    focusFlag = NONE
End Sub

Private Sub Text2_GotFocus()
    focusFlag = TXT2
End Sub

Private Sub Text2_LostFocus()
    focusFlag = NONE
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape: Command2_Click: KeyAscii = 0
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                        Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1: Command2_Click: KeyCode = 0
    End Select
End Sub

Private Sub addItems(t As TextBox, l As ListBox)
    Dim j%, tmp$
    
    tmp = LCase$(Trim$(t))
    If tmp = "" And focusFlag = TXT2 Then Exit Sub
    
    With l
        For j = 0 To .ListCount - 1
            If tmp = .List(j) Then Exit For
        Next
        
        If j = .ListCount Then
            .AddItem tmp
            If focusFlag = TXT1 Then
                saveFlag1 = True
            Else
                saveFlag2 = True
            End If
        End If
    End With
    
    t.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim f%, j%
        
    If saveFlag1 Then
        f = FreeFile
        Set transExtens = Nothing
        Set transExtens = New clsFileInfo
        Open inifile2 For Output As f
        
        For j = 0 To List1.ListCount - 1
            transExtens.AddItem List1.List(j)
            Print #f, transExtens.FileName(j)
        Next
        
        Close f
    End If

    If saveFlag2 Then
        f = FreeFile
        Set transFNames = Nothing
        Set transFNames = New clsFileInfo
        Open inifile3 For Output As f
        
        For j = 0 To List2.ListCount - 1
            transFNames.AddItem List2.List(j)
            Print #f, transFNames.FileName(j)
        Next
        
        Close f
    End If
End Sub

Private Sub Form_Terminate()
    Set transExtens = Nothing
    Set transFNames = Nothing
End Sub

