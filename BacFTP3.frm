VERSION 5.00
Begin VB.Form frmSet 
   Caption         =   "FTP Settings"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "BacFTP3.frx":0000
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Server Address"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum infoType
    SERVER
    USER
    PASSW
    XFILE
End Enum

Dim HelpLoaded As Boolean
Dim helpFlag As Boolean

'The OK Button
Private Sub Command1_Click()
    Dim j%
    
    For j = 0 To 3
        strFTP(j) = Text1(j).Text
    Next
    
    Unload Me
End Sub

'The Help Button
Private Sub Command2_Click()
    Dim f%, otherFlag As Boolean
    Dim j As infoType
    
    otherFlag = helpFlag
    helpFlag = Not helpFlag
    
    Text2.Visible = helpFlag
    For j = SERVER To XFILE
        Label1(j).Visible = otherFlag
        Text1(j).Visible = otherFlag
    Next
    
    Command1.Visible = otherFlag
    Command3.Visible = otherFlag
            
    If helpFlag Then
        If Not HelpLoaded Then
            On Error GoTo hp1
            f = FreeFile
            Open txtfile3 For Input As f
            Text2.Text = Input(LOF(f), f)
            Close f
            HelpLoaded = True
        End If
    Else
        Text1(0).SetFocus
    End If
    
    GoTo hp2
    
hp1: Text2.Text = "Help File Not Found:" _
                & vbCrLf & vbCrLf & txtfile3
    Resume hp2
hp2: On Error GoTo 0
End Sub

'The Cancel Button
Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Text2_Click()
    Command2_Click
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Command2_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                        Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1: Command2_Click: KeyCode = 0
    End Select
End Sub

Private Sub Form_Activate()
    Dim j As infoType
    
    For j = SERVER To XFILE
        Text1(j).Text = strFTP(j)
    Next
    
    Move frmMain.Left + 1000, frmMain.Top + 1000
    
    HelpLoaded = False
    helpFlag = False
    Text1(0).SetFocus
End Sub

Private Sub Form_Load()
    Dim j As infoType
    Dim lt!, tp!, ht!, wd!
    
    Width = 3800
    
    ht = 375
    tp = SPCR
    lt = 1100 + SPCR * 2
    wd = ScaleWidth - 1100 - SPCR * 3
    
    For j = SERVER To XFILE
        If j Then
            Load Label1(j)
            Load Text1(j)
        End If
        
        With Label1(j)
            .Move SPCR, tp + 40, 1100, ht
            .Visible = True
            Select Case j
                Case SERVER: .Caption = "Server Address"
                Case USER: .Caption = "User"
                Case PASSW: .Caption = "Password"
                Case XFILE: .Caption = "Index Name"
            End Select
        End With
        With Text1(j)
            .Move lt, tp, wd, ht
            .Text = ""
            .Visible = True
        End With
        
        tp = tp + SPCR + ht
    Next
    
    wd = (ScaleWidth - SPCR * 4) / 3
    With Command1
        .Caption = "OK"
        .Move SPCR, tp, wd, ht
    End With
    lt = wd + SPCR * 2
    With Command2
        .Caption = "Help"
        .Move lt, tp, wd, ht
    End With
    lt = lt + wd + SPCR
    With Command3
        .Caption = "Cancel"
        .Move lt, tp, wd, ht
    End With
    With Command4
        .Move Width, SPCR, 1, 1
        .Default = True
        .TabStop = False
    End With
        
    Height = Height - ScaleHeight + tp + ht + SPCR
    
    With Text2
        .BackColor = LYELLOW
        .Move SPCR, SPCR, ScaleWidth - SPCR * 2, _
                ScaleHeight - SPCR * 2
        .Visible = False
    End With
End Sub
