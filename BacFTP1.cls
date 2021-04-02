VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Const DIMINCR = 32

Private Type indexType
    Name1 As String
    Date1 As Date
    Date2 As Date
End Type

Public Count%
Dim dimCntr%
Dim info() As indexType

Private Sub Class_Initialize()
    dimCntr = DIMINCR
    ReDim info(dimCntr)
End Sub

Public Sub AddItem(nam$, Optional oldD As Date = 0, _
                        Optional newD As Date = 0)
                        
    If Count >= dimCntr Then
        dimCntr = dimCntr + DIMINCR
        ReDim Preserve info(dimCntr + 1)
    End If
    
    info(Count).Name1 = LCase$(nam)
    info(Count).Date1 = oldD
    info(Count).Date2 = newD
    
    Count = Count + 1
End Sub

Public Sub RemoveItem(ByVal nam$)
    Dim j%
    
    nam = LCase$(nam)
    
    For j = 0 To Count - 1
        If nam = info(j).Name1 Then Exit For
    Next
    
    If j < Count Then
        Count = Count - 1
        For j = j To Count
            info(j) = info(j + 1)
        Next
    End If
End Sub

Public Function Exists(ByVal nam$, n%) As Boolean
    nam = LCase$(nam)
    
    For n = Count - 1 To 0 Step -1
        If nam = info(n).Name1 Then Exit For
    Next
    
    Exists = (n >= 0)
End Function

Public Property Get FileName$(j%)
    If j < 0 Or j >= Count Then
        FileName = ""
    Else
        FileName = info(j).Name1
    End If
End Property

Public Property Get OldDate(j%) As Date
    If j < 0 Or j >= Count Then
        OldDate = 0
    Else
        OldDate = info(j).Date1
    End If
End Property

Public Property Get NewDate(j%) As Date
    If j < 0 Or j >= Count Then
        NewDate = 0
    Else
        NewDate = info(j).Date2
    End If
End Property

Public Property Let FileName(j%, nam$)
    If j >= 0 And j < Count Then
        info(j).Name1 = LCase$(nam)
    End If
End Property

Public Property Let OldDate(j%, d As Date)
    If j >= 0 And j < Count Then
        info(j).Date1 = d
    End If
End Property

Public Property Let NewDate(j%, d As Date)
    If j >= 0 And j < Count Then
        info(j).Date2 = d
    End If
End Property
