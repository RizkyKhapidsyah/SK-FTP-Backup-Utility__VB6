VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Dim d As Date

Public Function getMyTime$()
    Dim min&, sec&
    
    sec = DateDiff("s", d, Now)
    min = sec \ 60
    sec = sec Mod 60
    
    getMyTime = min & " min " & sec & " sec"
End Function

Private Sub Class_Initialize()
    d = Now
End Sub
