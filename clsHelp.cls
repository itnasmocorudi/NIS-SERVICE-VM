VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BclsHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function compareMinute(minuteFirst As String, minuteLast As String) As Boolean
Dim firstMinute As Integer
Dim lastMinute As Integer

    ' FIRST MINUTE > LAST MINUTE
    compareMinute = False
    
    firstMinute = (CDbl(Left(minuteFirst, InStr(minuteFirst, ":") - 1)) * 60) + CDbl(Mid(minuteFirst, InStr(minuteFirst, ":") + 1, Len(minuteFirst) - InStr(minuteFirst, ":")))
    lastMinute = (Left(minuteLast, InStr(minuteLast, ":") - 1) * 60) + Mid(minuteLast, InStr(minuteLast, ":") + 1, Len(minuteLast) - InStr(minuteLast, ":"))
    
    If firstMinute < lastMinute Then
        compareMinute = True
    End If
End Function
