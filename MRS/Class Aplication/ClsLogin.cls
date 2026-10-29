VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSQL As String

Public Function FetchItemPassword(isUser As String) As ADODB.Recordset
    On Error GoTo ErrHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "Select PasswordUser FROM T_User Where NamaUser = '" & UCase(isUser) & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchItemPassword = rs
        
    Exit Function
ErrHandler:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
