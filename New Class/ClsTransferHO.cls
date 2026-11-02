VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTransferHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private mConnNas As New ADODB.Connection
Private mConnNrm As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSQL As String

Public Sub AddNew_LogHO(noWO As String, Status As String, tipe As String, isSQl As String)
1     On Error GoTo errHandler
2        If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
            Dim cm As New Command
3             strSQL = "Select * FROM LogHO WITH (NOLOCK) where kode = '" & noWO & "' and tipe = '" & tipe & "'"
          
4             Set rs = New ADODB.Recordset
5             rs.CursorLocation = adUseClient
6             rs.Open strSQL, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
7             Set rs.ActiveConnection = Nothing
8             If rs.EOF Then
9                 strSQL = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                     "Values ( '" & tipe & "','" & noWO & "', '" & Status & "', 0,'01/01/1900', '" & Replace(isSQl, "'", "~") & "')"
            
10                cm.CommandText = strSQL
11                cm.CommandType = adCmdText
                    
12                Set cm.ActiveConnection = mConn
13                cm.Execute
14            End If
            
15          If mConn.State = adStateOpen Then mConn.Close
16          Set cm = Nothing
17          Exit Sub
errHandler:
18          If mConn.State = adStateOpen Then mConn.Close
19          mConn.RollbackTrans
20        If Not cm Is Nothing Then Set cm = Nothing
21            err.Raise err.Number, err.Source, err.Description
          
End Sub

Public Function FetchItemTransfer() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSQL = "Select Top 1 * FROM LogHO WITH (NOLOCK) where LogHO.LogHO = 0"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchItemTransfer = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function


Public Function OpenTransactionLog_NRM() As ADODB.Recordset
    errHO = False
    If mConnNrm.State = adStateClosed Then mConnNrm.Open (SQLConBengkel99)
    mConnNrm.BeginTrans
End Function

Public Function CloseTransactionLog_NRM() As ADODB.Recordset
    On Error Resume Next
    If errHO = True Then
        mConnNrm.RollbackTrans
    Else
        mConnNrm.CommitTrans
    End If
    If mConnNrm.State = adStateOpen Then mConnNrm.Close
End Function

Public Function OpenTransactionLog() As ADODB.Recordset
    errHO = False
    If mConnNas.State = adStateClosed Then mConnNas.Open (SQLCONBengkel)
    If mConnNrm.State = adStateClosed Then mConnNrm.Open (SQLConBengkel99)
    mConnNas.BeginTrans
    mConnNrm.BeginTrans
End Function

Public Function CloseTransactionLog() As ADODB.Recordset
    On Error Resume Next
    mConnNas.CommitTrans
    mConnNrm.CommitTrans
    If mConnNas.State = adStateOpen Then mConnNas.Close
    If mConnNrm.State = adStateOpen Then mConnNrm.Close
End Function

'==================================================================
'REAL TIME
'==================================================================
Public Sub AddNew_WO(strSQLx As String)
1        On Error GoTo errHandler
            Dim cm As New Command
2           cm.CommandText = strSQLx
3           cm.CommandType = adCmdText
              
4           Set cm.ActiveConnection = mConnNrm
5           cm.Execute
6           Set cm = Nothing
7           Exit Sub
errHandler:
8         If Not cm Is Nothing Then Set cm = Nothing
9         errHO = True
10        mConnNrm.RollbackTrans
              'err.Raise err.Number, err.Source, err.Description
11            'MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure AddNew_Billing of Class Module ClsTransferHO")
End Sub


Public Sub AddNew_Billing(strSQLx As String)
1        On Error GoTo errHandler
            Dim cm As New Command
2           cm.CommandText = Replace(Replace(strSQLx, "')", "','" & Log.LogBengkel & "')"), ") VALUES (", ", CABANG) VALUES (")
3           cm.CommandType = adCmdText
              
4           Set cm.ActiveConnection = mConnNrm
5           cm.Execute
6           Set cm = Nothing
7           Exit Sub
errHandler:
8         If Not cm Is Nothing Then Set cm = Nothing
9         errHO = True
10        mConnNrm.RollbackTrans
              'err.Raise err.Number, err.Source, err.Description
11        'MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure AddNew_Billing of Class Module ClsTransferHO")
End Sub

Public Sub AddNew_Part(strSQLx As String)
1        On Error GoTo errHandler
            Dim cm As New Command
2           cm.CommandText = strSQLx
3           cm.CommandType = adCmdText
              
4           Set cm.ActiveConnection = mConnNrm
5           cm.Execute
6           Set cm = Nothing
7           Exit Sub
errHandler:
8         If Not cm Is Nothing Then Set cm = Nothing
9         errHO = True
10        mConnNrm.RollbackTrans
              'err.Raise err.Number, err.Source, err.Description
11        'MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure AddNew_Billing of Class Module ClsTransferHO")
End Sub

Public Function FetchDataHO(strSQL) As ADODB.Recordset
    On Error GoTo errHandler
     
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, mConnNrm, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataHO = rs
    Exit Function
errHandler:
          errHO = True
         mConnNrm.RollbackTrans
End Function

Public Sub Remove_WO(strSQLx As String)
1        On Error GoTo errHandler
            Dim cm As New Command
2           cm.CommandText = strSQLx
3           cm.CommandType = adCmdText
              
4           Set cm.ActiveConnection = mConnNrm
5           cm.Execute
6           Set cm = Nothing
7           Exit Sub
errHandler:
8         If Not cm Is Nothing Then Set cm = Nothing
9         errHO = True
10        mConnNrm.RollbackTrans
            '  err.Raise err.Number, err.Source, err.Description
11        'MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure AddNew_Billing of Class Module ClsTransferHO")
End Sub

