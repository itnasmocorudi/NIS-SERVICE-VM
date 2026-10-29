VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsManPower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Public Function FetchListManPower() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select * From manPower WITH (NOLOCK)"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchListManPower = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_ManPower(nkode As String, nNama As String, nTipe As String, nLevel As String, nik As String)
   On Error GoTo errHandler
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Insert Into manPower (Kode,Nama,Tipe,[Level],nik) " & _
               "Values('" & nkode & "','" & nNama & "','" & nTipe & "','" & nLevel & "','" & nik & "')"
               
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub Update_ManPower(nkode As String, nNama As String, nTipe As String, nLevel As String, nik As String)
   On Error GoTo errHandler
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Update ManPower " & _
               "Set Nama = '" & nNama & "',Tipe = '" & nTipe & "',[Level] = '" & nLevel & "',nik = '" & nik & "' " & _
               "Where Kode = '" & nkode & "'"
               
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub


Public Sub Delete_ItemManPower(nkode As String)
   On Error GoTo errHandler
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      
        strSql = "DELETE ManPower WHERE Kode = '" & nkode & "'"

      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub


