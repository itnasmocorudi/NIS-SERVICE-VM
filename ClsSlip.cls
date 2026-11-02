VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSQL As String

Public Function FetchDataSlip() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "Select IsNull(SlipA,'') as SlipA, IsNull(SlipAr,'') as SlipAr, IsNull(SlipB,'') as SlipB, IsNull(SlipBr,'') as SlipBr, " & _
                    "IsNull(SlipC,'') as SlipC, IsNull(SlipCr,'') as SlipCr, IsNull(SlipE,'') as SlipE, IsNull(SlipEr,'') as SlipEr, " & _
                    "IsNull(SlipG,'') as SlipG, IsNull(SlipGr,'') as SlipGr from SlipCode WITH (NOLOCK) Order By Urut ASC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataSlip = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function


Public Sub AddNew_MasterSlip(nUrut As Integer)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSQL = "Insert Into SlipCode values (" & nUrut & ",'','','','','','','','','','')"
               
      cm.CommandText = strSQL
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
      If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub Update_MasterSlipItem(nUrut As Integer, nData As String)
   On Error GoTo ErrUpdate
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      
      strSQL = "Update SlipCode Set " & nData & " Where Urut = " & nUrut
               
      cm.CommandText = strSQL
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrUpdate:
      If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub


Public Function FetchtipePart() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "Select TipePart FROM TipePart WITH (NOLOCK)"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchtipePart = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_TipePart(nPart As String)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSQL = "Insert Into TipePart (TipePart) values ('" & nPart & "')"
               
      cm.CommandText = strSQL
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
      If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function FetchListPart() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "Select TipePart,KodeItem,namaPart,qty,[ID] FROM TipePartList WITH (NOLOCK)"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchListPart = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchListPart_NISPART(iskode As Boolean, isCari As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    strSQL = "Select Distinct Right(Kodeitem, len(KodeItem)-2) as Kode, NamaBarang from INMST WITH (NOLOCK)"
    
    If iskode = True Then
        strSQL = strSQL & " Where Kodeitem like '%" & isCari & "%' and Left(Kodeitem,1) <> 'X'"
    Else
        strSQL = strSQL & " Where NamaBarang like '%" & isCari & "%' and Left(Kodeitem,1) <> 'X'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchListPart_NISPART = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchNamaPart_NISPART(iskode As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    strSQL = "Select NamaBarang from INMST WITH (NOLOCK) Where Right(Kodeitem, len(KodeItem)-2) = '" & iskode & "'"
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchNamaPart_NISPART = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_ListPart(TipePart As String, KodeItem As String, namaPart As String, qty As Double)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSQL = "Insert Into TipePartList (TipePart,KodeItem,namaPart,qty) values ('" & TipePart & "','" & KodeItem & "','" & namaPart & "'," & qty & ")"
               
      cm.CommandText = strSQL
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
      If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub


Public Sub UpdateItem_ListPart(TipePart As String, KodeItem As String, namaPart As String, qty As Double, nID As Double)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSQL = "update TipePartList set TipePart='" & TipePart & "',KodeItem='" & KodeItem & "',namaPart='" & namaPart & "',qty=" & qty & " Where [ID] = " & nID
               
      cm.CommandText = strSQL
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
      If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub DeleteItem_ListPart(nID As Double)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSQL = "Delete FROM TipePartList Where [ID] = " & nID
               
      cm.CommandText = strSQL
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
      If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub


