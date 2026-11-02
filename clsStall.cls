VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Public Function FetchQtyStall() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "Select top 1 * from Qty_Stall WITH (NOLOCK) Order By Tanggal DESC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchQtyStall = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_QtyStall(nEM As Double, nNP As Double, nPro As Double, nNpro As Double, nTHS As Double, nSP As Double)
   On Error GoTo errHandler
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
      strSql = "Insert Into QTY_Stall values (getdate()," & nEM & ", " & nNP & ", " & nPro & ", " & nNpro & ",'" & gSA & "', " & nTHS & ", " & nSP & ")"
                 
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
      If Not cm Is Nothing Then Set cm = Nothing
err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub AddNew_QtyStall_Log(nEM As Double, nNP As Double, nPro As Double, nNpro As Double, nTHS As Double, nSP As Double)
   On Error GoTo errHandler
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLConProfilBK)
      strSql = "Insert Into LogStall values (getdate()," & nEM & ", " & nNP & ", " & nPro & ", " & nNpro & ",'" & gSA & "','" & ComptName & "', " & nTHS & "," & nSP & ")"
                 
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
      If Not cm Is Nothing Then Set cm = Nothing
err.Raise err.Number, err.Source, err.Description
End Sub

Public Function FetchQtyStall_aDay(nHari As Double, nBulan As Double, nTahun As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "Select top 1 * from Qty_Stall WITH (NOLOCK) Where Day(Tanggal) <= " & nHari & " and Month(Tanggal) <= " & nBulan & " and Year(Tanggal) <= " & nTahun & " Order By Tanggal DESC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchQtyStall_aDay = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchTotalUnit_NP_EM(nHari As Double, nBulan As Double, nTahun As Double, isEM As Integer) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "Select Count(No_WO) from JOBDTA WITH (NOLOCK) Where FS_LOG_DEL = '' and FS_RC not in ('','R001') and Day([date]) = " & nHari & " and Month([date]) = " & nBulan & " and Year([date]) = " & nTahun & " AND No_WO in " & _
    "(select No_WO From T_invoice Where Day(InvDate) = " & nHari & " and Month(InvDate) = " & nBulan & " and Year(InvDate) = " & nTahun & " and FS_LOG_DEL = '' ) "
    
    If isEM = 1 Then
        strSql = strSql & " And Code = 1"
    Else
        strSql = strSql & " And Kd_K = 1"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount = 0 Then
        FetchTotalUnit_NP_EM = 0
    Else
        FetchTotalUnit_NP_EM = rs.Fields(0).value
    End If
        
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchTotalJA_NP_EM(nHari As Double, nBulan As Double, nTahun As Double, isEM As Integer) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "Select IsNull(Sum(Z.Menit),0) FROM " & _
             "( SELECT  NOORDER AS  NoWo,  ISNULL(SUM(DATEDIFF(n,START,PAUSE)),0)  AS  MENIT " & _
             "FROM  TEMPPROGRESS WHERE  FLAGP <>  'M'  AND PAUSE <> '' AND  NOORDER IN  ( " & _
             "SELECT  DISTINCT  A.No_Wo FROM  T_INVOICE A INNER  JOIN  JOBDTA B ON  A.NO_WO =  B.NO_WO " & _
             "WHERE  Day(A.InvDate)  =  " & nHari & "  AND  MONTH(A.INVDATE)  =  " & nBulan & " AND  YEAR(A.INVDATE)  =  " & nTahun & " And " & _
             " " & IIf(isEM = 1, " B.Code = 1 ", " B.Kd_K = 1 ") & " and  A.FS_RC not  in  ('','R001')  and  A.FS_LOG_DEL =  '') " & _
             "GROUP  BY  NOORDER) Z"
    
  
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount = 0 Then
        FetchTotalJA_NP_EM = 0
    Else
        FetchTotalJA_NP_EM = rs.Fields(0).value
    End If
        
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchTotalJF_NP_EM(nHari As Double, nBulan As Double, nTahun As Double, isEM As Integer) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    If isEM = 1 Then
    strSql = "select IsNull((Q.RTE_1 + W.RTE_2),0) FROM " & _
             "(Select sum(A.RTE) as RTE_2, 'OKE' as OKE  FROM " & _
             "(Select N.RTE FROm Hist_Jobadd N Inner Join JobDTA J On N.No_WO = J.No_WO Where N.FS_LOG_DEL = '' and " & _
             "J.FS_LOG_DEL = '' and J.Code = 1 and J.FS_RC not in ('','R001') And " & _
             "No_Inv in (select No_Inv from T_invoice Where FS_Log_Del ='' And Day(InvDate) = " & nHari & " and Month(InvDate) = " & nBulan & " and Year(InvDate) = " & nTahun & ")) A) W " & _
             "Inner Join " & _
             "(select (X.RTE1a + X.RTE2a + X.RTE3a + X.RTE4a + X.RTE5a + X.RTE6a + X.RTE7a + X.RTE8a) as RTE_1, 'OKE' as OKE FROM " & _
             "(Select sum(A.RTE1) as RTE1a, sum(A.RTE2) as RTE2a,sum(A.RTE3) as RTE3a,sum(A.RTE4) as RTE4a, " & _
             "sum(A.RTE5) as RTE5a, sum(A.RTE6) as RTE6a,sum(A.RTE7) as RTE7a,sum(A.RTE8) as RTE8a FROM " & _
             "(Select No_WO, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8 FROm T_INVOICE Where FS_Log_Del ='' and " & _
             "right(NOINV,1) <> 'B' And Day(InvDate) = " & nHari & " and Month(InvDate) = " & nBulan & " and Year(InvDate) = " & nTahun & ") A Inner Join JobDTA j On A.No_WO =J.NO_WO " & _
             "Where J.FS_LOG_DEL = '' and J.Code = 1 and J.FS_RC not in ('','R001')) X) Q On Q.OKE = W.OKE "
    Else
    strSql = "select isNull((Q.RTE_1 + W.RTE_2),0) FROM " & _
             "(Select sum(A.RTE) as RTE_2, 'OKE' as OKE  FROM " & _
             "(Select N.RTE FROm Hist_Jobadd N Inner Join JobDTA J On N.No_WO = J.No_WO Where N.FS_LOG_DEL = '' and " & _
             "J.FS_LOG_DEL = '' and J.Kd_K = 1 and J.FS_RC not in ('','R001') And " & _
             "No_Inv in (select No_Inv from T_invoice Where FS_Log_Del ='' And Day(InvDate) = " & nHari & " and Month(InvDate) = " & nBulan & " and Year(InvDate) = " & nTahun & ")) A) W " & _
             "Inner Join " & _
             "(select (X.RTE1a + X.RTE2a + X.RTE3a + X.RTE4a + X.RTE5a + X.RTE6a + X.RTE7a + X.RTE8a) as RTE_1, 'OKE' as OKE FROM " & _
             "(Select sum(A.RTE1) as RTE1a, sum(A.RTE2) as RTE2a,sum(A.RTE3) as RTE3a,sum(A.RTE4) as RTE4a, " & _
             "sum(A.RTE5) as RTE5a, sum(A.RTE6) as RTE6a,sum(A.RTE7) as RTE7a,sum(A.RTE8) as RTE8a FROM " & _
             "(Select No_WO, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8 FROm T_INVOICE Where FS_Log_Del ='' and " & _
             "right(NOINV,1) <> 'B' And Day(InvDate) = " & nHari & " and Month(InvDate) = " & nBulan & " and Year(InvDate) = " & nTahun & ") A Inner Join JobDTA j On A.No_WO =J.NO_WO " & _
             "Where J.FS_LOG_DEL = '' and J.Kd_K = 1 and J.FS_RC not in ('','R001')) X) Q On Q.OKE = W.OKE "
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount = 0 Then
        FetchTotalJF_NP_EM = 0
    Else
        FetchTotalJF_NP_EM = rs.Fields(0).value
    End If
        
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_Sumary_NP_EM(nHari As Double, nBulan As Double, nTahun As Double, isEM As Integer, nTot As Double, nJA As Double, nJF As Double, nRPS As Double, nLTr As Double, nOPS As Double, nDay As Double, nStEM As Double, nStNP As Double, nStGR As Double)
   On Error GoTo errHandler
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
      
      If isEM = 0 Then
        strSql = "Insert Into LapNP_HO values (" & nTahun & "," & nBulan & ", " & nHari & ", " & nStNP & ", " & nStGR & "," & RepRp(nTot) & "," & RepRp(nJA) & "," & RepRp(nJF) & "," & RepRp(nRPS) & "," & RepRp(nLTr) & "," & RepRp(nOPS) & "," & RepRp(nDay) & ")"
      Else
        strSql = "Insert Into LapEM_HO values (" & nTahun & "," & nBulan & ", " & nHari & ", " & nStEM & ", " & nStGR & "," & RepRp(nTot) & "," & RepRp(nJA) & "," & RepRp(nJF) & "," & RepRp(nRPS) & "," & RepRp(nLTr) & "," & RepRp(nOPS) & "," & RepRp(nDay) & ")"
      End If
      
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
      If Not cm Is Nothing Then Set cm = Nothing
err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub Update_Sumary_NP_EM(nHari As Double, nBulan As Double, nTahun As Double, isEM As Integer, nTot As Double, nJA As Double, nJF As Double, nRPS As Double, nLTr As Double, nOPS As Double, nDay As Double, nStEM As Double, nStNP As Double, nStGR As Double)
   On Error GoTo errHandler
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
           
      If isEM = 0 Then
        strSql = "Update LapNP_HO set StallNP=" & nStNP & ", StallGR=" & nStGR & ",Tot=" & RepRp(nTot) & ",JA=" & RepRp(nJA) & ",JF=" & RepRp(nJF) & ",RPS=" & RepRp(nRPS) & ",LTR=" & RepRp(nLTr) & ",OPS=" & RepRp(nOPS) & ",JAM=" & RepRp(nDay) & " Where Tahun=" & nTahun & " AND bulan=" & nBulan & " AND Tanggal=" & nHari & ""
      Else
        strSql = "Update LapEM_HO set StallEM=" & nStEM & ", StallGR=" & nStGR & ",Tot=" & RepRp(nTot) & ",JA=" & RepRp(nJA) & ",JF=" & RepRp(nJF) & ",RPS=" & RepRp(nRPS) & ",LTR=" & RepRp(nLTr) & ",OPS=" & RepRp(nOPS) & ",JAM=" & RepRp(nDay) & " Where Tahun=" & nTahun & " AND bulan=" & nBulan & " AND Tanggal=" & nHari & ""
      End If
    
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
      If Not cm Is Nothing Then Set cm = Nothing
err.Raise err.Number, err.Source, err.Description
End Sub

Public Function FetchDiscNRM() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "Select NRM, NON from DiscNRM WITH (NOLOCK)"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDiscNRM = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub Update_DiscNRM(nNRM As Double, nNON As Double)
   On Error GoTo errHandler
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
           
        strSql = "Update DiscNRM Set NRM = " & nNRM & ", NON = " & nNON
    
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
      If Not cm Is Nothing Then Set cm = Nothing
err.Raise err.Number, err.Source, err.Description
End Sub

