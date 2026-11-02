VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormDealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Public Function searchNoShow(minuteNoShow As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM BOOK WITH(NOLOCK) WHERE [DATE] + [TIME] <= DateAdd(n, (" & minuteNoShow & " * -1), getdate()) AND DELIV <> 'BTL' AND DELIV <> 'OKE' AND NOFAK NOT IN (SELECT NOBOK FROM " & Log.LogQueue & "..DATABOOKING WITH(NOLOCK) WHERE AKTIV <> 0)"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchNoShow = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Sub canceledNoShow_book(minuteNoShow As Integer, noBook As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    'strSql = "UPDATE BOOK SET DELIV = 'BTL', ALASAN = 'NO SHOW' WHERE [DATE] <= DateAdd(n, (" & minuteNoShow & " * -1), getdate()) AND [TIME] <= convert(varchar(5), DateAdd(n, (" & minuteNoShow & " * -1), getdate()),8) AND NOFAK = '" & Trim(noBook) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'"
    strSql = "UPDATE BOOK SET DELIV = 'BTL', ALASAN = 'NO SHOW' WHERE [DATE] + [TIME] <= DateAdd(n, (" & minuteNoShow & " * -1), getdate()) AND NOFAK = '" & Trim(noBook) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'"
             
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

Public Sub canceledNoShow_Stall(minuteNoShow As Integer, noBook As String)
On Error GoTo errHandler

    Call canceledNoShow_book(minuteNoShow, noBook)
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "UPDATE STALL SET FLAG = 'X' WHERE TANGGAL <= convert(varchar(10), DateAdd(n, (" & minuteNoShow & " * -1), getdate()),1) AND KETERANGAN LIKE '%|" & Trim(noBook) & "|%'"
             
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

Public Sub canceledNoShow(minuteNoShow As Integer, noBook As String)
On Error GoTo errHandler

    Call canceledNoShow_Stall(minuteNoShow, Trim(noBook))
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "INSERT INTO NOTEPART (NOSO, TIPE) VALUES ('" & Trim(noBook) & "', 'BB')"
             
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

Public Function searchRate(jobOrder As String, modelVehicle As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
          
3        ' If Log.RPLHistori = 0 Then
4             strSql = "SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH(NOLOCK) WHERE FS_ITEM = '" & Trim(jobOrder) & "' AND MDL = '" & Trim(modelVehicle) & "'"
5        ' Else
6        '     strSql = "SELECT ISNULL(RATE,0)*60 AS RATE FROM [" & dbReplikasiHO & "]..ACTRATE WITH(NOLOCK) WHERE FS_ITEM = '" & Trim(jobOrder) & "' AND MDL = '" & Trim(modelVehicle) & "'"
7        ' End If
8         Set rs = New ADODB.Recordset
9         rs.CursorLocation = adUseClient
10        rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
11        Set rs.ActiveConnection = Nothing
12        Set searchRate = rs
13        If mConn.State <> adStateClosed Then mConn.Close
14        Exit Function
          
errHandler:
15        err.Raise err.Number, err.Source, err.Description
End Function

Public Sub insertBookQueue(policeNumber As String, tipe As String, customerName As String, dateBook As String, timeBook As String, active As String, bookNumber As String)
1     On Error GoTo errHandler
          Dim cm As New Command
2         If mConn.State = adStateClosed Then mConn.Open (SQLCon)
          
3         strSql = "INSERT INTO " & Log.LogQueue & "..DATABOOKING (NOPOL, TIPE, NAMA, TGLBOK, JAMBOK, AKTIV, NOBOK ) VALUES ('" & Trim(policeNumber) & "','" & Trim(tipe) & "', '" & Trim(customerName) & "', '" & dateBook & "','" & timeBook & "','" & CDbl(active) & "','" & Trim(bookNumber) & "')"
                   
4         cm.CommandText = strSql
5         cm.CommandType = adCmdText
6         Set cm.ActiveConnection = mConn
7         cm.Execute
8         Set cm = Nothing
9         If mConn.State = adStateOpen Then mConn.Close
10        Exit Sub
          
errHandler:
11        If Not cm Is Nothing Then Set cm = Nothing
12        ErrMsg.judul = ""
13        ErrMsg.message = err.Number & vbCrLf & err.Source & vbCrLf & err.Description
14        ErrMsg.show
          
          ''    err.Raise err.Number, err.Source, err.Description
End Sub


Public Sub setQueue(nSTR As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    cm.CommandText = nSTR
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

Public Sub updateBookQueue(policeNumber As String, tipe As String, customerName As String, dateBook As String, timeBook As String, active As String, bookNumber As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "UPDATE " & Log.LogQueue & "..DATABOOKING SET TIPE = '" & Trim(tipe) & "', NAMA = '" & Trim(customerName) & "', TGLBOK = '" & dateBook & "', JAMBOK = '" & timeBook & "', AKTIV = '" & active & "' WHERE NOPOL = '" & Trim(policeNumber) & "' AND NOBOK = '" & Trim(bookNumber) & "'"
             
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

Public Sub deleteBookQueue(policeNumber As String, bookNumber As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "DELETE " & Log.LogQueue & "..DATABOOKING WHERE NOPOL = '" & Trim(policeNumber) & "' AND NOBOK = '" & Trim(bookNumber) & "'"
             
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

Public Function searchMechTime(minuteNoShow As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM BOOK WITH(NOLOCK) WHERE [DATE] <= DateAdd(n, (" & minuteNoShow & " * -1), getdate()) AND [TIME] <= convert(varchar(5), DateAdd(n, (" & minuteNoShow & " * -1), getdate()),8) AND DELIV <> 'BTL' AND DELIV <> 'OKE' AND NOFAK NOT IN (SELECT NOBOK FROM " & Log.LogQueue & "..DATABOOKING WITH(NOLOCK) WHERE AKTIV <> 0)"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchMechTime = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchEvent(dateNow As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    'strSQL = "SELECT * FROM HO_EVENT WHERE DATEDIFF(d,BERLAKU1,'" & dateNow & "') >= 0 AND DATEDIFF(d,'" & dateNow & "',BERLAKU2) >= 0 and AKTIF = 1 ORDER BY noEvent"
    
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT * FROM HO_EVENT A WITH(NOLOCK) INNER JOIN HO_EVTDEALER B WITH(NOLOCK) ON A.NOEVENT = B.NOEVENT " & _
             "WHERE DATEDIFF(d,BERLAKU1,'" & dateNow & "') >= 0 AND DATEDIFF(d,'" & dateNow & "',BERLAKU2) >= 0 and AKTIF = 1 " & _
             "AND B.KDdEALER = '" & Trim(gKdCbg) & "' ORDER BY A.noEvent"
    Else
        
        strSql = "SELECT * FROM [" & dbReplikasiHO & "]..HO_EVENT A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..HO_EVTDEALER B WITH(NOLOCK) ON A.NOEVENT = B.NOEVENT " & _
             "WHERE DATEDIFF(d,BERLAKU1,'" & dateNow & "') >= 0 AND DATEDIFF(d,'" & dateNow & "',BERLAKU2) >= 0 and AKTIF = 1 " & _
             "AND B.KDdEALER = '" & Trim(gKdCbg) & "' ORDER BY A.noEvent"
       ' strSql = "SELECT * FROM [" & dbReplikasiHO & "]..HO_EVENT WITH (NOLOCK) WHERE DATEDIFF(d,BERLAKU1,'" & dateNow & "') >= 0 AND DATEDIFF(d,'" & dateNow & "',BERLAKU2) >= 0 and AKTIF = 1 ORDER BY noEvent"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchEvent = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchModelEvent(noEvent As String) As String
Dim model As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT MDL FROM HO_EVTMODEL WITH(NOLOCK) WHERE noEvent = '" & Trim(noEvent) & "'"
    Else
        strSql = "SELECT MDL FROM [" & dbReplikasiHO & "]..HO_EVTMODEL WITH(NOLOCK) WHERE noEvent = '" & Trim(noEvent) & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    model = ""
    Do While Not rs.EOF
        model = model + IIf(model <> "", " ,", "") + rs("MDL")
        rs.MoveNext
    Loop
    
    searchModelEvent = model
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchKriteriaEvent(criteria As String, noRangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT NO_CHASSIS FROM CUSTDAT A WITH(NOLOCK) WHERE NO_CHASSIS = '" & noRangka & "' " & Trim(criteria)
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchKriteriaEvent = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function eventKodeProgram(eventNo As String) As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT KodeProgram FROM HO_EVENT WITH(NOLOCK) WHERE NoEvent = '" & Replace(Trim(eventNo), vbNewLine, "") & "' "
    Else
        strSql = "SELECT KodeProgram FROM [" & dbReplikasiHO & "]..HO_EVENT WITH(NOLOCK) WHERE NoEvent = '" & Replace(Trim(eventNo), vbNewLine, "") & "' "
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    eventKodeProgram = Trim(rs("KodeProgram"))
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function eventProgram(eventNo As String) As ADODB.Recordset
Dim programNo As String
    On Error GoTo errHandler
    programNo = eventKodeProgram(Replace(Trim(eventNo), vbNewLine, ""))
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT ID_PROGRAM,PROGRAM FROM HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & Trim(programNo) & "'"
    Else
        strSql = "SELECT ID_PROGRAM,PROGRAM FROM [" & dbReplikasiHO & "]..HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & Trim(programNo) & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set eventProgram = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Sub dropTable(tableName As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & tableName & "]')  ) " _
           & " drop table [tempdb]..[" & tableName & "]"
             
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

Public Sub createTableValidasi(tableName As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "CREATE TABLE [tempdb]..[" & tableName & "] ( " _
            & "[NO_DIAG] [VARCHAR] (10) NOT NULL DEFAULT '', " _
            & "[NO_POLISI] [VARCHAR] (12) NOT NULL DEFAULT '', " _
            & "[MODEL] [VARCHAR] (100) NOT NULL DEFAULT ''," _
            & "[NO_CHASSIS] [VARCHAR] (25) NOT NULL DEFAULT '', " _
            & "[DATE] [DATETIME] NOT NULL DEFAULT '01/01/3000', " _
            & "[TGLAPPROVE] [DATETIME] NOT NULL DEFAULT '01/01/3000', " _
            & "[CMPL1] [VARCHAR] (500) NOT NULL DEFAULT '', " _
            & "[NO_WO] [VARCHAR] (10) NOT NULL DEFAULT '', " _
            & "[STATUS_PREDIAG] [VARCHAR] (10) NOT NULL DEFAULT '', " _
            & "[ALASAN] [VARCHAR] (500) NOT NULL DEFAULT '', " _
            & "[NOFAK] [VARCHAR] (10) NOT NULL DEFAULT '')"
            
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

Public Sub createTableRepo(tableName As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "CREATE TABLE [tempdb]..[" & tableName & "] ( " _
         & "[NO_CHASSIS] [varchar] (25) NOT NULL DEFAULT '', " _
         & "[NO_POLISI][varchar](11) NOT NULL DEFAULT '', " _
         & "[FS_KD_PLG][varchar](13) NOT NULL DEFAULT '', " _
         & "[FS_KD_WIL][char](3) NOT NULL DEFAULT '', " _
         & "[CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
         & "[MDL][varchar](25) NOT NULL DEFAULT '', " _
         & "[FMDL][varchar](50) NOT NULL DEFAULT '', " _
         & "[TGLSVC][datetime] NOT NULL DEFAULT '01/01/3000', " _
         & "[NO_WO][char](12) NOT NULL DEFAULT '', " _
         & "[NOINV][char](12) NOT NULL DEFAULT '', " _
         & "[SA][char](5) NOT NULL DEFAULT '', "
        ' & "[TUNGGUSVC][varchar](10) NOT NULL DEFAULT ''," _

          
    strSql = strSql + "[KLPKERJA][char](5) NOT NULL DEFAULT '', " _
         & "[REPS][varchar](50) NOT NULL DEFAULT '', " _
         & "[MECHS][char](5) NOT NULL DEFAULT '', " _
         & "[ACTRATE][DECIMAL](18,2) NOT NULL DEFAULT '0', " _
         & "[FLATRATE][DECIMAL](18,2) NOT NULL DEFAULT '0', " _
         & "[DISC][numeric](9) NOT NULL DEFAULT '0', " _
         & "[OFFS][DECIMAL](18,2) NOT NULL DEFAULT '0', " _
         & "[TPSVC][char](1) NOT NULL DEFAULT '', " _
         & "[THS][char](1) NOT NULL DEFAULT '', " _
         & "[FN_RETURN][char](1) NOT NULL DEFAULT '', " _
         & "[PROSES][char](1) NOT NULL DEFAULT '', " _
         & "[NOFAK][char](8) NOT NULL DEFAULT '', " _
         & "[NOSTALL][varchar](10) NOT NULL DEFAULT '', " _
         & "[FS_KD_DEALER][char](25) NOT NULL DEFAULT '', " _
         & "[FS_NM_DEALER][varchar](100) NOT NULL DEFAULT '', " _
         & "[TIME][varchar](5) NOT NULL DEFAULT '00:00'," _
         & "[TIMESERAH][varchar](5) NOT NULL DEFAULT '00:00') ON [PRIMARY]"
             
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

Public Sub createTableRepoUnitMasuk(tableName As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "CREATE TABLE [tempdb]..[" & tableName & "] ( " _
         & "[NO_CHASSIS] [varchar] (25) NOT NULL DEFAULT '', " _
         & "[NO_POLISI][varchar](11) NOT NULL DEFAULT '', " _
         & "[FS_KD_PLG][varchar](13) NOT NULL DEFAULT '', " _
         & "[FS_KD_WIL][char](3) NOT NULL DEFAULT '', " _
         & "[CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
         & "[MDL][varchar](25) NOT NULL DEFAULT '', " _
         & "[FMDL][varchar](50) NOT NULL DEFAULT '', " _
         & "[TGLSVC][datetime] NOT NULL DEFAULT '01/01/3000', " _
         & "[NO_WO][char](12) NOT NULL DEFAULT '', " _
         & "[NOINV][char](12) NOT NULL DEFAULT '', " _
         & "[TUNGGUSVC][varchar](10) NOT NULL DEFAULT ''," _
         & "[SA][char](5) NOT NULL DEFAULT '', "
         

          
    strSql = strSql + "[KLPKERJA][char](5) NOT NULL DEFAULT '', " _
         & "[REPS][varchar](50) NOT NULL DEFAULT '', " _
         & "[MECHS][char](5) NOT NULL DEFAULT '', " _
         & "[ACTRATE][DECIMAL](18,2) NOT NULL DEFAULT '0', " _
         & "[FLATRATE][DECIMAL](18,2) NOT NULL DEFAULT '0', " _
         & "[DISC][numeric](9) NOT NULL DEFAULT '0', " _
         & "[OFFS][DECIMAL](18,2) NOT NULL DEFAULT '0', " _
         & "[TPSVC][char](1) NOT NULL DEFAULT '', " _
         & "[THS][char](1) NOT NULL DEFAULT '', " _
         & "[FN_RETURN][char](1) NOT NULL DEFAULT '', " _
         & "[PROSES][char](1) NOT NULL DEFAULT '', " _
         & "[NOFAK][char](8) NOT NULL DEFAULT '', " _
         & "[NOSTALL][varchar](10) NOT NULL DEFAULT '', " _
         & "[FS_KD_DEALER][char](25) NOT NULL DEFAULT '', " _
         & "[FS_NM_DEALER][varchar](100) NOT NULL DEFAULT '', " _
         & "[TIME][varchar](5) NOT NULL DEFAULT '00:00'," _
         & "[TIMESERAH][varchar](5) NOT NULL DEFAULT '00:00') ON [PRIMARY]"
             
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

Public Sub createTableInv(tableName As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "CREATE TABLE [dbo].[" & tableName & "] (" _
         & " [CUSTOMER] [varchar] (100) NOT NULL ," _
         & " [CUST_ADRES] [varchar] (250) NOT NULL ," _
         & " [KOTA] [varchar] (30) NOT NULL ," _
         & " [TELPON] [varchar] (30) NOT NULL ," _
         & " [NO_POLISI] [varchar] (12) NOT NULL ," _
         & " [NO_ENGINE] [varchar] (15) NOT NULL ," _
         & " [NO_CHASSIS] [varchar] (25) NOT NULL ," _
         & " [MDL] [varchar] (20) NOT NULL ," _
         & " [MODEL] [varchar] (6) NOT NULL ," _
         & " [FMDL] [varchar] (50) NOT NULL ," _
         & " [COLOR] [varchar] (20) NOT NULL ," _
         & " [FS_KD_DEALER] [varchar] (15) NOT NULL ," _
         & " [FS_NM_DEALER] [varchar] (30) NOT NULL ," _
         & " [DDATE] [varchar] (10) NOT NULL ," _
         & " [CODE] [varchar] (9) NOT NULL ," _
         & " [NPWP] [varchar] (15) NOT NULL ," _
         & " [PPN] [decimal](18, 2) NOT NULL ," _
         & " [DISC] [decimal](18, 2) NOT NULL ,"
 
    strSql = strSql + " [FRONT] [varchar] (4) NOT NULL ," _
         & " [GROUP] [varchar] (4) NOT NULL ," _
         & " [MECH] [varchar] (4) NOT NULL ," _
         & " [KM] [varchar] (8) NOT NULL ," _
         & " [DATE] [datetime] NOT NULL ," _
         & " [TIME] [varchar] (6) NOT NULL ," _
         & " [TIME_M] [varchar] (6) NOT NULL ," _
         & " [TIME_S] [varchar] (6) NOT NULL ," _
         & " [TIME_A] [varchar] (6) NOT NULL ," _
         & " [PART] [decimal](18, 2) NOT NULL ," _
         & " [MTRL] [decimal](18, 2) NOT NULL ," _
         & " [OIL] [decimal](18, 2) NOT NULL ," _
         & " [MET] [decimal](18, 2) NOT NULL ," _
         & " [SBLT] [decimal](18, 2) NOT NULL ," _
         & " [LABOR] [decimal](18, 2) NOT NULL ," _
         & " [BODY] [decimal](18, 2) NOT NULL ," _
         & " [TOTAL] [decimal](18, 2) NOT NULL ," _
         & " [JOB] [varchar] (4) NOT NULL ," _
         & " [REQ1] [varchar] (30) NOT NULL ," _
         & " [REQ2] [varchar] (30) NOT NULL ," _
         & " [REQ3] [varchar] (30) NOT NULL ," _
         & " [REQ4] [varchar] (30) NOT NULL ," _
         & " [REQ5] [varchar] (30) NOT NULL ,"
    
    strSql = strSql + " [REP1] [varchar] (25) NOT NULL ," _
         & " [REP2] [varchar] (25) NOT NULL ," _
         & " [REP3] [varchar] (25) NOT NULL ," _
         & " [REP4] [varchar] (25) NOT NULL ," _
         & " [REP5] [varchar] (25) NOT NULL ," _
         & " [REP6] [varchar] (25) NOT NULL ," _
         & " [REP7] [varchar] (25) NOT NULL ," _
         & " [REP8] [varchar] (25) NOT NULL ," _
         & " [SLIP] [varchar] (30) NOT NULL ," _
         & " [RTE1] [decimal](18, 2) NOT NULL ," _
         & " [RTE2] [decimal](18, 2) NOT NULL ," _
         & " [RTE3] [decimal](18, 2) NOT NULL ," _
         & " [RTE4] [decimal](18, 2) NOT NULL ," _
         & " [RTE5] [decimal](18, 2) NOT NULL ," _
         & " [RTE6] [decimal](18, 2) NOT NULL ," _
         & " [RTE7] [decimal](18, 2) NOT NULL ," _
         & " [RTE8] [decimal](18, 2) NOT NULL ," _
         & " [COND1] [varchar] (8) NOT NULL ," _
         & " [COND2] [varchar] (8) NOT NULL ," _
         & " [COND3] [varchar] (8) NOT NULL ," _
         & " [COND4] [varchar] (8) NOT NULL ," _
         & " [COND5] [varchar] (8) NOT NULL ," _
         & " [COND6] [varchar] (8) NOT NULL ,"
    'NoInovoice +1 digit untuk Retur
    strSql = strSql + " [COND7] [varchar] (8) NOT NULL ," _
         & " [INVDATE] [datetime] NOT NULL ," _
         & " [NO_WO] [varchar] (8) NOT NULL ," _
         & " [NOINV] [varchar] (9) NOT NULL ," _
         & " [STATUS] [varchar] (4) NOT NULL ," _
         & " [PAID] [varchar] (2) NOT NULL ," _
         & " [NOFAK] [varchar] (8) NOT NULL ," _
         & " [REMARK1] [varchar] (20) NOT NULL ," _
         & " [SUGGEST] [varchar] (25) NOT NULL ," _
         & " [DELIV] [varchar] (3) NOT NULL ," _
         & " [SPK] [varchar] (5) NOT NULL ," _
         & " [RPR_COND] [varchar] (1) NOT NULL ," _
         & " [MEKANIK1] [varchar] (4) NOT NULL ," _
         & " [MEKANIK2] [varchar] (4) NOT NULL ," _
         & " [MEKANIK3] [varchar] (4) NOT NULL ," _
         & " [MEKANIK4] [varchar] (4) NOT NULL ," _
         & " [MEKANIK5] [varchar] (4) NOT NULL ," _
         & " [MEKANIK6] [varchar] (4) NOT NULL ," _
         & " [MEKANIK7] [varchar] (4) NOT NULL ," _
         & " [MEKANIK8] [varchar] (4) NOT NULL ," _
         & " [GRP1] [varchar] (4) NOT NULL ,"
     
    strSql = strSql + " [GRP2] [varchar] (4) NOT NULL ," _
         & " [GRP3] [varchar] (4) NOT NULL ," _
         & " [GRP4] [varchar] (4) NOT NULL ," _
         & " [GRP5] [varchar] (4) NOT NULL ," _
         & " [GRP6] [varchar] (4) NOT NULL ," _
         & " [GRP7] [varchar] (4) NOT NULL ," _
         & " [GRP8] [varchar] (4) NOT NULL ," _
         & " [ACTRATE1] [decimal](18, 2) NOT NULL ," _
         & " [ACTRATE2] [decimal](18, 2) NOT NULL ," _
         & " [ACTRATE3] [decimal](18, 2) NOT NULL ," _
         & " [ACTRATE4] [decimal](18, 2) NOT NULL ," _
         & " [ACTRATE5] [decimal](18, 2) NOT NULL ," _
         & " [ACTRATE6] [decimal](18, 2) NOT NULL ," _
         & " [ACTRATE7] [decimal](18, 2) NOT NULL ," _
         & " [ACTRATE8] [decimal](18, 2) NOT NULL ," _
         & " [OFF1] [decimal](18, 2) NOT NULL ," _
         & " [OFF2] [decimal](18, 2) NOT NULL ," _
         & " [OFF3] [decimal](18, 2) NOT NULL ," _
         & " [OFF4] [decimal](18, 2) NOT NULL ," _
         & " [OFF5] [decimal](18, 2) NOT NULL ," _
         & " [OFF6] [decimal](18, 2) NOT NULL ," _
         & " [OFF7] [decimal](18, 2) NOT NULL ,"
       
    strSql = strSql + " [OFF8] [decimal](18, 2) NOT NULL ," _
         & " [KLPKERJA1] [varchar] (4) NOT NULL ," _
         & " [KLPKERJA2] [varchar] (4) NOT NULL ," _
         & " [KLPKERJA3] [varchar] (4) NOT NULL ," _
         & " [KLPKERJA4] [varchar] (4) NOT NULL ," _
         & " [KLPKERJA5] [varchar] (4) NOT NULL ," _
         & " [KLPKERJA6] [varchar] (4) NOT NULL ," _
         & " [KLPKERJA7] [varchar] (4) NOT NULL ," _
         & " [KLPKERJA8] [varchar] (4) NOT NULL ," _
         & " [FS_KREDIT] [varchar] (1) NOT NULL ," _
         & " [FN_PAJAK] [decimal](18, 2) NOT NULL ," _
         & " [FS_TH_MOBIL] [varchar] (4) NOT NULL ," _
         & " [FN_RETURN] [tinyint] NOT NULL ," _
         & " [FS_CATATAN] [varchar] (50) NOT NULL ," _
         & " [FS_LOG_ADD] [varchar] (4) NOT NULL ," _
         & " [FS_LOG_DEL] [varchar] (4) NOT NULL ," _
         & " [FN_JML_ORDER] [decimal](18, 0) NOT NULL ," _
         & " [FN_JML_RATE] [decimal](18, 2) NOT NULL ,"
         
    strSql = strSql + " [ALASAN_INV] [VarChar](150) NOT NULL, " _
         & " [FS_RC] [VarChar](4) NOT NULL, " _
         & " [FS_KD_PLG] [varchar] (13) NOT NULL , " _
         & " [FS_KD_WIL] [varchar] (3) NOT NULL , " _
         & " [FD_TGL_PART] [datetime] NOT NULL , " _
         & " [FS_NO_RETUR] [varchar] (11) NOT NULL , " _
         & " [FD_RETUR] [datetime] NOT NULL , " _
         & " [FS_KD_RETUR] [VarChar](8) NOT NULL " _
         & " ) ON [PRIMARY]"
         
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

Public Function listFieldInTable(tableName As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM " & Log.LogBengkel & ".." & Trim(tableName)
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenDynamic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set listFieldInTable = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function partPrice(kodePart As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    strSql = "SELECT hargajuala FROM inmst WITH(NOLOCK) WHERE kodeItem = '" & kodePart & "'"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set partPrice = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function partPackage(jobTipe As String, jobOrder As String, modelVehicle As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT * FROM ho_partPack WITH(NOLOCK) WHERE job = '" & Trim(jobTipe) & "' AND FS_ITEm = '" & Trim(jobOrder) & "' AND mdl = '" & modelVehicle & "' and FS_LOG_DEL = ''"
    Else
        strSql = "SELECT * FROM [" & dbReplikasiHO & "]..ho_partPack WITH(NOLOCK) WHERE job = '" & Trim(jobTipe) & "' AND FS_ITEm = '" & Trim(jobOrder) & "' AND mdl = '" & modelVehicle & "' and FS_LOG_DEL = ''"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, mConn, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set partPackage = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function partFopPackage(jobTipe As String, jobOrder As String, modelVehicle As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT * FROM HO_FOP_PROGRAM WITH(NOLOCK) WHERE job = '" & Trim(jobTipe) & "' AND FS_ITEm = '" & Trim(jobOrder) & "' AND mdl = '" & modelVehicle & "' and FS_LOG_DEL = ''"
    Else
        strSql = "SELECT * FROM [" & dbReplikasiHO & "]..HO_FOP_PROGRAM WITH(NOLOCK) WHERE FS_ITEm = '" & Trim(jobOrder) & "' AND mdl = '" & modelVehicle & "' and FS_LOG_DEL = ''"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, mConn, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set partFopPackage = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function listMOU(modelVehicle As String, chassisNumber As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    ' ================================================================================================================================
    ' 01 Mar 2017 | Thania | buka kuncian pencarian mou berdasarkan model hanya berdasarkan noka (req. 13 Feb 2017 - Setyo)
    ' ================================================================================================================================
'    strSql = "SELECT A.FLAGDEALER,B.NO_MOU FROM mouhdr A WITH(NOLOCK) INNER JOIN moudtl B WITH(NOLOCK) ON A.no_mou = B.no_mou and A.FS_LOG_DEL = '' " _
'           & "WHERE mdl = '" & modelVehicle & "' AND no_Rangka = '" & chassisNumber & "' ORDER BY mdl, no_Rangka"
    If Log.RPLHistori = 0 Then
        strSql = "SELECT A.FLAGDEALER,B.NO_MOU FROM mouhdr A WITH(NOLOCK) INNER JOIN moudtl B WITH(NOLOCK) ON A.no_mou = B.no_mou and A.FS_LOG_DEL = '' " _
               & "WHERE no_Rangka = '" & chassisNumber & "' ORDER BY mdl, no_Rangka"
    Else
        strSql = "SELECT A.FLAGDEALER,B.NO_MOU FROM [" & dbReplikasiHO & "]..mouhdr A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..moudtl B WITH(NOLOCK) ON A.no_mou = B.no_mou and A.FS_LOG_DEL = '' " _
               & "WHERE no_Rangka = '" & chassisNumber & "' ORDER BY mdl, no_Rangka"

    End If
    ' ================================================================================================================================
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set listMOU = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function listDealerMOU(mouNumber As String, dealerCode As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT NO_MOU FROM MOUCBG WITH(NOLOCK) WHERE NO_MOU = '" & mouNumber & "' AND KODEDEALER = '" & dealerCode & "'"
    Else
        strSql = "SELECT NO_MOU FROM [" & dbReplikasiHO & "]..MOUCBG WITH(NOLOCK) WHERE NO_MOU = '" & mouNumber & "' AND KODEDEALER = '" & dealerCode & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set listDealerMOU = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function countCustomerFU(advisorName As String) As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    countCustomerFU = 0
    strSql = "SELECT [DATE], a.CUSTOMER, a.CUST_ADRES, TELPFU, JAMFU, KETFU, A.MDL + ' | ' + A.FMDL, A.COLOR, A.NO_CHASSIS, A.NO_WO, A.NO_POLISI, KM FROM JOBDTA A WITH(NOLOCK) " _
           & "INNER JOIN JOBDTL B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO INNER JOIN CUSTDAT C WITH(NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS AND A.NO_POLISI = C.NO_POLISI " _
           & "WHERE TGLFU <= CONVERT(VarChar(15), GETDATE(), 1) AND RPR_COND <> '' AND DELIV <> 'BTL' AND FRONT = '" & Trim(advisorName) & "' AND YEAR(TGLFU) > 1900 " _
           & "AND NOT EXISTS (SELECT * FROM DTLFU WITH(NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS AND NO_WO = A.NO_WO AND (SAMBUNG = 'Y' OR COUNTER = 3))  AND A.FS_LOG_DEL = '' ORDER BY TGLFU DESC, JAMFU ASC"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    countCustomerFU = CDbl(rs.RecordCount)
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function countUnitEntry() As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    countUnitEntry = 0
    strSql = "SELECT count(no_wo) AS JUMLAH FROM jobdta WITH(NOLOCK) WHERE date = CONVERT(varchar(10),GETDATE(),1) AND DELIV <> 'BTL'  AND FS_LOG_DEL = ''"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    
    countUnitEntry = CDbl(rs!JUMLAH)
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function countUnitInvoice() As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    countUnitInvoice = 0
    strSql = "SELECT * FROM T_Invoice WITH(NOLOCK) WHERE INVdate = CONVERT(varchar(10),GETDATE(),1) AND FS_LOG_DEL = '' AND NOINV NOT LIKE '%-B'"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    countUnitInvoice = CDbl(rs.RecordCount)
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function countUnitWaitInvoice(id As Integer) As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    countUnitWaitInvoice = 0
    Select Case id
        Case 0
            strSql = "SELECT * FROM JOBDTA WITH(NOLOCK) WHERE (DELIV = 'OKE' OR DELIV = 'PRC') AND RPR_COND = 'S' AND FS_LOG_DEL = '' AND NOINV = '' AND NO_WO IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '')"
        Case 1
            strSql = "SELECT * FROM JOBDTA WITH(NOLOCK) WHERE (DELIV = 'OKE' OR DELIV = 'PRC') AND RPR_COND = 'S' AND FS_LOG_DEL = '' AND NOINV = '' AND NO_WO IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '') AND DATE <= CONVERT(varchar(10),GETDATE(),1)"
    End Select
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    countUnitWaitInvoice = CDbl(rs.RecordCount)
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function


Public Function countUnitWaitInvoiceBySA(id As Integer, sa As String) As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    countUnitWaitInvoiceBySA = 0
    Select Case id
        Case 0
            strSql = "SELECT * FROM JOBDTA WITH(NOLOCK) WHERE (DELIV = 'OKE' OR DELIV = 'PRC') AND RPR_COND = 'S' AND FS_LOG_DEL = '' AND NOINV = '' AND NO_WO IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '') AND FRONT = '" & sa & "'"
        Case 1
            strSql = "SELECT * FROM JOBDTA WITH(NOLOCK) WHERE (DELIV = 'OKE' OR DELIV = 'PRC') AND RPR_COND = 'S' AND FS_LOG_DEL = '' AND NOINV = '' AND NO_WO IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '') AND DATE <= CONVERT(varchar(10),GETDATE(),1) AND FRONT = '" & sa & "'"
    End Select
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    countUnitWaitInvoiceBySA = CDbl(rs.RecordCount)
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function


Public Function countUnitWaitClosing(id As Integer) As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    countUnitWaitClosing = 0
    Select Case id
        Case 0
            strSql = "SELECT * FROM JOBDTA WITH(NOLOCK) WHERE (DELIV = 'OKE' OR DELIV = 'PRC') AND RPR_COND <> 'K' AND NOINV <> '' AND FS_LOG_DEL = ''"
        Case 1
            strSql = "SELECT * FROM JOBDTA WITH(NOLOCK) WHERE (DELIV = 'OKE' OR DELIV = 'PRC') AND RPR_COND <> 'K' AND NOINV <> '' AND DATE = CONVERT(varchar(10),GETDATE(),1) AND FS_LOG_DEL = ''"
    End Select
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    countUnitWaitClosing = CDbl(rs.RecordCount)
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function countUnitOnProcess() As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    countUnitOnProcess = 0
    strSql = "SELECT * FROM STALL WITH(NOLOCK) WHERE TANGGAL = CONVERT(varchar(10),GETDATE(),1) AND FLAG = 'S'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    countUnitOnProcess = CDbl(rs.RecordCount)
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function countUnitWaitProcess() As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    countUnitWaitProcess = 0
    strSql = "SELECT * FROM STALL WITH(NOLOCK) WHERE TANGGAL = CONVERT(varchar(10),GETDATE(),1) AND FLAG = 'N'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    countUnitWaitProcess = CDbl(rs.RecordCount)
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function countUnitOnStopped(id As Integer) As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    countUnitOnStopped = 0
    Select Case id
        Case 0
            strSql = "SELECT * FROM STALL WITH(NOLOCK) WHERE TANGGAL <= CONVERT(varchar(10),GETDATE(),1) AND FLAG = 'P'"
        Case 1
            strSql = "SELECT * FROM STALL WITH(NOLOCK) WHERE TANGGAL = CONVERT(varchar(10),GETDATE(),1) AND FLAG = 'P'"
    End Select
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    countUnitOnStopped = CDbl(rs.RecordCount)
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function kopOrder(id As Integer) As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    kopOrder = ""
    Select Case id
        Case 0
            ' all for wo
            strSql = "SELECT top 1 FS_NM_PERUSH+', '+FS_ALM_PERUSH+' '+FS_KOTA+' Telp. '+FS_TELP AS KOP FROM T_PROFIL WITH(NOLOCK) "
        Case 1
            ' nama PT
            strSql = "SELECT top 1 FS_NM_PERUSH AS KOP FROM T_PROFIL WITH(NOLOCK) "
        Case 2
            ' ALMAT PT
            strSql = "SELECT top 1 FS_ALM_PERUSH + ' ' +FS_KOTA AS KOP FROM T_PROFIL WITH(NOLOCK) "
        Case 3
            ' tELP
            strSql = "SELECT top 1 ' Telp. '+FS_TELP AS KOP FROM T_PROFIL WITH(NOLOCK) "
        Case 4
            ' alamat pt jalan
            strSql = "SELECT top 1 FS_ALM_PERUSH AS KOP FROM T_PROFIL WITH(NOLOCK) "
        Case 5
            ' alamat pt kota
            strSql = "SELECT top 1 FS_KOTA AS KOP FROM T_PROFIL WITH(NOLOCK) "
        Case 6
            ' TELP HOTLINE
            strSql = "SELECT top 1 ' Hotline. '+FS_HOTLINE AS KOP FROM T_PROFIL WITH(NOLOCK) "
    End Select
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    kopOrder = Trim(rs("KOP"))
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function kopBank(id As Integer) As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    kopBank = ""
    Select Case id
        Case 0
            ' all for wo
            strSql = "SELECT top 1 FS_BANK+' '+FS_ACC AS KOP FROM T_PROFIL WITH(NOLOCK) "
    End Select
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    kopBank = Trim(rs("KOP"))
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

