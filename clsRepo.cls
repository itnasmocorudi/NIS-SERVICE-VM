VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Public Function searchWorkTech(tglAwal As String, tglAkhir As String, ketrg As String, Txt As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
            
'    strSql = "Select X.GRP, X.JOB, S.* FROM(Select Z.GRP as GRP, Z. Job as JOB, Z.NoWo as xNOWO FROM (SELECT KLPKERJA1 AS GRP, REP1 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) " _
'            & "UNION SELECT KLPKERJA2 AS GRP, REP2 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) UNION SELECT KLPKERJA3 AS GRP, REP3 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) " _
'            & "UNION SELECT KLPKERJA4 AS GRP, REP4 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) UNION SELECT KLPKERJA5 AS GRP, REP5 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) " _
'            & "UNION SELECT KLPKERJA6 AS GRP, REP6 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) UNION SELECT KLPKERJA7 AS GRP, REP7 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) " _
'            & "UNION SELECT KLPKERJA8 AS GRP, REP8 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) UNION SELECT KLPKERJA  AS GRP, REP  AS JOB, NO_WO as NOWO FROM JOBADD2 WITH(NOLOCK) " _
'            & "UNION SELECT 'SUB', NAMABARANG+' '+KODESATUAN ,NO_WO as NOWO FROM JOBSUBLET WITH(NOLOCK)) Z) X INNER JOIN ( " _
'            & "SELECT A.NoOrder as NOORDER, A.Tanggal as TANGGAL, A.NoStall as NOSTALL, A.Start as START, A.FLAGP as FLAGP , A.Pause as PAUSE, A.Mech as MECH, " _
'            & "GRP AS FOREMAN, DATEDIFF(n,START,CASE WHEN PAUSE = '' THEN '" & Format(isnow, "HH:NN") & "' ELSE PAUSE END) AS JAMKERJA, NO_POLISI, " _
'            & "NO_CHASSIS AS NO_RANGKA, CASE WHEN FN_RETURN = 1 THEN (RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7+RTE8) * -1 ELSE RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7+RTE8 END AS RATE1, " _
'            & "ISNULL((SELECT CASE WHEN FN_RETURN = 1 THEN (ISNULL(SUM(RTE),0)) * -1 ELSE ISNULL(SUM(RTE),0) END AS TOTALRATE FROM JOBADD2 G  WITH(NOLOCK) INNER JOIN JOBDTA H WITH(NOLOCK) ON G.NO_WO = H.NO_WO WHERE G.NO_WO = A.NOORDER GROUP BY G.NO_WO, FN_RETURN),0) AS RATE2, ISNULL(B.NAME,'') AS NAME FROM TEMPPROGRESS A WITH(NOLOCK) LEFT JOIN JOBDTA J WITH(NOLOCK) ON A.NOORDER = J.NO_WO LEFT JOIN MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH WHERE " _
'            & "TANGGAL BETWEEN '" & TglAwal & "' AND '" & TglAkhir & "' AND A.FLAGP <> 'M' " & ketrg & ") S ON X.XNOWO = S.NOORDER AND X.GRP <> '' " & Trim(txt) & " ORDER BY TANGGAL, FOREMAN, MECH, S.NOORDER, START, SUBSTRING(X.GRP,2,1)"
    
    ''2016-09-15 !! penambahan flag code / EM
    
    strSql = "Select X.GRP, X.JOB, S.* FROM(Select Z.GRP as GRP, Z. Job as JOB, Z.NoWo as xNOWO FROM (SELECT KLPKERJA1 AS GRP, REP1 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) " _
            & "UNION SELECT KLPKERJA2 AS GRP, REP2 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) UNION SELECT KLPKERJA3 AS GRP, REP3 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) " _
            & "UNION SELECT KLPKERJA4 AS GRP, REP4 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) UNION SELECT KLPKERJA5 AS GRP, REP5 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) " _
            & "UNION SELECT KLPKERJA6 AS GRP, REP6 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) UNION SELECT KLPKERJA7 AS GRP, REP7 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) " _
            & "UNION SELECT KLPKERJA8 AS GRP, REP8 AS JOB, NO_WO as NOWO FROM JOBDTA WITH(NOLOCK) UNION SELECT KLPKERJA  AS GRP, REP  AS JOB, NO_WO as NOWO FROM JOBADD2 WITH(NOLOCK) " _
            & "UNION SELECT 'SUB', NAMABARANG+' '+KODESATUAN ,NO_WO as NOWO FROM JOBSUBLET WITH(NOLOCK)) Z) X INNER JOIN ( " _
            & "SELECT j.code, A.NoOrder as NOORDER, A.Tanggal as TANGGAL, A.NoStall as NOSTALL, A.Start as START, A.FLAGP as FLAGP , A.Pause as PAUSE, A.Mech as MECH, " _
            & "GRP AS FOREMAN, DATEDIFF(n,START,CASE WHEN PAUSE = '' THEN '" & Format(isnow, "HH:NN") & "' ELSE PAUSE END) AS JAMKERJA, NO_POLISI, " _
            & "NO_CHASSIS AS NO_RANGKA, CASE WHEN FN_RETURN = 1 THEN (RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7+RTE8) * -1 ELSE RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7+RTE8 END AS RATE1, " _
            & "ISNULL((SELECT CASE WHEN FN_RETURN = 1 THEN (ISNULL(SUM(RTE),0)) * -1 ELSE ISNULL(SUM(RTE),0) END AS TOTALRATE FROM JOBADD2 G  WITH(NOLOCK) INNER JOIN JOBDTA H WITH(NOLOCK) ON G.NO_WO = H.NO_WO WHERE G.NO_WO = A.NOORDER GROUP BY G.NO_WO, FN_RETURN),0) AS RATE2, ISNULL(B.NAME,'') AS NAME FROM TEMPPROGRESS A WITH(NOLOCK) LEFT JOIN JOBDTA J WITH(NOLOCK) ON A.NOORDER = J.NO_WO LEFT JOIN MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH WHERE " _
            & "TANGGAL BETWEEN '" & tglAwal & "' AND '" & tglAkhir & "' AND A.FLAGP <> 'M' " & ketrg & ") S ON X.XNOWO = S.NOORDER AND X.GRP <> '' " & Trim(Txt) & " ORDER BY TANGGAL, FOREMAN, MECH, S.NOORDER, START, SUBSTRING(X.GRP,2,1)"
                            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchWorkTech = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
Exit Function
End Function

Public Function printRepo(namaTabel As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT NO, NOWO, TANGGAL, NOSTALL, START, STATUS, SELESAI, MECH, FOREMAN, MENIT, " & _
             "NOPOLISI, NORANGKA, RATE1, RATE2, [GROUP], JOB, REASON, " & _
             "CASE WHEN RATE1+RATE2 > 0 THEN (SELECT ISNULL(SUM(LABOR),0) FROM T_INVOICE WITH(NOLOCK) WHERE NO_WO = A.NOWO AND FS_LOG_DEL = '') ELSE 0 END AS Nominal " & _
             "FROM [tempdb]..[" & namaTabel & "] A"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set printRepo = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchGLtrnNonWo() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    strSql = "SELECT top 1 NOBUKTI FROM GLTRN WITH(NOLOCK) WHERE ISNULL(NOWO,'') = '' OR ISNULL(NOFPS,'') = ''"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchGLtrnNonWo = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
'err.Raise err.Number, err.Source, err.Description
Exit Function
End Function

Public Function searchSyaratBayar(kdPelanggan As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    If OnTopSV = 0 Then
        strSql = "SELECT SYARATBAYAR FROM ARMST WITH(NOLOCK) WHERE KODELGN = '" & Trim(kdPelanggan) & "'"
    Else
        'UNTUK FEBUARI
        strSql = "SELECT TOPGRP as SYARATBAYAR FROM ARMST WITH(NOLOCK) WHERE KODELGN = '" & Trim(kdPelanggan) & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchSyaratBayar = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    Exit Function
End Function


Public Function FetchICSK(nThn As Double, nBln As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT PartFillRate,Dissatified,Satified FROm ICSK_HO WITH(NOLOCK) Where Tahun = " & nThn & " and bulan = " & nBln

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchICSK = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function



Public Function FetchDiFollow_FIR(nThn As Double, nBln As Double) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT COUNT(*) AS JUMLAH " & _
             "FROM dbo.JOBDTL AS A WITH(NOLOCK) INNER JOIN dbo.JOBDTA AS B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND B.FS_LOG_DEL = '' " & _
             "INNER JOIN dbo.DTLFU AS C WITH(NOLOCK) ON A.NO_WO = C.NO_WO AND A.NO_CHASSIS = C.NO_CHASSIS " & _
             "WHERE (C.SAMBUNG = 'Y') AND MONTH(A.TGLFU)  = " & nBln & " AND YEAR(A.TGLFU) = " & nThn & " and MONTH(C.TANGGAL)  = " & nBln & " AND YEAR(C.TANGGAL) = " & nThn & " " & _
             "GROUP BY MONTH(A.TGLFU), YEAR(A.TGLFU), MONTH(C.TANGGAL), YEAR(C.TANGGAL)"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount = 0 Then
        FetchDiFollow_FIR = 0
    Else
        FetchDiFollow_FIR = rs.Fields(0).value
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchNoQ_FIR(nThn As Double, nBln As Double, nQ As Integer) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT COUNT(*) AS JUMLAH " & _
             "FROM dbo.JOBDTL AS A WITH(NOLOCK) INNER JOIN dbo.JOBDTA AS B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND B.FS_LOG_DEL = '' " & _
             "INNER JOIN dbo.DTLFU AS C WITH(NOLOCK) ON A.NO_WO = C.NO_WO AND A.NO_CHASSIS = C.NO_CHASSIS "
    Select Case nQ
        Case 1
            strSql = strSql & "WHERE (C.SAMBUNG = 'Y') AND (C.KUALITAS <> 'P') AND MONTH(A.TGLFU) = " & nBln & " AND YEAR(A.TGLFU) = " & nThn & " and MONTH(C.TANGGAL)  = " & nBln & " AND YEAR(C.TANGGAL) = " & nThn & " "
        Case 2
            strSql = strSql & "WHERE (C.SAMBUNG = 'Y') AND (C.KUALITAS = 'P') AND (C.TEPAT <> 'P') AND MONTH(A.TGLFU)  = " & nBln & " AND YEAR(A.TGLFU) = " & nThn & " and MONTH(C.TANGGAL)  = " & nBln & " AND YEAR(C.TANGGAL) = " & nThn & " "
        Case 3
            strSql = strSql & "WHERE (C.SAMBUNG = 'Y') AND (C.KUALITAS = 'P') AND (C.TEPAT = 'P') AND (C.LAMA <> 'P') AND MONTH(A.TGLFU)  = " & nBln & " AND YEAR(A.TGLFU) = " & nThn & " and MONTH(C.TANGGAL)  = " & nBln & " AND YEAR(C.TANGGAL) = " & nThn & " "
    End Select
    strSql = strSql & "GROUP BY MONTH(A.TGLFU), YEAR(A.TGLFU), MONTH(C.TANGGAL), YEAR(C.TANGGAL)"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount = 0 Then
        FetchNoQ_FIR = 0
    Else
        FetchNoQ_FIR = rs.Fields(0).value
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchSLT_ICSK(nThn As Double, nBln As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "Select Sum(X.minutes), (Sum(X.minutes) - (Sum(X.minutes)%60))/60 as Jam,  (Sum(X.minutes)%60) as Menit " & _
             "FROm (Select A.WO, B.Tgl_In, T.InVDate+ T.InvTime as TanggalInv, " & _
             "datediff(n, B.Tgl_In, T.InVDate + T.InvTime) As minutes " & _
             "from ListAntrian A WITH(NOLOCK) Inner Join " & Log.LogQueue & "..Queue B WITH(NOLOCK) On A.IDPRINT = B.[ID] " & _
             "Inner Join T_invoice T WITH(NOLOCK) On A.WO = T.No_WO " & _
             "Inner Join JobDta J WITH(NOLOCK) On T.No_WO = J.No_WO " & _
             "Where month(T.InVDate) = " & nBln & " and Year(T.InVDate) = " & nThn & " and T.fs_log_del = '' and right(T.NoInv,1) <> 'B' and J.[date] = T.InVDate) X"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchSLT_ICSK = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchSLT_SPLIT_ICSK(nThn As Double, nBln As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    

    strSql = "Select Sum(Y.Tunggu), (Sum(Y.Tunggu) - (Sum(Y.Tunggu)%60))/60 as JamTunggu,  (Sum(Y.Tunggu)%60) as MenitTunggu, Sum(Y.Proses), (Sum(Y.Proses) - (Sum(Y.Proses)%60))/60 as JamProses,  (Sum(Y.Proses)%60) as MenitProses, sum(Y.Tunggu) as totTunggu, sum(Y.Proses) as TotProses FROM " & _
             "(Select J.no_wo,datediff(n, B.Tgl_In, B.Tgl_Panggil) + datediff(n, J.[date] + J.[time],K.TanggalStart) + datediff(n, (J.[date] + J.Time_S),T.InVDate+ T.InvTime) + (U.FullProses-G.MENITKerja) as Tunggu,datediff(n, B.Tgl_Panggil, J.[date] + J.[time]) + G.MENITKerja as Proses, " & _
                    "datediff(n, B.Tgl_In, B.Tgl_Panggil) As TungguDilayani,datediff(n, B.Tgl_Panggil, J.[date] + J.[time]) As ProsesWO," & _
                    "datediff(n, J.[date] + J.[time],K.TanggalStart) As TungguStart,datediff(n, K.TanggalStart,(J.[date] + J.Time_S)) As ProsesKerja," & _
                    "datediff(n, (J.[date] + J.Time_S),T.InVDate+ T.InvTime) As TungguInv,G.MENITKerja, U.FullProses, " & _
                    "B.Tgl_In as Queue, B.Tgl_Panggil as Panggil, J.[date] + J.[time] as WO_Cetak,K.TanggalStart as Start,(J.[date] + J.Time_S) as Finish, T.InVDate+ T.InvTime as TanggalInv  from " & _
             "ListAntrian A WITH(NOLOCK) Inner Join " & Log.LogQueue & "..Queue B WITH(NOLOCK) On A.IDPRINT = B.[ID] Inner Join " & _
             "T_invoice T WITH(NOLOCK) On A.WO = T.no_wo Inner Join " & _
             "JobDTA J WITH(NOLOCK) On A.WO = J.no_wo Inner Join " & _
             "(select A.NoOrder, A.Flag, A.Tanggal +  Min(A.Jam) as TanggalStart From(select Distinct  noorder, Flag, Tanggal, Jam from Progress WITH(NOLOCK) Where FLAG = 'S') A  group by A.NoOrder, A.Flag, A.Tanggal) K On A.WO = K.NoOrder Inner Join " & _
             "(SELECT  NOORDER AS  NoWo,  ISNULL(SUM(DATEDIFF(n,START,PAUSE)),0)  AS  MENITKerja FROM  TEMPPROGRESS WITH(NOLOCK) WHERE  FLAGP <>  'M'  AND  NOORDER IN (SELECT  DISTINCT  A.no_wo FROM  T_INVOICE A WITH(NOLOCK) INNER  JOIN  JOBDTA B WITH(NOLOCK) ON  A.no_wo =  B.no_wo " & _
                "WHERE  MONTH(A.INVDATE)  =  " & nBln & " AND  YEAR(A.INVDATE)  =  " & nThn & " and  A.FS_LOG_DEL =  '') GROUP  BY  NOORDER) G On G.NoWo = A.wo Inner Join " & _
             "(Select DateDiff(n,TglMin, TglMax) as FullProses, O.NoOrder FROm " & _
                 "(select Min(Tanggal+Jam) as TglMin, Noorder from Progress WITH(NOLOCK) Where NoOrder in (select No_WO FROm T_Invoice WITH(NOLOCK) where MONTH(INVDATE)= " & nBln & " AND YEAR(INVDATE) = " & nThn & ") Group By NoOrder) O Inner Join " & _
                 "(select Max(Tanggal+Jam) as TglMax, Noorder from Progress WITH(NOLOCK) Where NoOrder in (select No_WO FROm T_Invoice WITH(NOLOCK) where MONTH(INVDATE)= " & nBln & " AND YEAR(INVDATE) = " & nThn & ") Group By NoOrder) P On O.NoOrder = P.NoOrder) U On U.Noorder = A.WO " & _
                 "Where month(T.InVDate) = " & nBln & " and Year(T.InVDate) = " & nThn & " and T.fs_log_del = '' and right(T.NoInv,1) <> 'B' and J.[date] =T.InVDate) Y"
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchSLT_SPLIT_ICSK = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchOTD_Notif_ICSK(nThn As Double, nBln As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT  SUM(CASE WHEN datediff(n,CONVERT(datetime, TglPanggil), dateadd(n, 0, CONVERT(datetime, DATESERAH + ' ' + TIMESERAH))) >= 0 " & _
                      "AND datediff(n, CONVERT(datetime, tglfinish + ' ' + jamfinish), dateadd(n, 0, CONVERT(datetime, DATESERAH + ' ' + TIMESERAH))) <= 30 " & _
                      "THEN 1 ELSE 0 END) AS jumlah " & _
             "FROM dbo.LISTANTRIAN AS A WITH(NOLOCK) INNER JOIN dbo.T_INVOICE AS D WITH(NOLOCK) ON A.WO = D.NO_WO AND D.FS_LOG_DEL = '' AND D.NOINV NOT LIKE '%-B' INNER JOIN " & _
                      Log.LogQueue & "..Queue AS B WITH(NOLOCK) ON A.IDPRINT = B.ID INNER JOIN " & Log.LogQueue & "..Panggilan AS C WITH(NOLOCK) ON A.IDCALL = C.Id LEFT OUTER JOIN " & _
                      Log.LogQueue & "..FinishService AS E WITH(NOLOCK) ON D.NO_POLISI = E.NoPol AND E.TglFinish = D.INVDATE AND E.Panggil <> 0 LEFT OUTER JOIN " & _
                      "dbo.JOBDTL AS F WITH(NOLOCK) ON A.WO = F.NO_WO " & _
             "WHERE (B.Tgl_In <> '01/01/1900') AND (E.TglFinish <> '01/01/1900') AND (D.PDS = 0) AND (D.THS = 0) AND (D.TPSVC <> 'W') AND ((CASE WHEN datediff(n, " & _
                      "CONVERT(datetime, TglPanggil), dateadd(n, 0, CONVERT(datetime, DATESERAH + ' ' + TIMESERAH))) >= 0 AND datediff(n, CONVERT(datetime, " & _
                      "tglfinish + ' ' + jamfinish), dateadd(n, 0, CONVERT(datetime, DATESERAH + ' ' + TIMESERAH))) <= 30 THEN 1 ELSE 0 END) = 1) " & _
                      "and MONTH(D.INVDATE) = " & nBln & " and  YEAR(D.INVDATE) = " & nThn

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchOTD_Notif_ICSK = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchOTD_THS_ICSK(nThn As Double, nBln As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "Select ( " & _
             "SELECT  COUNT(*) AS JUMLAH FROM T_INVOICE WITH(NOLOCK) WHERE (THS = 1) AND (RIGHT(NOINV, 2) <> '-B') " & _
             "and MONTH(INVDATE) = " & nBln & " and YEAR(INVDATE) = " & nThn & ") - " & _
             "(SELECT COUNT(*) AS JUMLAHR FROM T_INVOICE WITH(NOLOCK) " & _
             "WHERE (THS = 1) And MONTH(CONVERT(char(10), FD_RETUR, 101)) = " & nBln & " and  YEAR(CONVERT(char(10), FD_RETUR, 101)) = " & nThn & ") as THS"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchOTD_THS_ICSK = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub InsertItemICSK(Tahun As Double, Bulan As Double, FIR As Double, LeadTimeH As Double, LeadTimeM As Double, StagnasiH As Double, StagnasiM As Double, ProcessH As Double, ProcessM As Double, PartFillRate As Double, OTD As Double, Dissatified As Double, Satified As Double)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "INSERT INTO ICSK_HO Values (" & Tahun & "," & Bulan & "," & RepRp(FIR) & "," & LeadTimeH & "," & LeadTimeM & "," & StagnasiH & "," & StagnasiM & "," & ProcessH & "," & ProcessM & "," & RepRp(PartFillRate) & "," & RepRp(OTD) & "," & RepRp(Dissatified) & "," & RepRp(Satified) & ")"
             
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

Public Sub UpdateItemICSK(Tahun As Double, Bulan As Double, FIR As Double, LeadTimeH As Double, LeadTimeM As Double, StagnasiH As Double, StagnasiM As Double, ProcessH As Double, ProcessM As Double, PartFillRate As Double, OTD As Double, Dissatified As Double, Satified As Double)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "UPDATE ICSK_HO set FIR = " & RepRp(FIR) & ",LeadTimeH= " & LeadTimeH & ",LeadTimeM=" & LeadTimeM & ",StagnasiH=" & StagnasiH & ",StagnasiM=" & StagnasiM & ",ProcessH=" & ProcessH & ",ProcessM=" & ProcessM & ",PartFillRate=" & RepRp(PartFillRate) & ",OTD=" & RepRp(OTD) & ",Dissatified=" & RepRp(Dissatified) & ",Satified=" & RepRp(Satified) & _
             " Where Tahun=" & Tahun & " AND BULAN = " & Bulan
             
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

Public Sub FetchICSK_Bulanan(nThn As Double, nBln As Double)
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT Bulan FROm ICSK_HO WITH(NOLOCK) Where Tahun = " & nThn & " and bulan = " & nBln

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount = 0 Then
        Call InsertItemICSK(nThn, nBln, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Sub
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub UpdateItemICSK_manual(Tahun As Double, Bulan As Double, PartFillRate As Double, Dissatified As Double, Satified As Double)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "UPDATE ICSK_HO set PartFillRate=" & RepRp(PartFillRate) & ",Dissatified=" & RepRp(Dissatified) & ",Satified=" & RepRp(Satified) & _
             " Where Tahun=" & Tahun & " AND BULAN = " & Bulan
             
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

'-------------------------- D Y N A
Public Sub FetchDYNA_Bulanan(nThn As Double, nBln As Double)
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT Bulan FROm DYNA_HO WITH(NOLOCK) Where Tahun = " & nThn & " and bulan = " & nBln

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount = 0 Then
        Call InsertItemDYNA(nThn, nBln, 0, 0, 0, 0, 0, 0, 0)
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Sub
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub InsertItemDYNA(Tahun As Double, Bulan As Double, SPK As Double, SB1 As Double, SB5 As Double, SBE As Double, grp As Double, SVCD As Double, SVCV As Double)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "INSERT INTO DYNA_HO Values (" & Tahun & "," & Bulan & "," & SPK & "," & SB1 & "," & SB5 & "," & SBE & "," & grp & "," & SVCD & "," & SVCV & ")"
             
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

Public Sub UpdateItemDYNA(Tahun As Double, Bulan As Double, SPK As Double, SB1 As Double, SB5 As Double, SBE As Double, grp As Double, SVCD As Double, SVCV As Double)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "UPDATE DYNA_HO set SPK = " & SPK & ",SB1= " & SB1 & ",SB5=" & SB5 & ",SBE=" & SBE & ",GRP=" & grp & ",SVCD=" & SVCD & ",SVCV=" & SVCV & _
             " Where Tahun=" & Tahun & " AND BULAN = " & Bulan
             
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

Public Function FetchSPK_DYNA(nThn As Double, nBln As Double) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
    strSql = "Select " & _
             "(Select Count(A.NoFaktur) from SIHDR A WITH(NOLOCK) Inner Join SISAL B WITH(NOLOCK) On A.NoFaktur = B.NoFaktur " & _
             "where Month(A.TglFaktur) = " & nBln & " and Year(A.TglFaktur) = " & nThn & " and B.Keterangan like '%DYNA%' And A.TypeTrn = 'A')- " & _
             "(Select Count(A.NoFaktur) from SIHDR A WITH(NOLOCK) Inner Join SISAL B WITH(NOLOCK) On A.NoFaktur = B.NoFaktur " & _
             "where Month(A.TglFaktur) = " & nBln & " and Year(A.TglFaktur) = " & nThn & " and B.Keterangan like '%DYNA%' And A.TypeTrn = 'K')"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchSPK_DYNA = rs.Fields(0).value
    Else
        FetchSPK_DYNA = 0
    End If
    
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchSB_DYNA(nThn As Double, nBln As Double, nR As String) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    If Log.RPLHistori = 0 Then
        strSql = "select " & _
             "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
             "Where FS_RC = '" & nR & "' and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
             "And MDl in (Select MDL from Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
             "And FS_Log_DEL = '' AND PDS = 0 And Right(NoInv,1) <> 'B') - " & _
             "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
             "Where FS_RC = '" & nR & "' and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
             "And MDl in (Select MDL from Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
             "And FS_Log_DEL <> '' AND PDS = 0 And Right(NoInv,1) <> 'B')"
    Else
        strSql = "select " & _
             "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
             "Where FS_RC = '" & nR & "' and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
             "And MDl in (Select MDL from [" & dbReplikasiHO & "]..Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
             "And FS_Log_DEL = '' AND PDS = 0 And Right(NoInv,1) <> 'B') - " & _
             "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
             "Where FS_RC = '" & nR & "' and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
             "And MDl in (Select MDL from [" & dbReplikasiHO & "]..Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
             "And FS_Log_DEL <> '' AND PDS = 0 And Right(NoInv,1) <> 'B')"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchSB_DYNA = rs.Fields(0).value
    Else
        FetchSB_DYNA = 0
    End If
    
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchSBE_DYNA(nThn As Double, nBln As Double) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    If Log.RPLHistori = 0 Then
        strSql = "select " & _
                 "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
                 "Where FS_RC not in ('R001', 'R005','') and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
                 "And MDl in (Select MDL from Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
                 "And FS_Log_DEL = '' AND PDS = 0 And Right(NoInv,1) <> 'B') - " & _
                 "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
                 "Where FS_RC not in ('R001', 'R005','') and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
                 "And MDl in (Select MDL from Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
                 "And FS_Log_DEL <> '' AND PDS = 0 And Right(NoInv,1) <> 'B')"
    Else
        strSql = "select " & _
             "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
             "Where FS_RC not in ('R001', 'R005','') and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
             "And MDl in (Select MDL from [" & dbReplikasiHO & "]..Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
             "And FS_Log_DEL = '' AND PDS = 0 And Right(NoInv,1) <> 'B') - " & _
             "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
             "Where FS_RC not in ('R001', 'R005','') and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
             "And MDl in (Select MDL from [" & dbReplikasiHO & "]..Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
             "And FS_Log_DEL <> '' AND PDS = 0 And Right(NoInv,1) <> 'B')"
    
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchSBE_DYNA = rs.Fields(0).value
    Else
        FetchSBE_DYNA = 0
    End If
    
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchTHS_DYNA(nThn As Double, nBln As Double) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    
    If Log.RPLHistori = 0 Then
        strSql = "select " & _
                 "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
                 "Where THS = 1 and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
                 "And MDl in (Select MDL from Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
                 "And FS_Log_DEL = '' AND PDS = 0 And Right(NoInv,1) <> 'B') - " & _
                 "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
                 "Where THS = 1 and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
                 "And MDl in (Select MDL from Model Where [desc] like '%DYNA%') " & _
                 "And FS_Log_DEL <> '' AND PDS = 0 And Right(NoInv,1) <> 'B')"
    Else
        strSql = "select " & _
             "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
             "Where THS = 1 and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
             "And MDl in (Select MDL from [" & dbReplikasiHO & "]..Model WITH(NOLOCK) Where [desc] like '%DYNA%') " & _
             "And FS_Log_DEL = '' AND PDS = 0 And Right(NoInv,1) <> 'B') - " & _
             "(Select Count(NoInv) from T_invoice WITH(NOLOCK) " & _
             "Where THS = 1 and Month(InvDate) = " & nBln & " and Year(InvDate) = " & nThn & " " & _
             "And MDl in (Select MDL from [" & dbReplikasiHO & "]..Model Where [desc] like '%DYNA%') " & _
             "And FS_Log_DEL <> '' AND PDS = 0 And Right(NoInv,1) <> 'B')"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchTHS_DYNA = rs.Fields(0).value
    Else
        FetchTHS_DYNA = 0
    End If
    
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchLapPartLama(nThn As Double, nBln As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (Replace(SQLConACC, Log.LogACC, Left(Log.LogACC, 6) & nThn & IIf(Len(Trim(nBln)) = 1, "0" & nBln, nBln)))
    
    strSql = "SELECT CASE WHEN ISNULL(P.TANGGAL,'') = '' THEN Q.TanggalFAKTUR ELSE P.TANGGAL END AS Tanggal, " & _
         "TOTC, TOTW, TOTAL,  ISNULL((Q.FSYNT),0) AS FSYNT, ISNULL((Q.SYNT),0) AS SYNT, ISNULL((Q.MINR),0) AS MINR, " & _
         "ISNULL((Q.TGBF),0) AS TGBF, ISNULL((Q.SLLC),0) AS SLLC, ISNULL((Q.TGGO),0) AS TGGO, ISNULL((Q.BATTERY),0) AS BATTERY " & _
         "FROM ( " & _
         "SELECT CASE WHEN ISNULL(K.TANGGAL,'') = '' THEN L.TGLTRN ELSE K.TANGGAL END AS Tanggal, " & _
         "ISNULL(K.TotalC,0) AS 'TotC', L.TotalW AS 'TotW', " & _
         "ISNULL(K.TotalC,0)+L.TotalW as Total FROM " & _
         "( " & _
         "Select A.TGLTRN, Sum(A.TotalW) as TotalW FROM ( " & _
         "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) AS 'TotalW' FROM GLTRN WITH(NOLOCK) " & _
         "WHERE NOBUKTI LIKE 'E-%' AND KODESUMBER = 'SV' AND " & _
         "KODEGL IN (Select SlipE as SlipE from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' " & _
         "       Union Select SlipEr as SlipE from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' ) " & _
         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
         "GROUP BY TGLTRN UNION " & _
         "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) AS 'TotalW' FROM GLTRN WITH(NOLOCK) " & _
         "WHERE NOBUKTI LIKE 'G-%' AND KODESUMBER = 'SI' AND " & _
         "KODEGL IN (Select SlipG as SlipG from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' " & _
         "       Union Select SlipGr as SlipG from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
         "GROUP BY TGLTRN) A GROUP BY A.TGLTRN " & _
         ") L FULL OUTER JOIN " & _
         "( "
         strSql = strSql & "SELECT TGLTRN AS TANGGAL, SUM(ISNULL(PIUTANG,0)) AS 'TotalC' FROM " & _
         "(SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) as piutang FROM GLTRN WITH(NOLOCK) " & _
         "WHERE NOBUKTI LIKE 'A%' AND KODESUMBER = 'SI' AND " & _
         "KODEGL IN (Select SlipA as SlipA from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' " & _
         "       Union Select SlipAr as SlipA from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '') " & _
         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
         "GROUP BY TGLTRN " & _
         "Union " & _
         "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) as piutang FROM GLTRN WITH(NOLOCK) " & _
         "WHERE NOBUKTI LIKE 'B%' AND KODESUMBER = 'SI' AND " & _
         "KODEGL IN (Select SlipB as SlipB from " & gDb & "..SlipCode Where SlipB <> '' " & _
         "       Union Select SlipBr as SlipB from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' ) " & _
         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
         "GROUP BY TGLTRN " & _
         "Union " & _
         "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) as piutang FROM GLTRN WITH(NOLOCK) " & _
         "WHERE NOBUKTI LIKE 'C%' AND KODESUMBER = 'SI' AND " & _
         "KODEGL IN (Select SlipC as SlipC from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' " & _
         "       Union Select SlipCr as SlipC from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '' ) " & _
         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
         "GROUP BY TGLTRN " & _
         ") A " & _
         "GROUP BY A.TGLTRN " & _
         ") K ON K.TANGGAL = L.TGLTRN " & _
         ") P "
         strSql = strSql & "FULL OUTER JOIN " & _
         "( " & _
         "SELECT S.TanggalFAKTUR, " & _
         "ISNULL(SUM(S.FSYNT),0) AS FSYNT, ISNULL(SUM(S.SYNT),0) AS SYNT, ISNULL(SUM(S.MINR),0) AS MINR, ISNULL(SUM(S.TGBF),0) AS TGBF, " & _
         "ISNULL(SUM(S.SLLC),0) AS SLLC, ISNULL(SUM(S.TGGO),0) AS TGGO, ISNULL(SUM(S.BATTERY),0) AS BATTERY " & _
         "FROM ( " & _
         "SELECT CASE WHEN ISNULL(Z.TGLFAKTUR,'') <> '' THEN Z.TGLFAKTUR ELSE CASE WHEN ISNULL(Y.TGLFAKTUR,'') <> '' THEN Y.TGLFAKTUR ELSE CASE WHEN ISNULL(X.TGLFAKTUR,'') <> '' THEN X.TGLFAKTUR ELSE CASE WHEN ISNULL(W.TGLFAKTUR,'') <> '' THEN W.TGLFAKTUR ELSE CASE WHEN ISNULL(V.TGLFAKTUR,'') <> '' THEN V.TGLFAKTUR ELSE CASE WHEN ISNULL(U.TGLFAKTUR,'') <> '' THEN U.TGLFAKTUR ELSE T.TGLFAKTUR END END END END END END AS TanggalFAKTUR, " & _
         "case When ISNULL(Z.KODEITEM,'') <> '' THEN Z.KODEITEM ELSE CASE WHEN ISNULL(Y.KODEITEM,'') <> '' THEN Y.KODEITEM ELSE CASE WHEN ISNULL(X.KODEITEM,'') <> '' THEN X.KODEITEM ELSE CASE WHEN ISNULL(W.KODEITEM,'') <> '' THEN W.KODEITEM ELSE CASE WHEN ISNULL(V.KODEITEM,'') <> '' THEN V.KODEITEM ELSE CASE WHEN ISNULL(U.KODEITEM,'') <> '' THEN U.KODEITEM ELSE T.KODEITEM END END END END END END as KODEITEM," & _
         "ISNULL((Z.FSYNT),0) AS FSYNT, ISNULL((Y.SYNT),0) AS SYNT, ISNULL((X.MINR),0) AS MINR, " & _
         "ISNULL((W.TGBF),0) AS TGBF, ISNULL((V.SLLC),0) AS SLLC, ISNULL((U.TGGO),0) AS TGGO, ISNULL((T.BATTERY),0) AS BATTERY " & _
         "From "
         strSql = strSql & "( " & _
         "    Select X.TGLFAKTUR, X.KODEITEM, X.FSYNT * Y.QTY as FSYNT from " & _
         "    ( " & _
         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS FSYNT " & _
         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B ON A.NOFAKTUR = B.NOFAKTUR " & _
         "           WHERE A.NOFAKTUR IN " & _
         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
         "           where KODEGL IN " & _
         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
         "          ) "
         strSql = strSql & " Union " & _
         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
         "           where G.KODEGL IN " & _
         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
         "          )) " & _
         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'FSYNT')) " & _
         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'FSYNT') " & _
         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
         "      ) Z "
         strSql = strSql & "FULL OUTER JOIN ( " & _
         "    Select X.TGLFAKTUR, X.KODEITEM, X.SYNT * Y.QTY as SYNT from ( " & _
         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS SYNT " & _
         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
         "           WHERE A.NOFAKTUR IN " & _
         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
         "           where KODEGL IN " & _
         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
         "          ) "
         strSql = strSql & " Union " & _
         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
         "           where G.KODEGL IN " & _
         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
         "          )) " & _
         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SYNT')) " & _
         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SYNT') " & _
         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
         "      ) Y ON Z.TGLFAKTUR = Y.TGLFAKTUR And Z.KodeItem = Y.KodeItem "
         strSql = strSql & "FULL OUTER JOIN ( " & _
         "    Select X.TGLFAKTUR, X.KODEITEM, X.MINR * Y.QTY as MINR from ( " & _
         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS MINR " & _
         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B  WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
         "           WHERE A.NOFAKTUR IN " & _
         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
         "           where KODEGL IN " & _
         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
         "          ) "
         strSql = strSql & " Union " & _
         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
         "           where G.KODEGL IN " & _
         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
         "          )) " & _
         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'MINR')) " & _
         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'MINR') " & _
         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
         "      ) X ON Z.TGLFAKTUR = X.TGLFAKTUR AND Z.KodeItem = X.KodeItem  "
         strSql = strSql & "FULL OUTER JOIN ( " & _
         "    Select X.TGLFAKTUR, X.KODEITEM, X.TGBF * Y.QTY as TGBF from ( " & _
         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS TGBF " & _
         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
         "           WHERE A.NOFAKTUR IN " & _
         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
         "           where KODEGL IN " & _
         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
         "          ) "
         strSql = strSql & " Union " & _
         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
         "           where G.KODEGL IN " & _
         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
         "          )) " & _
         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGBF')) " & _
         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGBF') " & _
         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
         "    ) W ON Z.TGLFAKTUR = W.TGLFAKTUR AND Z.KodeItem = W.KodeItem "
         strSql = strSql & "FULL OUTER JOIN ( " & _
         "    Select X.TGLFAKTUR, X.KODEITEM, X.SLLC * Y.QTY as SLLC from ( " & _
         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS SLLC " & _
         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
         "           WHERE A.NOFAKTUR IN " & _
         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
         "           where KODEGL IN " & _
         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
         "          ) "
         strSql = strSql & " Union " & _
         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
         "           where G.KODEGL IN " & _
         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
         "          )) " & _
         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SLLC')) " & _
         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SLLC') " & _
         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem  " & _
         "    ) V ON Z.TGLFAKTUR = V.TGLFAKTUR AND Z.KodeItem = V.KodeItem "
         strSql = strSql & "FULL OUTER JOIN ( " & _
         "    Select X.TGLFAKTUR, X.KODEITEM, X.TGGO * Y.QTY as TGGO from ( " & _
         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS TGGO " & _
         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
         "           WHERE A.NOFAKTUR IN " & _
         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
         "           where KODEGL IN " & _
         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
         "          ) "
         strSql = strSql & " Union " & _
         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
         "           where G.KODEGL IN " & _
         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
         "          )) " & _
         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGGO')) " & _
         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGGO') " & _
         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
         "      ) U ON Z.TGLFAKTUR = U.TGLFAKTUR AND Z.KodeItem = U.KodeItem "
         strSql = strSql & "FULL OUTER JOIN ( " & _
         "    Select X.TGLFAKTUR, X.KODEITEM, X.BATTERY * Y.QTY as BATTERY from ( " & _
         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS BATTERY " & _
         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
         "           WHERE A.NOFAKTUR IN " & _
         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
         "           where KODEGL IN " & _
         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
         "          ) "
         strSql = strSql & " Union " & _
         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
         "           where G.KODEGL IN " & _
         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
         "          )) " & _
         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'BATR' )) " & _
         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'BATR') " & _
         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
         "      ) T ON Z.TGLFAKTUR = T.TGLFAKTUR AND Z.KodeItem = T.KodeItem " & _
         ") S "
         strSql = strSql & " GROUP BY S.TanggalFAKTUR " & _
         ") Q ON P.TANGGAL = Q.TanggalFAKTUR " & _
         "ORDER BY P.TANGGAL, Q.TanggalFAKTUR"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, (Replace(SQLConACC, Log.LogACC, Left(Log.LogACC, 6) & nThn & IIf(Len(Trim(nBln)) = 1, "0" & nBln, nBln))), adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchLapPartLama = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchLapPart(nThn As Double, nBln As Double) As ADODB.Recordset
1         On Error GoTo errHandler
2         If mConn.State = adStateClosed Then mConn.Open (Replace(SQLConACC, Log.LogACC, Left(Log.LogACC, 6) & nThn & IIf(Len(Trim(nBln)) = 1, "0" & nBln, nBln)))
          
      '    strSql = "SELECT CASE WHEN ISNULL(P.TANGGAL,'') = '' THEN Q.TanggalFAKTUR ELSE P.TANGGAL END AS Tanggal, " & _
      '         "TOTC, TOTW, TOTAL,  ISNULL((Q.FSYNT),0) AS FSYNT, ISNULL((Q.SYNT),0) AS SYNT, ISNULL((Q.MINR),0) AS MINR, " & _
      '         "ISNULL((Q.TGBF),0) AS TGBF, ISNULL((Q.SLLC),0) AS SLLC, ISNULL((Q.TGGO),0) AS TGGO, ISNULL((Q.BATTERY),0) AS BATTERY " & _
      '         "FROM ( " & _
      '         "SELECT CASE WHEN ISNULL(K.TANGGAL,'') = '' THEN L.TGLTRN ELSE K.TANGGAL END AS Tanggal, " & _
      '         "ISNULL(K.TotalC,0) AS 'TotC', L.TotalW AS 'TotW', " & _
      '         "ISNULL(K.TotalC,0)+L.TotalW as Total FROM " & _
      '         "( " & _
      '         "Select A.TGLTRN, Sum(A.TotalW) as TotalW FROM ( " & _
      '         "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) AS 'TotalW' FROM GLTRN WITH(NOLOCK) " & _
      '         "WHERE NOBUKTI LIKE 'E-%' AND KODESUMBER = 'SV' AND " & _
      '         "KODEGL IN (Select SlipE as SlipE from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' " & _
      '         "       Union Select SlipEr as SlipE from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' ) " & _
      '         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
      '         "GROUP BY TGLTRN UNION " & _
      '         "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) AS 'TotalW' FROM GLTRN WITH(NOLOCK) " & _
      '         "WHERE NOBUKTI LIKE 'G-%' AND KODESUMBER = 'SI' AND " & _
      '         "KODEGL IN (Select SlipG as SlipG from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' " & _
      '         "       Union Select SlipGr as SlipG from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
      '         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
      '         "GROUP BY TGLTRN) A GROUP BY A.TGLTRN " & _
      '         ") L FULL OUTER JOIN " & _
'         "( "
      '         strSql = strSql & "SELECT TGLTRN AS TANGGAL, SUM(ISNULL(PIUTANG,0)) AS 'TotalC' FROM " & _
      '         "(SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) as piutang FROM GLTRN WITH(NOLOCK) " & _
      '         "WHERE NOBUKTI LIKE 'A%' AND KODESUMBER = 'SI' AND " & _
      '         "KODEGL IN (Select SlipA as SlipA from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' " & _
      '         "       Union Select SlipAr as SlipA from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '') " & _
      '         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
      '         "GROUP BY TGLTRN " & _
      '         "Union " & _
      '         "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) as piutang FROM GLTRN WITH(NOLOCK) " & _
      '         "WHERE NOBUKTI LIKE 'B%' AND KODESUMBER = 'SI' AND " & _
      '         "KODEGL IN (Select SlipB as SlipB from " & gDb & "..SlipCode Where SlipB <> '' " & _
      '         "       Union Select SlipBr as SlipB from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' ) " & _
      '         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
      '         "GROUP BY TGLTRN " & _
      '         "Union " & _
      '         "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) as piutang FROM GLTRN WITH(NOLOCK) " & _
      '         "WHERE NOBUKTI LIKE 'C%' AND KODESUMBER = 'SI' AND " & _
      '         "KODEGL IN (Select SlipC as SlipC from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' " & _
      '         "       Union Select SlipCr as SlipC from " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '' ) " & _
      '         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
      '         "GROUP BY TGLTRN " & _
      '         ") A " & _
      '         "GROUP BY A.TGLTRN " & _
      '         ") K ON K.TANGGAL = L.TGLTRN " & _
'         ") P "
      '         strSql = strSql & "FULL OUTER JOIN " & _
      '         "( " & _
      '         "SELECT S.TanggalFAKTUR, " & _
      '         "ISNULL(SUM(S.FSYNT),0) AS FSYNT, ISNULL(SUM(S.SYNT),0) AS SYNT, ISNULL(SUM(S.MINR),0) AS MINR, ISNULL(SUM(S.TGBF),0) AS TGBF, " & _
      '         "ISNULL(SUM(S.SLLC),0) AS SLLC, ISNULL(SUM(S.TGGO),0) AS TGGO, ISNULL(SUM(S.BATTERY),0) AS BATTERY " & _
      '         "FROM ( " & _
      '         "SELECT CASE WHEN ISNULL(Z.TGLFAKTUR,'') <> '' THEN Z.TGLFAKTUR ELSE CASE WHEN ISNULL(Y.TGLFAKTUR,'') <> '' THEN Y.TGLFAKTUR ELSE CASE WHEN ISNULL(X.TGLFAKTUR,'') <> '' THEN X.TGLFAKTUR ELSE CASE WHEN ISNULL(W.TGLFAKTUR,'') <> '' THEN W.TGLFAKTUR ELSE CASE WHEN ISNULL(V.TGLFAKTUR,'') <> '' THEN V.TGLFAKTUR ELSE CASE WHEN ISNULL(U.TGLFAKTUR,'') <> '' THEN U.TGLFAKTUR ELSE T.TGLFAKTUR END END END END END END AS TanggalFAKTUR, " & _
      '         "case When ISNULL(Z.KODEITEM,'') <> '' THEN Z.KODEITEM ELSE CASE WHEN ISNULL(Y.KODEITEM,'') <> '' THEN Y.KODEITEM ELSE CASE WHEN ISNULL(X.KODEITEM,'') <> '' THEN X.KODEITEM ELSE CASE WHEN ISNULL(W.KODEITEM,'') <> '' THEN W.KODEITEM ELSE CASE WHEN ISNULL(V.KODEITEM,'') <> '' THEN V.KODEITEM ELSE CASE WHEN ISNULL(U.KODEITEM,'') <> '' THEN U.KODEITEM ELSE T.KODEITEM END END END END END END as KODEITEM," & _
      '         "ISNULL((Z.FSYNT),0) AS FSYNT, ISNULL((Y.SYNT),0) AS SYNT, ISNULL((X.MINR),0) AS MINR, " & _
      '         "ISNULL((W.TGBF),0) AS TGBF, ISNULL((V.SLLC),0) AS SLLC, ISNULL((U.TGGO),0) AS TGGO, ISNULL((T.BATTERY),0) AS BATTERY " & _
      '         "From "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
3              strSql = "SELECT CASE WHEN ISNULL(P.TANGGAL,'') = '' THEN Q.TanggalFAKTUR ELSE P.TANGGAL END AS Tanggal, " & _
               "TOTC, TOTW, TOTAL,  ISNULL((Q.FSYNT),0) AS FSYNT, ISNULL((Q.SYNT),0) AS SYNT, ISNULL((Q.MINR),0) AS MINR, " & _
               "ISNULL((Q.TGBF),0) AS TGBF, ISNULL((Q.SLLC),0) AS SLLC, ISNULL((Q.TGGO),0) AS TGGO, ISNULL((Q.BATTERY),0) AS BATTERY " & _
               "FROM ( " & _
               "SELECT CASE WHEN ISNULL(K.TANGGAL,'') = '' THEN L.TGLTRN ELSE K.TANGGAL END AS Tanggal, " & _
               "ISNULL(K.TotalC,0) AS 'TotC', L.TotalW AS 'TotW', " & _
               "ISNULL(K.TotalC,0)+L.TotalW as Total FROM " & _
               "( " & _
               "Select A.TGLTRN, Sum(A.TotalW) as TotalW FROM ( " & _
               "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) AS 'TotalW' FROM GLTRN WITH(NOLOCK) " & _
               "WHERE NOBUKTI LIKE 'E-%' AND KODESUMBER = 'SV' AND " & _
               "KODEGL IN (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02) " & _
               "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
               "GROUP BY TGLTRN UNION "
               
4              strSql = strSql & " SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) AS 'TotalW' FROM GLTRN WITH(NOLOCK) " & _
               "WHERE NOBUKTI LIKE 'G-%' AND KODESUMBER = 'SI' AND " & _
               "KODEGL IN (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02) " & _
               "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
               "GROUP BY TGLTRN) A GROUP BY A.TGLTRN " & _
               ") L FULL OUTER JOIN " & _
               "( "
               
      '         strSql = strSql & "SELECT TGLTRN AS TANGGAL, SUM(ISNULL(PIUTANG,0)) AS 'TotalC' FROM " & _
      '         "(SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) as piutang FROM GLTRN WITH(NOLOCK) " & _
      '         "WHERE NOBUKTI LIKE 'A%' AND KODESUMBER = 'SI' AND " & _
      '         "KODEGL IN (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
      '         "               Union " & _
      '         "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
      '         "               Union " & _
      '         "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02) " & _
      '         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
      '         "GROUP BY TGLTRN " & _
      '         "Union " & _
      '         "SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) as piutang FROM GLTRN WITH(NOLOCK) " & _
      '         "WHERE NOBUKTI LIKE 'B%' AND KODESUMBER = 'SI' AND " & _
      '         "KODEGL IN (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
      '         "               Union " & _
      '         "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
      '         "               Union " & _
      '         "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02) " & _
      '         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
      '         "GROUP BY TGLTRN " & _
      '         "Union "
      '
      '         strSql = strSql & " SELECT TGLTRN, ISNULL(ROUND(SUM((jlhkredit-JlhDebit)) + (SUM((jlhkredit-JlhDebit))*0.1),0),0) as piutang FROM GLTRN WITH(NOLOCK) " & _
      '         "WHERE NOBUKTI LIKE 'C%' AND KODESUMBER = 'SI' AND " & _
      '         "KODEGL IN (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
      '         "               Union " & _
      '         "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
      '         "               Union " & _
      '         "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02) " & _
      '         "AND MONTH(TGLTRN) = " & nBln & " AND YEAR(TGLTRN) = " & nThn & " " & _
      '         "GROUP BY TGLTRN " & _
      '         ") A " & _
      '         "GROUP BY A.TGLTRN " & _
      '         ") K ON K.TANGGAL = L.TGLTRN " & _

5              strSql = strSql & " SELECT TglFaktur TANGGAL, sum(Dpp+PPN) as TotalC " & _
                     " From Sihdr " & _
                     " Where Month(tglFaktur) = " & nBln & " And year(tglFaktur) = " & nThn & _
                     " AND (NoFaktur LIKE 'A%' OR NoFaktur LIKE 'B%' OR NoFaktur LIKE 'C%') " & _
                     " group by TglFaktur " & _
               ") K ON K.TANGGAL = L.TGLTRN " & _
               ") P "
6              strSql = strSql & "FULL OUTER JOIN " & _
               "( " & _
               "SELECT S.TanggalFAKTUR, " & _
               "ISNULL(SUM(S.FSYNT),0) AS FSYNT, ISNULL(SUM(S.SYNT),0) AS SYNT, ISNULL(SUM(S.MINR),0) AS MINR, ISNULL(SUM(S.TGBF),0) AS TGBF, " & _
               "ISNULL(SUM(S.SLLC),0) AS SLLC, ISNULL(SUM(S.TGGO),0) AS TGGO, ISNULL(SUM(S.BATTERY),0) AS BATTERY " & _
               "FROM ( " & _
               "SELECT CASE WHEN ISNULL(Z.TGLFAKTUR,'') <> '' THEN Z.TGLFAKTUR ELSE CASE WHEN ISNULL(Y.TGLFAKTUR,'') <> '' THEN Y.TGLFAKTUR ELSE CASE WHEN ISNULL(X.TGLFAKTUR,'') <> '' THEN X.TGLFAKTUR ELSE CASE WHEN ISNULL(W.TGLFAKTUR,'') <> '' THEN W.TGLFAKTUR ELSE CASE WHEN ISNULL(V.TGLFAKTUR,'') <> '' THEN V.TGLFAKTUR ELSE CASE WHEN ISNULL(U.TGLFAKTUR,'') <> '' THEN U.TGLFAKTUR ELSE T.TGLFAKTUR END END END END END END AS TanggalFAKTUR, " & _
               "case When ISNULL(Z.KODEITEM,'') <> '' THEN Z.KODEITEM ELSE CASE WHEN ISNULL(Y.KODEITEM,'') <> '' THEN Y.KODEITEM ELSE CASE WHEN ISNULL(X.KODEITEM,'') <> '' THEN X.KODEITEM ELSE CASE WHEN ISNULL(W.KODEITEM,'') <> '' THEN W.KODEITEM ELSE CASE WHEN ISNULL(V.KODEITEM,'') <> '' THEN V.KODEITEM ELSE CASE WHEN ISNULL(U.KODEITEM,'') <> '' THEN U.KODEITEM ELSE T.KODEITEM END END END END END END as KODEITEM," & _
               "ISNULL((Z.FSYNT),0) AS FSYNT, ISNULL((Y.SYNT),0) AS SYNT, ISNULL((X.MINR),0) AS MINR, " & _
               "ISNULL((W.TGBF),0) AS TGBF, ISNULL((V.SLLC),0) AS SLLC, ISNULL((U.TGGO),0) AS TGGO, ISNULL((T.BATTERY),0) AS BATTERY " & _
               "From "
               ''---------------------------------------------------------------------
               
               ''
      '         strSql = strSql & "( " & _
      '         "    Select X.TGLFAKTUR, X.KODEITEM, X.FSYNT * Y.QTY as FSYNT from " & _
      '         "    ( " & _
      '         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS FSYNT " & _
      '         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B ON A.NOFAKTUR = B.NOFAKTUR " & _
      '         "           WHERE A.NOFAKTUR IN " & _
      '         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
      '         "           where KODEGL IN " & _
      '         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
      '         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
      '         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
      '         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
      '         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
      '         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
      '         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
      '         "          ) "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
7              strSql = strSql & "( " & _
               "    Select X.TGLFAKTUR, X.KODEITEM, X.FSYNT * Y.QTY as FSYNT from " & _
               "    ( " & _
               "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS FSYNT " & _
               "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B ON A.NOFAKTUR = B.NOFAKTUR " & _
               "           WHERE A.NOFAKTUR IN " & _
               "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
               "           where KODEGL IN " & _
               "               (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02) " & _
               "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
               "          ) "
               ''----------------------------------------------------------------------
               
      '         strSql = strSql & " Union " & _
      '         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
      '         "           where G.KODEGL IN " & _
      '         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
      '         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
      '         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
      '         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
      '         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
      '         "          )) " & _
      '         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'FSYNT')) " & _
      '         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
      '         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'FSYNT') " & _
      '         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
      '         "      ) Z "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
8              strSql = strSql & " Union " & _
               "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
               "           where G.KODEGL IN " & _
               "                (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02) " & _
               "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
               "          )) " & _
               "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'FSYNT')) " & _
               "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
               "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'FSYNT') " & _
               "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
               "      ) Z "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & "FULL OUTER JOIN ( " & _
      '         "    Select X.TGLFAKTUR, X.KODEITEM, X.SYNT * Y.QTY as SYNT from ( " & _
      '         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS SYNT " & _
      '         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
      '         "           WHERE A.NOFAKTUR IN " & _
      '         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
      '         "           where KODEGL IN " & _
      '         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
      '         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
      '         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
      '         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
      '         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
      '         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
      '         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
      '         "          ) "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
9              strSql = strSql & "FULL OUTER JOIN ( " & _
               "    Select X.TGLFAKTUR, X.KODEITEM, X.SYNT * Y.QTY as SYNT from ( " & _
               "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS SYNT " & _
               "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
               "           WHERE A.NOFAKTUR IN " & _
               "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
               "           where KODEGL IN " & _
               "               (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02) " & _
               "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
               "          ) "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & " Union " & _
      '         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
      '         "           where G.KODEGL IN " & _
      '         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
      '         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
      '         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
      '         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
      '         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
      '         "          )) " & _
      '         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SYNT')) " & _
      '         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
      '         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SYNT') " & _
      '         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
      '         "      ) Y ON Z.TGLFAKTUR = Y.TGLFAKTUR And Z.KodeItem = Y.KodeItem "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
10             strSql = strSql & " Union " & _
               "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
               "           where G.KODEGL IN " & _
               "                (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
               "          )) " & _
               "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SYNT')) " & _
               "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
               "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SYNT') " & _
               "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
               "      ) Y ON Z.TGLFAKTUR = Y.TGLFAKTUR And Z.KodeItem = Y.KodeItem "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & "FULL OUTER JOIN ( " & _
      '         "    Select X.TGLFAKTUR, X.KODEITEM, X.MINR * Y.QTY as MINR from ( " & _
      '         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS MINR " & _
      '         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B  WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
      '         "           WHERE A.NOFAKTUR IN " & _
      '         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
      '         "           where KODEGL IN " & _
      '         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
      '         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
      '         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
      '         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
      '         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
      '         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
      '         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
      '         "          ) "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
11             strSql = strSql & "FULL OUTER JOIN ( " & _
               "    Select X.TGLFAKTUR, X.KODEITEM, X.MINR * Y.QTY as MINR from ( " & _
               "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS MINR " & _
               "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B  WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
               "           WHERE A.NOFAKTUR IN " & _
               "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
               "           where KODEGL IN " & _
               "               (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
               "          ) "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & " Union " & _
      '         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
      '         "           where G.KODEGL IN " & _
      '         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
      '         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
      '         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
      '         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
      '         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
      '         "          )) " & _
      '         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'MINR')) " & _
      '         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
      '         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'MINR') " & _
      '         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
      '         "      ) X ON Z.TGLFAKTUR = X.TGLFAKTUR AND Z.KodeItem = X.KodeItem  "
               
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
12             strSql = strSql & " Union " & _
               "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
               "           where G.KODEGL IN " & _
               "                (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
               "          )) " & _
               "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'MINR')) " & _
               "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
               "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'MINR') " & _
               "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
               "      ) X ON Z.TGLFAKTUR = X.TGLFAKTUR AND Z.KodeItem = X.KodeItem  "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & "FULL OUTER JOIN ( " & _
      '         "    Select X.TGLFAKTUR, X.KODEITEM, X.TGBF * Y.QTY as TGBF from ( " & _
      '         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS TGBF " & _
      '         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
      '         "           WHERE A.NOFAKTUR IN " & _
      '         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
      '         "           where KODEGL IN " & _
      '         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
      '         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
      '         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
      '         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
      '         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
      '         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
      '         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
      '         "          ) "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
13             strSql = strSql & "FULL OUTER JOIN ( " & _
               "    Select X.TGLFAKTUR, X.KODEITEM, X.TGBF * Y.QTY as TGBF from ( " & _
               "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS TGBF " & _
               "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
               "           WHERE A.NOFAKTUR IN " & _
               "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
               "           where KODEGL IN " & _
               "               (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
               "          ) "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & " Union " & _
      '         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
      '         "           where G.KODEGL IN " & _
      '         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
      '         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
      '         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
      '         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
      '         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
      '         "          )) " & _
      '         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGBF')) " & _
      '         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
      '         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGBF') " & _
      '         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
      '         "    ) W ON Z.TGLFAKTUR = W.TGLFAKTUR AND Z.KodeItem = W.KodeItem "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
14             strSql = strSql & " Union " & _
               "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
               "           where G.KODEGL IN " & _
               "                (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
               "          )) " & _
               "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGBF')) " & _
               "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
               "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGBF') " & _
               "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
               "    ) W ON Z.TGLFAKTUR = W.TGLFAKTUR AND Z.KodeItem = W.KodeItem "
               ''---------------------------------------------------------------------
               
               
      '         strSql = strSql & "FULL OUTER JOIN ( " & _
      '         "    Select X.TGLFAKTUR, X.KODEITEM, X.SLLC * Y.QTY as SLLC from ( " & _
      '         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS SLLC " & _
      '         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
      '         "           WHERE A.NOFAKTUR IN " & _
      '         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
      '         "           where KODEGL IN " & _
      '         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
      '         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
      '         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
      '         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
      '         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
      '         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
      '         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
      '         "          ) "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
15             strSql = strSql & "FULL OUTER JOIN ( " & _
               "    Select X.TGLFAKTUR, X.KODEITEM, X.SLLC * Y.QTY as SLLC from ( " & _
               "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS SLLC " & _
               "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
               "           WHERE A.NOFAKTUR IN " & _
               "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
               "           where KODEGL IN " & _
               "               (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
               "          ) "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & " Union " & _
      '         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
      '         "           where G.KODEGL IN " & _
      '         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
      '         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
      '         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
      '         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
      '         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
      '         "          )) " & _
      '         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SLLC')) " & _
      '         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
      '         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SLLC') " & _
      '         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem  " & _
      '         "    ) V ON Z.TGLFAKTUR = V.TGLFAKTUR AND Z.KodeItem = V.KodeItem "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
16             strSql = strSql & " Union " & _
               "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
               "           where G.KODEGL IN " & _
               "                (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
               "          )) " & _
               "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SLLC')) " & _
               "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
               "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'SLLC') " & _
               "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem  " & _
               "    ) V ON Z.TGLFAKTUR = V.TGLFAKTUR AND Z.KodeItem = V.KodeItem "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & "FULL OUTER JOIN ( " & _
      '         "    Select X.TGLFAKTUR, X.KODEITEM, X.TGGO * Y.QTY as TGGO from ( " & _
      '         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS TGGO " & _
      '         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
      '         "           WHERE A.NOFAKTUR IN " & _
      '         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
      '         "           where KODEGL IN " & _
      '         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
      '         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
      '         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
      '         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
      '         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
      '         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
      '         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
      '         "          ) "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
17             strSql = strSql & "FULL OUTER JOIN ( " & _
               "    Select X.TGLFAKTUR, X.KODEITEM, X.TGGO * Y.QTY as TGGO from ( " & _
               "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS TGGO " & _
               "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
               "           WHERE A.NOFAKTUR IN " & _
               "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
               "           where KODEGL IN " & _
               "               (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
               "          ) "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & " Union " & _
      '         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
      '         "           where G.KODEGL IN " & _
      '         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
      '         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
      '         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
      '         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
      '         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
      '         "          )) " & _
      '         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGGO')) " & _
      '         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
      '         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGGO') " & _
      '         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
      '         "      ) U ON Z.TGLFAKTUR = U.TGLFAKTUR AND Z.KodeItem = U.KodeItem "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
18             strSql = strSql & " Union " & _
               "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
               "           where G.KODEGL IN " & _
               "                (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
               "          )) " & _
               "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGGO')) " & _
               "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
               "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'TGGO') " & _
               "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
               "      ) U ON Z.TGLFAKTUR = U.TGLFAKTUR AND Z.KodeItem = U.KodeItem "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & "FULL OUTER JOIN ( " & _
      '         "    Select X.TGLFAKTUR, X.KODEITEM, X.BATTERY * Y.QTY as BATTERY from ( " & _
      '         "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS BATTERY " & _
      '         "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
      '         "           WHERE A.NOFAKTUR IN " & _
      '         "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
      '         "           where KODEGL IN " & _
      '         "               (Select SlipA as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipA <> '' Union " & _
      '         "                Select SlipAr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipAr <> '' Union " & _
      '         "                Select SlipB as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipB <> '' Union " & _
      '         "                Select SlipBr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipBr <> '' Union " & _
      '         "                Select SlipC as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipC <> '' Union " & _
      '         "                Select SlipCr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipCr <> '') " & _
      '         "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
      '         "          ) "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
19             strSql = strSql & "FULL OUTER JOIN ( " & _
               "    Select X.TGLFAKTUR, X.KODEITEM, X.BATTERY * Y.QTY as BATTERY from ( " & _
               "    SELECT B.KODEITEM, A.TGLFAKTUR, SUM(B.QTY) AS BATTERY " & _
               "           FROM SIHDR A WITH(NOLOCK) INNER JOIN SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR " & _
               "           WHERE A.NOFAKTUR IN " & _
               "           ((select NOBUKTI from gltrn WITH(NOLOCK) " & _
               "           where KODEGL IN " & _
               "               (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'A%' OR nobukti like 'B%' OR nobukti like 'C%') " & _
               "          ) "
               ''---------------------------------------------------------------------
               
      '         strSql = strSql & " Union " & _
      '         "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
      '         "           where G.KODEGL IN " & _
      '         "                (Select SlipE as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipE <> '' Union " & _
      '         "                Select SlipEr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipEr <> '' Union " & _
      '         "                Select SlipG as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipG <> '' Union " & _
      '         "                Select SlipGr as Slip from  " & gDb & "..SlipCode WITH(NOLOCK) Where SlipGr <> '' ) " & _
      '         "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
      '         "          )) " & _
      '         "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'BATR' )) " & _
      '         "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
      '         "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'BATR') " & _
      '         "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
      '         "      ) T ON Z.TGLFAKTUR = T.TGLFAKTUR AND Z.KodeItem = T.KodeItem " & _
      '         ") S "
               
               ''---------------------------------------------------------------------
               '' 2016-03-02 || ganti query pengambilan part
               ''---------------------------------------------------------------------
20             strSql = strSql & " Union " & _
               "           (select S.NOFAKTUR as NOBUKTI From SIHDR S WITH(NOLOCK) Inner Join GLTRN G WITH(NOLOCK) On G.NOWO = S.NOSO " & _
               "           where G.KODEGL IN " & _
               "                (SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN SPARE PART%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN Accesories%' AND TYPEREK = 02 " & _
               "               Union " & _
               "               SELECT KODEGL FROM Glmst WHERE NamaGl LIKE '%PENJUALAN OLIE%' AND TYPEREK = 02 ) " & _
               "          and (nobukti like 'E-%' OR nobukti like 'G-%') " & _
               "          )) " & _
               "          AND (Right(B.KODEITEM, LEN(B.KODEITEM)-2) IN (Select KodeItem from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'BATR' )) " & _
               "          GROUP BY A.TGLFAKTUR, B.KODEITEM, B.SATUANNS) " & _
               "    X INNER Join (Select KodeItem, Qty from " & gDb & "..TipePartList WITH(NOLOCK) Where TipePart = 'BATR') " & _
               "    Y On Right(X.KodeItem,len(X.KodeItem)-2) = Y.kodeItem " & _
               "      ) T ON Z.TGLFAKTUR = T.TGLFAKTUR AND Z.KodeItem = T.KodeItem " & _
               ") S "
               ''---------------------------------------------------------------------
               
21             strSql = strSql & " GROUP BY S.TanggalFAKTUR " & _
               ") Q ON P.TANGGAL = Q.TanggalFAKTUR " & _
               "ORDER BY P.TANGGAL, Q.TanggalFAKTUR"

      '    strSql = "SELECT TglFaktur, sum(Dpp+PPN) as rev"
      '    strSql = strSql & " From Sihdr"
      '    strSql = strSql & " Where Month(tglFaktur) = " & nBln & " And year(tglFaktur) = " & nThn & ""
      '    strSql = strSql & " AND (NoFaktur LIKE 'A%' OR NoFaktur LIKE 'B%' OR NoFaktur LIKE 'C%')"
      '    strSql = strSql & " group by TglFaktur"

22        Set rs = New ADODB.Recordset
23        rs.CursorLocation = adUseClient
24        rs.Open strSql, (Replace(SQLConACC, Log.LogACC, Left(Log.LogACC, 6) & nThn & IIf(Len(Trim(nBln)) = 1, "0" & nBln, nBln))), adOpenStatic, adLockBatchOptimistic, adCmdText
25        Set rs.ActiveConnection = Nothing
26        Set FetchLapPart = rs
27        If mConn.State <> adStateClosed Then mConn.Close
28        Exit Function
          
errHandler:
29    err.Raise err.Number, err.Source, err.Description
End Function





