VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEstimasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Public xKm As Double

Public Function FetchCUSTDAT() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "SELECT Distinct a.No_Chassis FROM CUSTDAT a WITH(NOLOCK) INNER JOIN history b WITH(NOLOCK) on A.no_chassis = b.no_chassis AND A.FS_LOG_DEL = '' WHERE a.No_Chassis <> '' Order By a.No_Chassis ASC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchCUSTDAT = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchHISTORY(Rangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "SELECT (Rep1 + Rep2 + Rep3 + Rep4 + Rep5 + Rep6 + Rep7) as Reps FROM History  WITH(NOLOCK) WHERE (Rep1 + Rep2 + Rep3 + Rep4 + Rep5 + Rep6 + Rep7) Like '%SERVIS BERKALA%' " & _
             "And No_Chassis = '" & Rangka & "' AND FS_LOG_DEL = '' Order By [Date] Desc"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchHISTORY = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchAllHISTORY(Rangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "SELECT * FROM HistoryUnitFunc_2 ('" & Rangka & "')"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchAllHISTORY = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Sub UpdateItemFlagKm(Rangka As String)
    Dim cm As New Command
On Error GoTo errHandler

    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    strSql = "Update CUSTDAT set FlagKm = " & xKm & " Where No_Chassis = '" & Rangka & "'"
   
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



' Get Data GRP - Adjustment SBE
Public Function GetDataGRP() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select a.NoInv, a.No_Wo, CAST(a.[Date] as Char(40)) as Tanggal, a.KM, " & _
             "a.Customer, a.Cust_adres, a.No_Polisi, a.No_Engine, a.No_Chassis, b.Telpon, b.GSM, b.CDMA, " & _
             "d.REP1, d.REP2, d.REP3, d.REP4, d.REP5, d.REP6, d.REP7, d.REP8, " & _
             "d.RTE1 , d.RTE2, d.RTE3, d.RTE4, d.RTE5, d.RTE6, d.RTE7, d.RTE8 " & _
             "FROM History a WITH(NOLOCK) Inner Join Custdat b WITH(NOLOCK) On a.No_Chassis = b.No_Chassis AND A.FS_LOG_DEL = '' " & _
             "INNER Join M_ToSBE c WITH(NOLOCK) On a.No_Wo = c.No_Wo " & _
             "INNER JOIN JOBDTA d WITH(NOLOCK) on c.No_Wo = d.No_Wo"
'    strSQL = "Select a.NoInv, a.No_Wo, CAST(a.[Date] as Char(40)) as Tanggal, a.KM, a.Customer, a.Cust_adres, a.No_Polisi, a.No_Engine, a.No_Chassis, b.Telpon, b.GSM, b.CDMA " & _
'             "FROM History a Inner Join Custdat b On a.No_Chassis = b.No_Chassis AND A.FS_LOG_DEL = '' INNER Join M_ToSBE c On a.No_Wo = c.No_Wo "
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GetDataGRP = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

' Get Data SBE - KM Terakhir
Public Function GetDatakmSBE(noRangka As String, noWO As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    If Log.RPLHistori = 0 Then
        strSql = "Select b.FS_Nilai FROM History a WITH(NOLOCK) INNER JOIN T_REGULAR b WITH(NOLOCK) ON a.FS_RC = b.FS_RC AND A.FS_LOG_DEL = '' " & _
             "Where a.SBE = 'SBE' And a.No_Chassis = '" & noRangka & "' Order By b.FS_Nilai Desc"
    Else
        strSql = "Select b.FS_Nilai FROM History a WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..T_REGULAR b WITH(NOLOCK) ON a.FS_RC = b.FS_RC AND A.FS_LOG_DEL = '' " & _
             "Where a.SBE = 'SBE' And a.No_Chassis = '" & noRangka & "' Order By b.FS_Nilai Desc"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GetDatakmSBE = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function



Public Sub DeleteItemGRP_On_ToSBE(noRangka As String, noWO As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "DELETE FROM M_ToSBE Where noRangka = '" & noRangka & "' AND No_Wo = '" & noWO & "'"
             
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

Public Sub AdjustGRP_ItemHistory(noRangka As String, noWO As String, IsKmSBE As String)
    Dim cm As New Command
On Error GoTo errHandler

    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    strSql = "Update History set SBE = 'SBE', FS_RC = " & IsKmSBE & " " & _
             "Where No_Chassis = '" & noRangka & "' And No_Wo = '" & noWO & "' AND FS_LOG_DEL = ''"
   
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = mConn
    cm.Execute
    
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    
    'Create M_Estimasi
    Dim RsE As ADODB.Recordset
    Set RsE = GetDataEstimasi(noRangka)
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    'Call mConn.Execute("EXEC NextEstimasi '" & RsE.Fields(0).value & "', '" & NullTgl(RsE.Fields(1).value) & "', " & Val(IIf(IsNumeric(RsE.Fields(2).value) = False, 0, RsE.Fields(2).value)) & ", " & RsE.Fields(3).value & ",5000,'" & RsE.Fields(4).value & "','" & RsE.Fields(5).value & "'")
    Call mConn.Execute("EXEC " & Trim(hTipePenghitunganDealer) & " '" & RsE.Fields(0).value & "', '" & NullTgl(RsE.Fields(1).value) & "', " & Val(IIf(IsNumeric(RsE.Fields(2).value) = False, 0, RsE.Fields(2).value)) & ", " & RsE.Fields(3).value & ",5000,'" & RsE.Fields(4).value & "','" & RsE.Fields(5).value & "'")
    If mConn.State = adStateOpen Then mConn.Close
   
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Private Function GetDataEstimasi(noRangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    strSql = "Select No_Chassis, [Date],KM ,NilSBE,NoINV, No_Wo " & _
             "FROM History WITH(NOLOCK) WHERE SBE = 'SBE' AND No_Chassis = '" & noRangka & "' AND FS_LOG_DEL = '' " & _
             "Order By No_Chassis, [Date] Desc"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GetDataEstimasi = rs
    If mConn.State = adStateOpen Then mConn.Close
    
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function panggilHistoryHO(noRgk As String) As ADODB.Recordset
    On Error GoTo vbError
    Dim rsHist As ADODB.Recordset
    Dim rsHlp As ADODB.Recordset
    Dim strSql As String
    ' AMBIL HISTORY DARI HO
    Dim rsHlp2 As ADODB.Recordset
    Dim dBaseHO As String
    Dim keteranganHide As String
    If Log.hideService = 1 Then
        keteranganHide = " AND A.KM > 10 "
    End If
    If Log.RPLHistori = 0 Then
        If Log.LogHOConn = 1 Then
            dBaseHO = HODb
            strSql = "SELECT CONVERT (varchar (10) , a.[Date],105) As TGL,a.no_wo, A.noinv,CAST(A.KM AS INTEGER) KM, "
            strSql = strSql & " A.FRONT, A.mech, A.FS_KD_DEALER,CASE WHEN A.FS_RC = 'R001' THEN 'SBI' WHEN A.FS_RC <> '' THEN 'SBE' ELSE (CASE WHEN A.TIPEWO = 'W' THEN 'TWC' ELSE 'GRP' END) END AS TIPE,"
            strSql = strSql & " REPLACE(A.REP1+' '+A.REP2+' '+A.REP3+' '+A.REP4+' '+A.REP5+' '+A.REP6+' '+A.REP7+' ','  ','') AS JOB,"
            strSql = strSql & " isnull(C.SUGGEST,'') AS SARAN, isnull(C.REMARK1,'') AS PESAN, isnull(C.CATAT,'') AS CATATAN FROM HISTORYnew A WITH(NOLOCK) "
            strSql = strSql & " LEFT JOIN JOBDTA C WITH(NOLOCK) ON A.NO_WO = C.NO_WO AND A.NOINV = C.NOINV WHERE A.FS_LOG_DEL = '' AND A.NO_CHASSIS = '" & Trim(noRgk) & "' " & keteranganHide & " ORDER BY a.[DATE] DESC"
            Set rsHlp = objO.FetchDataHO(strSql)
            Set panggilHistoryHO = rsHlp
        End If
    Else
        strSql = "SELECT CONVERT (varchar (10) , a.[Date],105) As TGL,a.no_wo, A.noinv,CAST(A.KM AS INTEGER) KM, "
        ''strSql = strSql & " A.FRONT, A.mech, A.FS_KD_DEALER,CASE WHEN A.FS_RC = 'R001' THEN 'SBI' WHEN A.FS_RC <> '' THEN 'SBE' ELSE (CASE WHEN A.TIPEWO = 'W' THEN 'TWC' ELSE 'GRP' END) END AS TIPE,"
        strSql = strSql & " A.FRONT, A.mech, A.FS_NM_DEALER as FS_KD_DEALER,CASE WHEN A.FS_RC = 'R001' THEN 'SBI' WHEN A.FS_RC <> '' THEN 'SBE' ELSE (CASE WHEN A.TIPEWO = 'W' THEN 'TWC' ELSE 'GRP' END) END AS TIPE,"
        strSql = strSql & " REPLACE(A.REP1+' '+A.REP2+' '+A.REP3+' '+A.REP4+' '+A.REP5+' '+A.REP6+' '+A.REP7+' ','  ','') AS JOB,"
        strSql = strSql & " isnull(C.SUGGEST,'') AS SARAN, isnull(C.REMARK1,'') AS PESAN, isnull(C.CATAT,'') AS CATATAN FROM [" & dbReplikasiHO & "]..HISTORYnew A WITH(NOLOCK) "
        strSql = strSql & " LEFT JOIN [" & dbReplikasiHO & "]..JOBDTA C WITH(NOLOCK) ON A.NO_WO = C.NO_WO AND A.NOINV = C.NOINV WHERE A.FS_LOG_DEL = '' AND A.NO_CHASSIS = '" & Trim(noRgk) & "' " & keteranganHide & " ORDER BY a.[DATE] DESC"
        Set rsHlp = objO.FetchData(strSql)
        Set panggilHistoryHO = rsHlp
    End If
    Exit Function
vbError:
    Exit Function
End Function

Public Sub panggilHistory(noRgk As String)
    On Error GoTo vbError
    Dim rsHist As ADODB.Recordset
    Dim rsHlp As ADODB.Recordset
        
    ' AMBIL HISTORY DARI HO
    Dim rsHlp2 As ADODB.Recordset
    Dim dBaseHO As String
    If Log.LogHOConn = 1 Then
        dBaseHO = HODb
        Set rsHlp = objO.FetchDataHO("SELECT * FROM HISTORYnew WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC")
        If Not rsHlp.EOF Then
            Do While Not rsHlp.EOF
                If rsHlp("NO_CHASSIS") <> "" Then
                    ' INPUT HISTORY
                    'AND KM = '" & rsHlp("KM") & "'
                    Set rsHlp2 = objO.FetchData("SELECT * FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC, NO_WO DESC")
                    If rsHlp2.EOF Then
                        Call objO.InsertData("INSERT INTO HISTORY VALUES ('" & rsHlp("CUSTOMER") & "', '" & rsHlp("CUST_ADRES") & "', '" & rsHlp("KOTA") & "','" & rsHlp("NO_POLISI") & "', '" & rsHlp("NO_ENGINE") & "','" & rsHlp("NO_CHASSIS") & "','" & rsHlp("FRONT") & "','" & rsHlp("MECH") & "','" & rsHlp("KM") & "','" & Format(rsHlp("DATE"), "mm/dd/yyyy") & "'," & rsHlp("PART_G") & ", " _
                                     & "" & rsHlp("PART_NG") & "," & rsHlp("OIL_G") & "," & rsHlp("OIL_NG") & "," & rsHlp("MTRL") & "," & rsHlp("LABOR") & "," & rsHlp("BODY") & "," & rsHlp("SBLT") & "," & rsHlp("MET") & "," & rsHlp("TOTAL") & ",'" & rsHlp("JOB") & "','" & rsHlp("REP1") & "','" & rsHlp("REP2") & "','" & rsHlp("REP3") & "','" & rsHlp("REP4") & "'," _
                                     & "'" & rsHlp("REP5") & "','" & rsHlp("REP6") & "','" & rsHlp("REP7") & "','" & rsHlp("MEKANIK1") & "','" & rsHlp("MEKANIK2") & "','" & rsHlp("MEKANIK3") & "'," _
                                     & "'" & rsHlp("MEKANIK4") & "','" & rsHlp("MEKANIK5") & "','" & rsHlp("MEKANIK6") & "','" & rsHlp("MEKANIK7") & "','" & rsHlp("MEKANIK8") & "','" & rsHlp("KLPKERJA1") & "','" & rsHlp("KLPKERJA2") & "','" & rsHlp("KLPKERJA3") & "','" & rsHlp("KLPKERJA4") & "','" & rsHlp("KLPKERJA5") & "','" & rsHlp("KLPKERJA6") & "'," _
                                     & "'" & rsHlp("KLPKERJA7") & "','" & rsHlp("KLPKERJA8") & "','" & rsHlp("NO_WO") & "','" & rsHlp("NOINV") & "','" & rsHlp("NOFAK") & "','" & rsHlp("STATUS") & "','" & rsHlp("PAID") & "','" & rsHlp("DISC") & "','" & rsHlp("DISCT") & "','" & rsHlp("REMARK1") & "','" & rsHlp("FN_RETURN") & "','" & rsHlp("FS_RC") & "'," _
                                     & "'" & rsHlp("FS_KD_DEALER") & "','" & rsHlp("FS_NM_DEALER") & "','" & rsHlp("TIPEWO") & "','" & rsHlp("KODETIPE") & "', '" & rsHlp("SBE") & "','" & rsHlp("FS_KD_PLG") & "','" & rsHlp("FS_KD_WIL") & "','" & rsHlp("THS") & "','" & rsHlp("MOU") & "','" & rsHlp("PDS") & "','" & rsHlp("FS_LOG_DEL") & "','" & rsHlp("FD_TGL_DEL") & "', " _
                                     & "'" & rsHlp("ALASANDEL") & "', '','" & Format(rsHlp("INVDATE"), "MM/DD/YYYY") & "')")
                    End If
                    
                    ' INPUT PART
                    Dim stringAwal As String
                    Dim stringPotong1 As String
                    Dim stringPotong2 As String
                    Dim nNext As Integer
                    Dim hasil(15) As String
                    If rsHlp("MEMOPART") <> "" Then
                        Set rsHlp2 = objO.FetchData("SELECT * FROM T_PART WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsHlp("NO_WO")) & "'")
                        If rsHlp2.EOF Then
                            stringAwal = Trim(rsHlp("MEMOPART"))
                            Do While Len(stringAwal) > 0
                                stringPotong1 = Left(stringAwal, InStr(stringAwal, "|") - 1)
                                nNext = 0
                                Do While Len(stringPotong1) > 0
                                    stringPotong2 = Left(stringPotong1, IIf(InStr(stringPotong1, "^") = 0, Len(stringPotong1), InStr(stringPotong1, "^") - 1))
                                    hasil(nNext) = Trim(stringPotong2)
                                    nNext = nNext + 1
                                    stringPotong1 = Mid(stringPotong1, InStr(stringPotong1, "^") + 1, Len(stringPotong1) - Len(stringPotong2))
                                Loop
                                stringAwal = Mid(stringAwal, InStr(stringAwal, "|") + 1, Len(stringAwal) - Len(stringPotong1))
                                
                                ' SIMPAN
                                Call objO.InsertData("INSERT INTO T_PART VALUES ('" & Trim(hasil(0)) & "','" & Trim(hasil(1)) & "','" & Trim(hasil(2)) & "','" & Trim(hasil(3)) & "','" & Trim(hasil(4)) & "', '" & IIf(Trim(hasil(5)), 1, 0) & "','" & Trim(hasil(6)) & "', " _
                                             & "'" & setKoma(Trim(hasil(7))) & "','" & Trim(hasil(8)) & "','" & Trim(hasil(9)) & "')")
                            Loop
                        End If
                    End If
                    
                    ' INPUT SUBLET
                    If rsHlp("MEMOSUBLET") <> "" Then
                        Set rsHlp2 = objO.FetchData("SELECT * FROM HIST_SUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsHlp("NO_WO")) & "' AND NO_CHASSIS = '" & noRgk & "'")
                        If rsHlp2.EOF Then
                            stringAwal = Trim(rsHlp("MEMOSUBLET"))
                            Do While Len(stringAwal) > 0
                                stringPotong1 = Left(stringAwal, InStr(stringAwal, "|") - 1)
                                nNext = 0
                                Do While Len(stringPotong1) > 0
                                    stringPotong2 = Left(stringPotong1, IIf(InStr(stringPotong1, "^") = 0, Len(stringPotong1), InStr(stringPotong1, "^") - 1))
                                    hasil(nNext) = Trim(stringPotong2)
                                    nNext = nNext + 1
                                    stringPotong1 = Mid(stringPotong1, InStr(stringPotong1, "^") + 1, Len(stringPotong1) - Len(stringPotong2))
                                Loop
                                stringAwal = Mid(stringAwal, InStr(stringAwal, "|") + 1, Len(stringAwal) - Len(stringPotong1))
                                
                                ' SIMPAN
                                Call objO.InsertData("INSERT INTO HIST_SUBLET VALUES ('" & Trim(hasil(0)) & "','" & Trim(hasil(1)) & "','" & Trim(hasil(2)) & "','" & Trim(hasil(3)) & "','" & Trim(hasil(4)) & "', '" & Trim(hasil(5)) & "','" & Trim(hasil(6)) & "', " _
                                             & "'" & Trim(hasil(7)) & "','" & Trim(hasil(8)) & "','" & Trim(hasil(9)) & "','" & Trim(hasil(10)) & "','" & Trim(hasil(11)) & "','" & Trim(hasil(12)) & "','" & Trim(hasil(13)) & "','" & Trim(hasil(14)) & "')")
                            Loop
                        End If
                    End If
                End If
                rsHlp.MoveNext
            Loop
        End If
    End If
    Exit Sub
vbError:
    Exit Sub
End Sub


