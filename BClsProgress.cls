VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BClsProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Public Function takeStallTime(namaHari As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM T_STALL WITH(NOLOCK) WHERE hari = '" & namaHari & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeStallTime = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeUnit(nTanggal As String, nTabel As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    Select Case nTabel
        Case 1
            ' STALL
            strSql = "SELECT * FROM stall WITH(NOLOCK) WHERE tanggal = '" & nTanggal & "' AND FLAG <> 'X' ORDER BY FLAG, noStall"
        Case 2
            ' TEMPPROGRESS
            strSql = "SELECT * FROM tempProgress WITH(NOLOCK) WHERE tanggal = '" & nTanggal & "' AND FLAGP <> 'M' ORDER BY NOSTALL, FLAGP DESC"   'START,
        Case 3
            ' HELPPROGRESS
            strSql = "SELECT * FROM HELPPROGRESS WITH(NOLOCK) WHERE tanggal = '" & nTanggal & "' AND FLAG = 'X' ORDER BY noStall"
    End Select

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    'rs.Open strSQL, oConn, adOpenDynamic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set takeUnit = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeMech(noOrder As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Trim(noOrder) & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeMech = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeDifference(noOrder As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & noOrder & "' AND DATEDIFF(n,START,PAUSE) > 0"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeDifference = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeDesc(stallNumber As String, timeUnit As String, sFlag As String, dTgl As Date) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
'    strSQL = "SELECT * FROM STALL WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(isnow, "mm/dd/yyyy") & "' AND " _
'           & " JamMulai <= '" & timeUnit & "' AND JamSelesai >= '" & timeUnit & "' AND FLAG LIKE '%" & Trim(sFlag) & "%' AND FLAG <> 'X' ORDER BY JAMMULAI DESC"
    strSql = "SELECT * FROM STALL WITH(NOLOCK) WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(dTgl, "mm/dd/yyyy") & "' AND " _
           & " JamMulai <= '" & timeUnit & "' AND JamSelesai >= '" & timeUnit & "' AND FLAG LIKE '%" & Trim(sFlag) & "%' AND FLAG <> 'X' ORDER BY JAMMULAI DESC"


    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeDesc = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeDescPending(stallNumber As String, timeUnit As String, dTgl As Date) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
'    strSQL = "SELECT SUBSTRING(KETERANGAN, CHARINDEX(' |', KETERANGAN)+2, (CHARINDEX('| ', KETERANGAN)-CHARINDEX(' |', KETERANGAN)-2)) AS NOWO, LEFT(KETERANGAN, CHARINDEX(' |', KETERANGAN)-1) AS NOPOL,  " _
'           & "CUSTOMER , noStall, jamMulai, a.MECH, REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, NOFAK " _
'           & "FROM STALL A INNER JOIN JOBDTA B ON SUBSTRING(KETERANGAN, CHARINDEX(' |', KETERANGAN)+2, (CHARINDEX('| ', KETERANGAN)-CHARINDEX(' |', KETERANGAN)-2)) = B.NO_WO " _
'           & "WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(isnow, "mm/dd/yyyy") & "' AND " _
'           & "JamMulai <= '" & timeUnit & "' AND (FLAG = 'N' OR FLAG = '')  ORDER BY JAMMULAI DESC"
    
    strSql = "SELECT SUBSTRING(KETERANGAN, CHARINDEX(' |', KETERANGAN)+2, (CHARINDEX('| ', KETERANGAN)-CHARINDEX(' |', KETERANGAN)-2)) AS NOWO, LEFT(KETERANGAN, CHARINDEX(' |', KETERANGAN)-1) AS NOPOL,  " _
           & "CUSTOMER , noStall, jamMulai, a.MECH, REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, NOFAK " _
           & "FROM STALL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON SUBSTRING(KETERANGAN, CHARINDEX(' |', KETERANGAN)+2, (CHARINDEX('| ', KETERANGAN)-CHARINDEX(' |', KETERANGAN)-2)) = B.NO_WO " _
           & "WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(dTgl, "mm/dd/yyyy") & "' AND " _
           & "JamMulai <= '" & timeUnit & "' AND (FLAG = 'N' OR FLAG = '')  ORDER BY JAMMULAI DESC"
           
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeDescPending = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

'takeDescAdd(Trim(rsU("NO_WO")))
Public Function takeDescAdd(nWO As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
'    strSQL = "SELECT * FROM STALL WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(isnow, "mm/dd/yyyy") & "' AND " _
'           & " JamMulai <= '" & timeUnit & "' AND (FLAG = 'N' OR FLAG = '')  ORDER BY JAMMULAI DESC"

    strSql = "SELECT * FROM JOBADD2 WITH(NOLOCK) WHERE NO_WO = '" & Trim(nWO) & "' ORDER BY [DATE], [TIME], KLPKERJA DESC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeDescAdd = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeDescBook(noBook As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM BOOK WITH(NOLOCK) WHERE NOFAK = '" & Trim(noBook) & "' ORDER BY [DATE], [TIME]"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeDescBook = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeDetailDiag(nOrder As String, nIdx As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    Select Case nIdx
        Case 1
            ' UTK ALL UNIT
            strSql = "SELECT A.*, TUNGGUSVC, DATESERAH, TIMESERAH FROM diagdta A WITH(NOLOCK) INNER JOIN diagDTL B WITH(NOLOCK) ON A.NO_diag = B.NO_diag WHERE A.no_diag = '" & nOrder & "' AND deliv <> 'BTL'"
        Case 2
            ' UNIT YG BELUM CLOSING
            strSql = "SELECT A.*, TUNGGUSVC, DATESERAH, TIMESERAH FROM diagdta A WITH(NOLOCK) INNER JOIN diagDTL B WITH(NOLOCK) ON A.NO_diag = B.NO_diag WHERE A.no_diag = '" & nOrder & "' AND deliv <> 'BTL' AND RPR_COND <> 'K'"
    End Select

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeDetailDiag = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function takeDetailJob(nOrder As String, nIdx As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    Select Case nIdx
        Case 1
            ' UTK ALL UNIT
            strSql = "SELECT A.*, TUNGGUSVC, DATESERAH, TIMESERAH FROM jobdta A WITH(NOLOCK) INNER JOIN JOBDTL B WITH(NOLOCK) ON A.NO_WO = B.NO_WO WHERE A.no_wo = '" & nOrder & "' AND deliv <> 'BTL'"
        Case 2
            ' UNIT YG BELUM CLOSING
            strSql = "SELECT A.*, TUNGGUSVC, DATESERAH, TIMESERAH FROM jobdta A WITH(NOLOCK) INNER JOIN JOBDTL B WITH(NOLOCK) ON A.NO_WO = B.NO_WO WHERE A.no_wo = '" & nOrder & "' AND deliv <> 'BTL' AND RPR_COND <> 'K'"
    End Select

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeDetailJob = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeDeffJob(jobGroup As String, jobName As String, unitModel As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
          
          
3       '  If Log.RPLHistori = 0 Then
4             strSql = "SELECT * FROM actrate WITH(NOLOCK) WHERE JOB = '" & jobGroup & "' AND FS_ITEM = '" & jobName & "' AND MDL = '" & unitModel & "' AND FS_LOG_DEL = ''"
5       '  Else
6        '     strSql = "SELECT * FROM [" & dbReplikasiHO & "]..actrate WITH(NOLOCK) WHERE JOB = '" & jobGroup & "' AND FS_ITEM = '" & jobName & "' AND MDL = '" & unitModel & "' AND FS_LOG_DEL = ''"
7        ' End If
8         Set rs = New ADODB.Recordset
9         rs.CursorLocation = adUseClient
10        rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
11        Set rs.ActiveConnection = Nothing
12        Set takeDeffJob = rs
13        If mConn.State <> adStateClosed Then mConn.Close
14        Exit Function
          
errHandler:
15    err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchFromBooking(nOrder As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM BOOK WITH(NOLOCK) WHERE NOFAK = '" & nOrder & "' AND deliv <> 'BTL'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchFromBooking = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function countDifference(nOrder As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set countDifference = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchTempProgress(nOrder As String) As Boolean
    On Error GoTo errHandler
    searchTempProgress = False
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.EOF Then
        searchTempProgress = True
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takePauseState(nOrder As String, nChassis As String, nTangl As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM PROGRESS WITH(NOLOCK) WHERE noorder = '" & nOrder & "' AND no_chassis = '" & nChassis & "' AND FLAG = 'P' ORDER BY TANGGAL DESC, JAM DESC" 'AND tanggal = '" & Format(nTangl, "mm/dd/yyyy") & "'

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takePauseState = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchHelpProgres(stallNumber As String, timeUnit As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM HELPPROGRESS WITH(NOLOCK) WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(isnow, "mm/dd/yyyy") & "' AND " _
           & " JamMulai <= '" & timeUnit & "' ORDER BY JAMMULAI DESC"

    'CAST(LEFT(JamMulai,2) as INTEGER)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchHelpProgres = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function pauseProses(nOrder As String, nameMech As String, stallNumber As String, dReason As String, nChassis As String, nTangl As String, nDay As String, timeStart As String, nActRt As Double, timeFinish As String)
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim rsCari As ADODB.Recordset
Dim nFO As String

    On Error GoTo errHandler
    
     
    Set rsCari = objO.FetchData("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND DATEDIFF(n,JAMMULAI, '" & Format(timeStart, "00:00") & "') < 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeFinish, "00:00") & "') > 0 AND FLAG IN ('','S','P','N','F')")
    If Not rsCari.EOF Then
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
        Exit Function
    Else
        Dim rsHlp3 As ADODB.Recordset
         
        Set rsHlp3 = objO.FetchData("SELECT GROUP FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
        If Not rsHlp3.EOF Then
            nFO = Trim(rsHlp3("GROUP"))
        End If
        rsHlp3.Close
        
        ' SIMPAN PAUSE
        oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','P','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "', '" & gSA & "','" & dReason & "','" & nameMech & "')")
        oConn.Execute ("UPDATE tempProgress SET flagp = 'P', pause = '" & Format(isnow, "HH:NN") & "', MECH = '" & nameMech & "',GRP = '" & Trim(nFO) & "' WHERE noorder = '" & nOrder & "' AND tanggal = '" & Format(isnow, "mm/dd/yyyy") & "' AND NOSTALL = '" & stallNumber & "' AND FLAGP = ''")
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
         
        Set rs = oConn.Execute("SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
        If Not rs.EOF Then
            nSelisih = IIf(Nul(rs("SELISIH")) <> "", rs("SELISIH"), 0)
            totalRate = CDbl(nActRt) * 60
            sisaRate = totalRate - CDbl(nSelisih)
            rateBaru = Round(CDbl(nSelisih) / 60, 1)
        End If
        rs.Close
        
        ' FLAG PAUSE
         
        Set rs = oConn.Execute("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
        If Not rs.EOF Then
            'oConn.Execute ("UPDATE STALL SET flag = 'P', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 20, Round((20 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 20, Round((20 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
            '' ok sebelum ada Qc
            'oConn.Execute ("UPDATE STALL SET flag = 'P', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
            oConn.Execute ("UPDATE STALL SET flag = '" & IIf(Left(dReason, 5) = "Final", "H", "P") & "', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
        End If
        rs.Close
        
        oConn.Execute ("UPDATE jobdta SET time_a = '" & Format(isnow, "HH:NN") & "', rpr_cond = 'T' WHERE no_chassis = '" & nChassis & "' AND no_wo = '" & nOrder & "' ")
        
        ' UPDATE JAM TERPAKAI MEKANIK
        oConn.Execute ("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI - " & Replace(CDbl(sisaRate), ",", ".") & ", ALOKASI = ALOKASI - " & Replace(CDbl(nSelisih), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & nameMech & "'")
        
        ' INSERT JOB YG DIKERJAKAN KE TEMP
        Dim nMinBaru As Integer
        nMinBaru = 0
         
        Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
        If Not rsCari.EOF Then
            nMinBaru = CDbl(rsCari("SELISIH"))
        Else
            nMinBaru = nSelisih
        End If
        'nMinBaru = nMinBaru + (10 - (nMinBaru Mod 10))
        rsCari.Close
        Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 1, Trim(stallNumber))
        Exit Function
        
'        If oConn.Errors.Count = 0 Then
'            oConn.CommitTrans
'        Else
'            oConn.RollbackTrans
'            MsgBox "Transaksi Gagal, Ulangi Proses."
'        End If
    End If
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function stopInspection(nOrder As String, nameMech As String, stallNumber As String, dReason As String, nChassis As String, nTangl As String, nDay As String, timeStart As String, nActRt As Double, timeFinish As String)
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim rsCari As New ADODB.Recordset
Dim nFO As String

    On Error GoTo errHandler
     
    Set rsCari = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND DATEDIFF(n,JAMMULAI, '" & Format(timeStart, "00:00") & "') < 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeFinish, "00:00") & "') > 0 AND FLAG IN ('','S','P','N','F')")
    If Not rsCari.EOF Then
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
        Exit Function
    Else
        oConn.BeginTrans
        oConn.Errors.Clear
        Dim rsHlp3 As New ADODB.Recordset
         
        Set rsHlp3 = oConn.Execute("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
        If Not rsHlp3.EOF Then
            nFO = Trim(rsHlp3("GROUP"))
        End If
        rsHlp3.Close
        
        ' SIMPAN PAUSE
        oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','O','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "', '" & gSA & "','" & dReason & "','" & gSA & "')")
        oConn.Execute ("UPDATE tempProgress SET flagp = 'P', pause = '" & Format(isnow, "HH:NN") & "', MECH = '" & gSA & "',GRP = '" & Trim(nFO) & "' WHERE noorder = '" & nOrder & "' AND tanggal = '" & Format(isnow, "mm/dd/yyyy") & "' AND NOSTALL = '" & stallNumber & "' AND FLAGP = ''")
                
        ' FLAG PAUSE
         
        Set rs = oConn.Execute("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
        If Not rs.EOF Then
            oConn.Execute ("UPDATE STALL SET flag = 'G', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
        End If
        rs.Close
        
        oConn.Execute ("UPDATE jobdta SET rpr_cond = 'T' WHERE no_chassis = '" & nChassis & "' AND no_wo = '" & nOrder & "' ")
        
        ' INSERT JOB YG DIKERJAKAN KE TEMP
'        Dim nMinBaru As Integer
'        nMinBaru = 0
'
'        Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
'        If Not rsCari.EOF Then
'            nMinBaru = CDbl(rsCari("SELISIH"))
'        Else
'            nMinBaru = nSelisih
'        End If
'        rsCari.Close
'        Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 1, Trim(stallNumber))
        
        If oConn.Errors.Count = 0 Then
            oConn.CommitTrans
        Else
            oConn.RollbackTrans
            MsgBox "Transaksi Gagal, Ulangi Proses."
        End If
    End If
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function pauseProses2(nOrder As String, stallNumber As String, dReason As String, nChassis As String, nTangl As String, nDay As String, timeStart As String, nActRt As Double, timeFinish As String)
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim rsCari As New ADODB.Recordset
Dim nFO As String
Dim rsSearch As New ADODB.Recordset
Dim nameMech As String
Dim i As Integer

    On Error GoTo errHandler
    
     
    'Set rsCari = oConn.Execute("SELECT * FROM STALL WHERE tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & stallNumber & "' AND DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeStart, "00:00") & "') < 0 AND FLAG IN ('','S','P','N','F')")
    Set rsCari = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND DATEDIFF(n,JAMMULAI, '" & Format(timeStart, "00:00") & "') < 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeFinish, "00:00") & "') > 0 AND FLAG IN ('','S','P','N','F')")
    If Not rsCari.EOF Then
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
        Exit Function
    Else
'        oConn.BeginTrans
'        oConn.Errors.Clear
         
        Set rsSearch = oConn.Execute("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & Trim(nOrder) & "' AND FLAGP = '' ORDER BY MECH")
        If Not rsSearch.EOF Then
            i = 0
            Do While Not rsSearch.EOF
                i = i + 1
                nameMech = Trim(rsSearch("MECH"))
                
                Dim rsHlp3 As New ADODB.Recordset
                 
                Set rsHlp3 = oConn.Execute("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
                If Not rsHlp3.EOF Then
                    nFO = Trim(rsHlp3("GROUP"))
                End If
                rsHlp3.Close
                
                ' SIMPAN PAUSE
                oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','P','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "', '" & gSA & "','" & dReason & "','" & nameMech & "')")
                oConn.Execute ("UPDATE tempProgress SET flagp = 'P', pause = '" & Format(isnow, "HH:NN") & "', GRP = '" & Trim(nFO) & "' WHERE noorder = '" & nOrder & "' AND tanggal = '" & Format(isnow, "mm/dd/yyyy") & "' AND NOSTALL = '" & stallNumber & "' AND FLAGP = '' AND MECH = '" & nameMech & "'")
                
                If i = 1 Then
                    Set rs = New ADODB.Recordset
                    rs.CursorLocation = adUseClient
                     
                    Set rs = oConn.Execute("SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
                    If Not rs.EOF Then
                        nSelisih = IIf(Nul(rs("SELISIH")) <> "", rs("SELISIH"), 0)
                        totalRate = CDbl(nActRt) * 60
                        sisaRate = totalRate - CDbl(nSelisih)
                        rateBaru = Round(CDbl(nSelisih) / 60, 1)
                    End If
                    rs.Close
                
                
                    ' FLAG PAUSE
                     
                    Set rs = oConn.Execute("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
                    If Not rs.EOF Then
                        'oConn.Execute ("UPDATE STALL SET flag = 'P', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 20, Round((20 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 20, Round((20 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
                        'oConn.Execute ("UPDATE STALL SET flag = 'P', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
                        oConn.Execute ("UPDATE STALL SET flag = '" & IIf(Left(dReason, 5) = "Final", "H", "P") & "', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
                    End If
                    rs.Close
                    
                    oConn.Execute ("UPDATE jobdta SET time_a = '" & Format(isnow, "HH:NN") & "', rpr_cond = 'T' WHERE no_chassis = '" & nChassis & "' AND no_wo = '" & nOrder & "' ")
                End If
                
                ' UPDATE JAM TERPAKAI MEKANIK
                oConn.Execute ("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI - " & Replace(CDbl(sisaRate), ",", ".") & ", ALOKASI = ALOKASI - " & Replace(CDbl(nSelisih), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & nameMech & "'")
                
                If i = 1 Then
                    ' INSERT JOB YG DIKERJAKAN KE TEMP
                    Dim nMinBaru As Integer
                    'Dim rsCari As New ADODB.Recordset
                    nMinBaru = 0
                     
                    Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
                    If Not rsCari.EOF Then
                        nMinBaru = CDbl(rsCari("SELISIH"))
                    Else
                        nMinBaru = nSelisih
                    End If
                    'nMinBaru = nMinBaru + (10 - (nMinBaru Mod 10))
                    rsCari.Close
                    Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 1, Trim(stallNumber))
                End If
                rsSearch.MoveNext
            Loop
        End If
        
'        If oConn.Errors.Count = 0 Then
'            oConn.CommitTrans
'        Else
'            oConn.RollbackTrans
'            MsgBox "Transaksi Gagal, Ulangi Proses."
'        End If
        Exit Function
    End If
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function changeMech(nOrder As String, nTangl As String, nChassis As String, stallNumber As String, nameMech As String, nActRt As Double, nDay As String, timeStart As Double, timeFinish As Double) As Boolean
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim nFO As String
Dim rsHlp2 As New ADODB.Recordset
Dim rsHlp As New ADODB.Recordset
Dim okProses As Boolean

    'On Error GoTo errHandler
    okProses = True
    Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') < 0)) AND FLAG IN ('','S','P','N','F')")
    If Not rsHlp.EOF Then
        okProses = False
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
    End If
    If rsHlp.State <> adStateClosed Then rsHlp.Close
    
    If okProses Then
        changeMech = False
         
        Set rsHlp2 = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE NOSTALL = '" & Trim(stallNumber) & "' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAG = 'S'")
        If rsHlp2.EOF Then
    '        oConn.BeginTrans
    '        oConn.Errors.Clear
            Set rs = New ADODB.Recordset
             
            Set rs = oConn.Execute("SELECT * FROM Stall WITH(NOLOCK) WHERE keterangan LIKE '%|" & nOrder & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "'")
            If Not rs.EOF Then
                
                ' CARI GROUP / FO MEKANIK
                 
                Set rsHlp = oConn.Execute("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
                If Not rsHlp.EOF Then
                    nFO = Trim(rsHlp("GROUP"))
                End If
                rsHlp.Close
                
                ' SIMPAN MEKANIK
                oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','S','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "', '" & gSA & "','GANTI MEKANIK','" & nameMech & "')")
                oConn.Execute ("INSERT INTO TEMPPROGRESS VALUES ('" & nOrder & "', '" & Format(nTangl, "MM/DD/YYYY") & "', '" & stallNumber & "', '" & Format(isnow, "HH:NN") & "', '','', '" & nameMech & "','" & Trim(nFO) & "')")
                
                ' UPDATE JAM TERPAKAI MEKANIK
                 
                Set rsHlp = oConn.Execute("SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
                If Not rsHlp.EOF Then
                    nSelisih = IIf(rsHlp("SELISIH") > 0, rsHlp("SELISIH"), 0)
                    totalRate = CDbl(nActRt) * 60
                    sisaRate = totalRate - CDbl(nSelisih)
                    rateBaru = Round(CDbl(IIf(sisaRate < 20, 20, sisaRate)) / 60, 1)
                End If
                rsHlp.Close
                
                ', NOSTALL = '" & Trim(stallBaru) & "'
                oConn.Execute ("UPDATE STALL SET MECH = '" & nameMech & "', JAMMULAI = '" & pembulatanJam(Format(isnow, "HH:NN")) & "', RATE = " & Replace(CDbl(rateBaru), ",", ".") & ", flag = 'S', JAMSELESAI = '" & Left(hitungWaktu(CDbl(IIf(rateBaru < 20, 20, rateBaru)), Format(isnow, "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
                oConn.Execute ("UPDATE JOBDTA SET MECH = '" & nameMech & "', [GROUP] = '" & Trim(nFO) & "', RPR_COND = 'P' WHERE NO_WO = '" & Trim(nOrder) & "' AND NO_CHASSIS = '" & nChassis & "'")
                oConn.Execute ("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI + " & Replace(CDbl(sisaRate), ",", ".") & ", ALOKASI = ALOKASI + " & Replace(CDbl(sisaRate), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & nameMech & "'")
                If Trim(rs("FLAG")) = "N" Then
                    oConn.Execute ("UPDATE T_MEKANIK SET ALOKASI = ALOKASI - " & Replace(CDbl(totalRate), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & Trim(rs("MECH")) & "'")
                End If
                
                ' UPDATE FLAG HELPPROGRESS
                Dim nMinBaru As Integer
                Dim rsCari As New ADODB.Recordset
                nMinBaru = 0
                 
                Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
                If Not rsCari.EOF Then
                    nMinBaru = CDbl(rsCari("SELISIH"))
                Else
                    nMinBaru = nSelisih
                End If
                rsCari.Close
                Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 2, Trim(stallNumber))
                changeMech = True
            End If
            rs.Close
            
    '        If oConn.Errors.Count = 0 Then
    '            oConn.CommitTrans
    '        Else
    '            oConn.RollbackTrans
    '            MsgBox "Transaksi Gagal, Ulangi Proses"
    '        End If
        Else
            MsgBox "Stall Dipakai unit lain. " & vbCr & "Kendaraan dengan No. Polisi : " & Left(Trim(rsHlp2("KETERANGAN")), InStr(Trim(rsHlp2("KETERANGAN")), " |"))
        End If
        If rsHlp2.State <> adStateClosed Then rsHlp2.Close
    End If
    Exit Function

errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function changeMech2(nOrder As String, nTangl As String, nChassis As String, stallNumber As String, nameMech2 As String, nActRt As Double, nDay As String, timeStart As Double, timeFinish As Double) As Boolean
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim nFO As String
Dim rsHlp2 As New ADODB.Recordset
Dim j As Integer
Dim nameMech As String
Dim rsHlp As New ADODB.Recordset
Dim okProses As Boolean

    On Error GoTo errHandler
    okProses = True
     
    'Set rsHlp = oConn.Execute("SELECT * FROM STALL WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeFinish, "00:00") & "') > 0 AND FLAG IN ('','S','P','N')")
    'Set rsHlp = oConn.Execute("SELECT * FROM STALL WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 OR DATEDIFF(n,JAMMULAI, '" & Format(timeFinish, "00:00") & "') > 0) AND FLAG IN ('','S','P','N')")
    'Set rsHlp = oConn.Execute("SELECT * FROM STALL WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0)) AND FLAG IN ('','S','P','N')")
    Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') < 0)) AND FLAG IN ('','S','P','N','F')")
    If Not rsHlp.EOF Then
        okProses = False
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
    End If
    If rsHlp.State <> adStateClosed Then rsHlp.Close
    
    If okProses Then
        changeMech2 = False
         
        Set rsHlp2 = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE NOSTALL = '" & Trim(stallNumber) & "' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAG = 'S'")
        If rsHlp2.EOF Then
            Set rs = New ADODB.Recordset
             
            Set rs = oConn.Execute("SELECT * FROM Stall WITH(NOLOCK) WHERE keterangan LIKE '%|" & nOrder & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "'")
            If Not rs.EOF Then
                
                For j = 1 To jmlMekanikEM(nameMech2)
                    nameMech = pecahMekanik(nameMech2, j)
                    
                    ' CARI GROUP / FO MEKANIK
                     
                    Set rsHlp = oConn.Execute("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
                    If Not rsHlp.EOF Then
                        nFO = Trim(rsHlp("GROUP"))
                    End If
                    rsHlp.Close
                    
                    ' SIMPAN MEKANIK
                    oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','S','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "', '" & gSA & "','GANTI MEKANIK','" & nameMech & "')")
                    oConn.Execute ("INSERT INTO TEMPPROGRESS VALUES ('" & nOrder & "', '" & Format(nTangl, "MM/DD/YYYY") & "', '" & stallNumber & "', '" & Format(isnow, "HH:NN") & "', '','', '" & nameMech & "','" & Trim(nFO) & "')")
                    
                    ' UPDATE JAM TERPAKAI MEKANIK
                     
                    Set rsHlp = oConn.Execute("SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
                    If Not rsHlp.EOF Then
                        nSelisih = IIf(rsHlp("SELISIH") > 0, rsHlp("SELISIH"), 0)
                        totalRate = CDbl(nActRt) * 60
                        sisaRate = totalRate - CDbl(nSelisih)
                        rateBaru = Round(CDbl(IIf(sisaRate < 20, 20, sisaRate)) / 60, 1)
                    End If
                    rsHlp.Close
                    
                    ', NOSTALL = '" & Trim(stallBaru) & "'
                    If j = 1 Then
                        oConn.Execute ("UPDATE STALL SET MECH = '" & nameMech & "', JAMMULAI = '" & pembulatanJam(Format(isnow, "HH:NN")) & "', RATE = " & Replace(CDbl(rateBaru), ",", ".") & ", flag = 'S', JAMSELESAI = '" & Left(hitungWaktu(CDbl(IIf(rateBaru < 20, 20, rateBaru)), Format(isnow, "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
                        oConn.Execute ("UPDATE JOBDTA SET MECH = '" & nameMech & "', [GROUP] = '" & Trim(nFO) & "', SPK = '" & Replace(Trim(nameMech2), "'", "") & "', RPR_COND = 'P' WHERE NO_WO = '" & Trim(nOrder) & "' AND NO_CHASSIS = '" & nChassis & "'")
                    End If
                    oConn.Execute ("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI + " & Replace(CDbl(sisaRate), ",", ".") & ", ALOKASI = ALOKASI + " & Replace(CDbl(sisaRate), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & nameMech & "'")
                    If Trim(rs("FLAG")) = "N" Then
                        oConn.Execute ("UPDATE T_MEKANIK SET ALOKASI = ALOKASI - " & Replace(CDbl(totalRate), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & Trim(rs("MECH")) & "'")
                    End If
                    
                    ' UPDATE FLAG HELPPROGRESS
                    Dim nMinBaru As Integer
                    Dim rsCari As New ADODB.Recordset
                    nMinBaru = 0
                     
                    Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
                    If Not rsCari.EOF Then
                        nMinBaru = CDbl(rsCari("SELISIH"))
                    Else
                        nMinBaru = nSelisih
                    End If
                    rsCari.Close
                    Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 2, Trim(stallNumber))
                Next
                changeMech2 = True
            End If
            rs.Close
        Else
            MsgBox "Stall Dipakai unit lain. " & vbCr & "Kendaraan dengan No. Polisi : " & Left(Trim(rsHlp2("KETERANGAN")), InStr(Trim(rsHlp2("KETERANGAN")), " |"))
        End If
        If rsHlp2.State <> adStateClosed Then rsHlp2.Close
    End If
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function finishProses(nOrder As String, nameMech As String, stallNumber As String, dReason As String, nChassis As String, nTangl As String, nDay As String, timeStart As String, nActRt As Double, PRC As Boolean, timeFinish As String)
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim rsCari As New ADODB.Recordset
Dim nFO As String
    
     
        Set rsCari = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') < 0)) AND FLAG IN ('','S','P','N','F')")
    If Not rsCari.EOF Then
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
        Exit Function
    Else
        oConn.Errors.Clear
        oConn.BeginTrans
        
        Dim rsHlp3 As New ADODB.Recordset
         
        Set rsHlp3 = oConn.Execute("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
        If Not rsHlp3.EOF Then
            nFO = Trim(rsHlp3("GROUP"))
        End If
        rsHlp3.Close
    
        ' SIMPAN FINISH
        oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','F','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "','" & gSA & "' , '','" & nameMech & "')")
        Set rs = New ADODB.Recordset
         
        Set rs = oConn.Execute("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE noorder = '" & nOrder & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND noStall = '" & stallNumber & "' AND FLAGP = ''")
        If Not rs.EOF Then
            oConn.Execute ("UPDATE tempProgress SET flagp = 'F', pause = CONVERT(varchar(5), GETDATE(),8), MECH = '" & nameMech & "', GRP = '" & Trim(nFO) & "' WHERE noorder = '" & nOrder & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND noStall = '" & stallNumber & "' AND FLAGP = ''")
        Else
            oConn.Execute ("INSERT INTO tempProgress VALUES ('" & nOrder & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & stallNumber & "', '" & Format(DateAdd("n", -3, isnow), "HH:NN") & "', 'F','" & Format(isnow, "HH:NN") & "','" & nameMech & "','" & Trim(nFO) & "')")
        End If
        rs.Close
        
        Set rs = New ADODB.Recordset
         
        Set rs = oConn.Execute("SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
        If Not rs.EOF Then
            nSelisih = IIf(rs("SELISIH") > 0, rs("SELISIH"), 0)
            totalRate = CDbl(nActRt) * 60
            sisaRate = Abs(totalRate - CDbl(nSelisih))
        End If
        rs.Close
        
        ' FLAG FINISH
        Set rs = New ADODB.Recordset
         
        Set rs = oConn.Execute("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
        If Not rs.EOF Then
            'oConn.Execute ("UPDATE STALL SET flag = 'F', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 20, Round((20 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 20, Round((20 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
            oConn.Execute ("UPDATE STALL SET flag = 'F', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
        End If
        rs.Close
        
        Dim namaMechAkhir As String
        namaMechAkhir = Trim(nameMech)
        Set rs = New ADODB.Recordset
         
        Set rs = oConn.Execute("SELECT TOP 1 A.MECH FROM TEMPPROGRESS A WITH(NOLOCK) INNER JOIN PROGRESS B WITH(NOLOCK) ON A.NOORDER = B.NOORDER AND A.NOSTALL = B.NOSTALL AND A.START = B.JAM AND A.MECH = B.MECH WHERE A.NOORDER = '" & Trim(nOrder) & "' AND FLAG NOT IN ('Q','A','O') GROUP BY A.MECH ORDER BY SUM(DATEDIFF(n,START,PAUSE)) DESC")
        If Not rs.EOF Then
            namaMechAkhir = Trim(rs("MECH"))
        End If
        rs.Close
        oConn.Execute ("UPDATE jobdta SET time_S = '" & Format(isnow, "HH:NN") & "', rpr_cond = 'S', deliv = '" & IIf(PRC, "PRC", "OKE") & "', MECH = '" & Trim(namaMechAkhir) & "' WHERE no_chassis = '" & nChassis & "' AND no_wo = '" & nOrder & "' ")
        
        ' UPDATE JAM TERPAKAI MEKANIK
        oConn.Execute ("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI - " & Replace(CDbl(sisaRate), ",", ".") & ", ALOKASI = ALOKASI - " & Replace(CDbl(sisaRate), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & nameMech & "'")
        
        ' INSERT JOB YG DIKERJAKAN KE TEMP
        Dim nMinBaru As Integer
        nMinBaru = 0
         
        Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
        If Not rsCari.EOF Then
            nMinBaru = CDbl(rsCari("SELISIH"))
        Else
            nMinBaru = nSelisih
        End If
        'nMinBaru = nMinBaru + (10 - (nMinBaru Mod 10))
        rsCari.Close
        'Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 1)
        Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 3, Trim(stallNumber))
        
        If oConn.Errors.Count = 0 Then
            oConn.CommitTrans
        Else
            oConn.RollbackTrans
            MsgBox "Transaksi Gagal, Ulangi Proses"
        End If
    End If
End Function

Public Function finishProses2(nOrder As String, stallNumber As String, dReason As String, nChassis As String, nTangl As String, nDay As String, timeStart As String, nActRt As Double, PRC As Boolean, timeFinish As String)
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim rsCari As New ADODB.Recordset
Dim nFO As String
Dim rsSearch As New ADODB.Recordset
Dim nameMech As String
Dim i As Integer
    
     
    'Set rsCari = oConn.Execute("SELECT * FROM STALL WHERE tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & stallNumber & "' AND DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeStart, "00:00") & "') < 0 AND FLAG IN ('','S','P','N','F')")
    'If Format(isnow, "HH:NN") > Format(timeStart, "00:00") Then
        'Set rsCari = oConn.Execute("SELECT * FROM STALL WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 or DATEDIFF(n,JAMMULAI, '" & Format(timeFinish, "00:00") & "') > 0) AND FLAG IN ('','S','P','N','F')")
        'Set rsCari = oConn.Execute("SELECT * FROM STALL WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0)) AND FLAG IN ('','S','P','N','F')")
        Set rsCari = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') < 0)) AND FLAG IN ('','S','P','N','F')")
    'Else
    '    Set rsCari = oConn.Execute("SELECT * FROM STALL WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND DATEDIFF(n,JAMMULAI, '" & Format(timeStart, "00:00") & "') < 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeFinish, "00:00") & "') > 0 AND FLAG IN ('','S','P','N','F')")
    'End If
    If Not rsCari.EOF Then
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
        Exit Function
    Else
        oConn.Errors.Clear
        oConn.BeginTrans
        
         
        Set rsSearch = oConn.Execute("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & Trim(nOrder) & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND noStall = '" & stallNumber & "' AND FLAGP <> 'M' ORDER BY MECH")
        If Not rsSearch.EOF Then
            i = 0
            Do While Not rsSearch.EOF
                i = i + 1
                nameMech = Trim(rsSearch("MECH"))
                
                Dim rsHlp3 As New ADODB.Recordset
                 
                Set rsHlp3 = oConn.Execute("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
                If Not rsHlp3.EOF Then
                    nFO = Trim(rsHlp3("GROUP"))
                End If
                rsHlp3.Close
            
                ' SIMPAN FINISH
                oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','F','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "','" & gSA & "' , '','" & nameMech & "')")
                If Trim(rsSearch("FLAGP")) = "" Then
                    oConn.Execute ("UPDATE tempProgress SET flagp = 'F', pause = CONVERT(varchar(5), GETDATE(),8), MECH = '" & nameMech & "', GRP = '" & Trim(nFO) & "' WHERE noorder = '" & nOrder & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND noStall = '" & stallNumber & "' AND FLAGP = ''")
                Else
                    oConn.Execute ("INSERT INTO tempProgress VALUES ('" & nOrder & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & stallNumber & "', '" & Format(DateAdd("n", -3, isnow), "HH:NN") & "', 'F','" & Format(isnow, "HH:NN") & "','" & nameMech & "','" & Trim(nFO) & "')")
                End If
                
                If i = 1 Then
                    Set rs = New ADODB.Recordset
                     
                    Set rs = oConn.Execute("SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
                    If Not rs.EOF Then
                        nSelisih = IIf(rs("SELISIH") > 0, rs("SELISIH"), 0)
                        totalRate = CDbl(nActRt) * 60
                        sisaRate = Abs(totalRate - CDbl(nSelisih))
                    End If
                    rs.Close
                    
                    ' FLAG FINISH
                    Set rs = New ADODB.Recordset
                     
                    Set rs = oConn.Execute("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
                    If Not rs.EOF Then
                        'oConn.Execute ("UPDATE STALL SET flag = 'F', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 20, Round((20 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 20, Round((20 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
                        oConn.Execute ("UPDATE STALL SET flag = 'F', RATE = " & Replace(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(rs("START")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(hitungRate(rs("START"), rs("PAUSE"))) < 10, Round((10 / 60), 1), Round(CDbl(hitungRate(rs("START"), rs("PAUSE"))) / 60, 1)), Format(rs("START"), "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
                    End If
                    rs.Close
                    
                    Dim namaMechAkhir As String
                    namaMechAkhir = Trim(nameMech)
                    Set rs = New ADODB.Recordset
                     
                    Set rs = oConn.Execute("SELECT TOP 1 A.MECH FROM TEMPPROGRESS A WITH(NOLOCK) INNER JOIN PROGRESS B WITH(NOLOCK) ON A.NOORDER = B.NOORDER AND A.NOSTALL = B.NOSTALL AND A.START = B.JAM AND A.MECH = B.MECH WHERE a.NOORDER = '" & Trim(nOrder) & "' AND FLAG NOT IN ('Q','A','O') GROUP BY A.MECH ORDER BY SUM(DATEDIFF(n,START,PAUSE)) DESC")
                    If Not rs.EOF Then
                        namaMechAkhir = Trim(rs("MECH"))
                    End If
                    rs.Close
                    oConn.Execute ("UPDATE jobdta SET time_S = '" & Format(isnow, "HH:NN") & "', rpr_cond = 'S', deliv = '" & IIf(PRC, "PRC", "OKE") & "', MECH = '" & Trim(namaMechAkhir) & "' WHERE no_chassis = '" & nChassis & "' AND no_wo = '" & nOrder & "' ")
                End If
                ' UPDATE JAM TERPAKAI MEKANIK
                oConn.Execute ("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI - " & Replace(CDbl(sisaRate), ",", ".") & ", ALOKASI = ALOKASI - " & Replace(CDbl(sisaRate), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & nameMech & "'")
                
                If i = 1 Then
                    ' INSERT JOB YG DIKERJAKAN KE TEMP
                    Dim nMinBaru As Integer
                    nMinBaru = 0
                     
                    Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
                    If Not rsCari.EOF Then
                        nMinBaru = CDbl(rsCari("SELISIH"))
                    Else
                        nMinBaru = nSelisih
                    End If
                    'nMinBaru = nMinBaru + (10 - (nMinBaru Mod 10))
                    rsCari.Close
                    'Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 1)
                    Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 3, Trim(stallNumber))
                End If
                rsSearch.MoveNext
            Loop
        End If
        
        If oConn.Errors.Count = 0 Then
            oConn.CommitTrans
        Else
            oConn.RollbackTrans
            MsgBox "Transaksi Gagal, Ulangi Proses"
        End If
    End If
End Function

Public Function moveProses(nOrder As String, nameMech As String, stallNumber As String, dReason As String, nChassis As String, nTangl As String, timeStart As String, nActRt As Double, timeFinish As String)
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim okProses As Boolean
Dim rsHlp As New ADODB.Recordset

    okProses = True
     
    'Set rsHlp = oConn.Execute("SELECT * FROM STALL WHERE tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeStart, "00:00") & "') < 0 AND FLAG IN ('','S','P','N','F')")
    Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND DATEDIFF(n,JAMMULAI, '" & Format(timeStart, "00:00") & "') < 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeFinish, "00:00") & "') > 0 AND FLAG IN ('','S','P','N','F')")
    If Not rsHlp.EOF Then
        okProses = False
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & vbCr & "No. Pol : " & Trim(Left(rsHlp("KETERANGAN"), InStr(rsHlp("KETERANGAN"), " |"))) & "  Jam : " & Trim(rsHlp("JAMMULAI")) & " - " & Trim(rsHlp("JAMSELESAI")) & vbCr & "Pindahkan unit Tersebut !"
    End If
    If rsHlp.State <> adStateClosed Then rsHlp.Close
    
    If okProses Then
        oConn.BeginTrans
        oConn.Errors.Clear
        Set rs = New ADODB.Recordset
         
        Set rs = oConn.Execute("SELECT ISNULL(SUM(DATEDIFF(n,START,PAUSE)),0) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
        If Not rs.EOF Then
            nSelisih = IIf(rs("SELISIH") > 0, rs("SELISIH"), 0)
            totalRate = CDbl(nActRt) * 60
            sisaRate = totalRate - CDbl(nSelisih)
            rateBaru = Round(CDbl(IIf(sisaRate < 20, 20, sisaRate)) / 60, 1)
        End If
        rs.Close
        
         
        Set rs = oConn.Execute("SELECT * FROM Stall WITH(NOLOCK) WHERE keterangan LIKE '%|" & nOrder & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "'")
        If Not rs.EOF Then
            ' UPDATE TEMPPROGRESS
            oConn.Execute ("UPDATE TEMPPROGRESS SET FLAGP = 'P', PAUSE = '" & Format(isnow, "HH:NN") & "', MECH = '" & nameMech & "' WHERE NOSTALL = '" & rs("NOSTALL") & "' AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' AND NOORDER = '" & Mid(rs("KETERANGAN"), InStr(rs("KETERANGAN"), " |") + 2, InStr(rs("KETERANGAN"), "| ") - (InStr(rs("KETERANGAN"), " |") + 2)) & "' AND FLAGP = ''")
            
            ' UPDATE STALL
            'JAMSELESAI = Left(timeFinish, 2) & ":" & Right(timeFinish, 2)
'            oConn.Execute ("UPDATE STALL SET noStall = '" & stallNumber & "', JamMulai = '" & Left(timeStart, 2) & ":" & Right(timeStart, 2) & "', JamSelesai = '" & Left(hitungWaktu(nActRt, Left(timeStart, 2) & ":" & Right(timeStart, 2), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "', Rate = '" & Replace(IIf(CDbl(nActRt) < 20, 0.3, Round(CDbl(nActRt) / 60, 1)), ",", ".") & "' " _
'                         & ", [Group] = '" & Trim(rs("GROUP")) & "', Mech = '" & rs("MECH") & "' WHERE Keterangan LIKE '%|" & nOrder & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "'")
                         
            oConn.Execute ("UPDATE STALL SET noStall = '" & stallNumber & "', JamMulai = '" & pembulatanJam(Format(timeStart, "00:00")) & "', JamSelesai = '" & pembulatanJam(Left(hitungWaktu(nActRt, Left(timeStart, 2) & ":" & Right(timeStart, 2), Format(isnow, "MM/DD/YYYY"), "progress"), 5)) & "', Rate = '" & Replace(IIf(CDbl(sisaRate) < 20, 0.3, Round(CDbl(sisaRate) / 60, 1)), ",", ".") & "' " _
                         & ", [Group] = '" & Trim(rs("GROUP")) & "', Mech = '" & rs("MECH") & "' WHERE Keterangan LIKE '%|" & nOrder & "|%' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "'")
        
            If Trim(rs("FLAG")) <> "" Then
                ' INSERT PROGRESS
                'oConn.Execute ("INSERT INTO progress VALUES ('" & nOrder & "', '" & nChassis & "', 'M', '" & stallNumber & "', '" & Format(nTangl, "MM/DD/YYYY") & "', '" & Format(isnow, "HH:NN") & "', '" & gSA & "', '','" & nameMech & "')")
                'oConn.Execute ("INSERT INTO TEMPPROGRESS VALUES ('" & nOrder & "', '" & Format(nTangl, "MM/DD/YYYY") & "', '" & stallNumber & "', '" & Format((Left(timeStart, 2) & ":" & Right(timeStart, 2)), "HH:NN") & "', 'M','" & Format((Left(timeStart, 2) & ":" & Right(timeStart, 2)), "HH:NN") & "', '" & nameMech & "')")
                oConn.Execute ("INSERT INTO TEMPPROGRESS VALUES ('" & nOrder & "', '" & Format(nTangl, "MM/DD/YYYY") & "', '" & stallNumber & "', '" & Format(isnow, "HH:NN") & "', 'M','" & Format(isnow, "HH:NN") & "', '" & nameMech & "','')")
            End If
            
            ' UPDATE FLAG HELPPROGRESS
'            Dim nMinBaru As Integer
'            Dim rsCari As New ADODB.Recordset
'            nMinBaru = 0
'
'            Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 ORDER BY TANGGAL DESC, START DESC")
'            If Not rsCari.EOF Then
'                nMinBaru = CDbl(rsCari("SELISIH"))
'            Else
'                nMinBaru = nSelisih
'            End If
'            rsCari.Close
'
'            Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 2)
        End If
        rs.Close
        
        If oConn.Errors.Count = 0 Then
            oConn.CommitTrans
        Else
            oConn.RollbackTrans
            MsgBox "Tidak berhasil dipindahkan !" & vbCr & "Keluar dari JPCB, ulangi proses."
        End If
    End If
End Function

Public Function startProses(nOrder As String, nameMech As String, stallNumber As String, dReason As String, nChassis As String, nTangl As String, nDay As String, timeStart As String, nActRt As Double, nStart1 As Boolean, timeFinish As String)
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim nFO As String
Dim rsHlp As New ADODB.Recordset
Dim okProses As Boolean
    
    okProses = True
     
    Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(nOrder) & "|%' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAG = 'S'")
    If Not rsHlp.EOF Then
        okProses = False
        MsgBox "Unit sudah diprogress User Lain ! "
    End If
    If rsHlp.State <> adStateClosed Then rsHlp.Close
    
     
    'Set rsHlp = oConn.Execute("SELECT * FROM STALL WHERE tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMMULAI, '" & Format(timeStart, "00:00") & "') < 0 AND FLAG IN ('','S','P','N','F')")
    'Set rsHlp = oConn.Execute("SELECT * FROM STALL WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 OR DATEDIFF(n,JAMMULAI, '" & Format(timeFinish, "00:00") & "') > 0) AND FLAG IN ('','S','P','N')")
    'Set rsHlp = oConn.Execute("SELECT * FROM STALL WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0)) AND FLAG IN ('','S','P','N','F')")
    Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') < 0)) AND FLAG IN ('','S','P','N','F')")
    If Not rsHlp.EOF Then
        okProses = False
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
    End If
    If rsHlp.State <> adStateClosed Then rsHlp.Close
    
    
    If okProses Then
        oConn.Errors.Clear
        oConn.BeginTrans
        
        Set rs = New ADODB.Recordset
         
        Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE NOSTALL = '" & Trim(stallNumber) & "' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAG = 'S'")
        If rsHlp.EOF Then
            Dim rsHlp3 As New ADODB.Recordset
             
            Set rsHlp3 = oConn.Execute("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
            If Not rsHlp3.EOF Then
                nFO = Trim(rsHlp3("GROUP"))
            End If
            rsHlp3.Close
            
            ' SIMPAN START
            oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','S','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "','" & gSA & "' , '', '" & nameMech & "')")
            oConn.Execute ("INSERT INTO tempProgress VALUES ('" & nOrder & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & stallNumber & "', '" & Format(isnow, "HH:NN") & "', '','','" & nameMech & "','" & Trim(nFO) & "')")
            
            ' UPDATE JAM TERPAKAI MEKANIK
             
            Set rs = oConn.Execute("SELECT ISNULL(SUM(DATEDIFF(n,START,PAUSE)),0) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
            If Not rs.EOF Then
                nSelisih = IIf(rs("SELISIH") > 0, rs("SELISIH"), 0)
                totalRate = CDbl(nActRt) * 60
                sisaRate = totalRate - CDbl(nSelisih)   'IIf(nStart1, 0, totalRate - CDbl(nSelisih))
        '        If sisaRate > 0 Then
        '            rateBaru = Round(CDbl(sisaRate) / 60, 1)
        '        Else
                    rateBaru = Round(CDbl(IIf(totalRate / 2 > 20, totalRate / 2, 20)) / 60, 1)
        '        End If
        
                totalRate = IIf(totalRate > 20, totalRate, 20)
            End If
            rs.Close
            
            ' FLAG START
            oConn.Execute ("UPDATE STALL SET flag = 'S', RATE = " & Replace(CDbl(rateBaru), ",", ".") & ", JAMMULAI = '" & pembulatanJam(Format(isnow, "HH:NN")) & "', JAMSELESAI = '" & Left(hitungWaktu(CDbl(rateBaru), Format(isnow, "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
            oConn.Execute ("UPDATE jobdta SET RPR_COND = 'P', MECH = '" & Trim(nameMech) & "' WHERE no_chassis = '" & nChassis & "' AND no_wo = '" & nOrder & "'")
            If nStart1 Then
                oConn.Execute ("UPDATE STALL SET flag = 'S', RATE = " & Replace(CDbl(Round(CDbl(totalRate) / 60, 1)), ",", ".") & ", JAMMULAI = '" & pembulatanJam(Format(isnow, "HH:NN")) & "', JAMSELESAI = '" & Left(hitungWaktu(CDbl(Round(CDbl(totalRate) / 60, 1)), Format(isnow, "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
                oConn.Execute ("UPDATE jobdta SET time_m = '" & Format(isnow, "HH:NN") & "', RPR_COND = 'P' WHERE no_chassis = '" & nChassis & "' AND no_wo = '" & nOrder & "' AND NOT EXISTS(SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & nOrder & "' AND NO_CHASSIS = '" & nChassis & "' AND TIME_M <> '')")
            End If
            
            oConn.Execute ("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI + " & Replace(CDbl(sisaRate), ",", ".") & ", ALOKASI = ALOKASI + " & Replace(CDbl(sisaRate), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & nameMech & "'")
            
            ' UPDATE FLAG HELPPROGRESS
            Dim nMinBaru As Integer
            Dim rsCari As New ADODB.Recordset
            nMinBaru = 0
             
            Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
            If Not rsCari.EOF Then
                nMinBaru = CDbl(rsCari("SELISIH"))
            Else
                nMinBaru = nSelisih
            End If
            'nMinBaru = nMinBaru + (10 - (nMinBaru Mod 10))
            rsCari.Close
            Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 2, Trim(stallNumber))
        Else
            MsgBox "Stall Dipakai unit lain. " & vbCr & "Kendaraan dengan No. Polisi : " & Left(Trim(rsHlp("KETERANGAN")), InStr(Trim(rsHlp("KETERANGAN")), " |"))
        End If
        
        If oConn.Errors.Count = 0 Then
            oConn.CommitTrans
        Else
            oConn.RollbackTrans
            MsgBox "Proses tidak berhasil ! " & vbCr & "Refresh JPCB dan ulangi Proses"
        End If
    End If
End Function

Public Function startInspection(nOrder As String, nameMech As String, stallNumber As String, dReason As String, nChassis As String, nTangl As String, nDay As String, timeStart As String, nActRt As Double, timeFinish As String)
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim nFO As String
Dim rsHlp As New ADODB.Recordset
Dim okProses As Boolean
    
    okProses = True
     
    Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(nOrder) & "|%' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAG = 'Q'")
    If Not rsHlp.EOF Then
        okProses = False
        MsgBox "Unit sudah diprogress User Lain ! "
    End If
    If rsHlp.State <> adStateClosed Then rsHlp.Close
    
     
    Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') < 0)) AND FLAG IN ('','S','P','N','F')")
    If Not rsHlp.EOF Then
        okProses = False
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
    End If
    If rsHlp.State <> adStateClosed Then rsHlp.Close
    
    
    If okProses Then
        oConn.Errors.Clear
        oConn.BeginTrans
        
        Set rs = New ADODB.Recordset
         
        Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE NOSTALL = '" & Trim(stallNumber) & "' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAG = 'S'")
        If rsHlp.EOF Then
            Dim rsHlp3 As New ADODB.Recordset
             
            Set rsHlp3 = oConn.Execute("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
            If Not rsHlp3.EOF Then
                nFO = Trim(rsHlp3("GROUP"))
            End If
            rsHlp3.Close
            
            ' SIMPAN START
            oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','Q','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "','" & gSA & "' , 'PROSES FINAL INSPECTION', '" & gSA & "')")
            oConn.Execute ("INSERT INTO tempProgress VALUES ('" & nOrder & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & stallNumber & "', '" & Format(isnow, "HH:NN") & "', '','','" & gSA & "','" & Trim(nFO) & "')")
                        
            ' FLAG START
            oConn.Execute ("UPDATE STALL SET flag = 'Q', RATE = " & Replace(CDbl(5 / 60), ",", ".") & ", JAMMULAI = '" & pembulatanJam(Format(isnow, "HH:NN")) & "', JAMSELESAI = '" & Left(hitungWaktu(CDbl(rateBaru), Format(isnow, "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
            oConn.Execute ("UPDATE jobdta SET RPR_COND = 'P' WHERE no_chassis = '" & nChassis & "' AND no_wo = '" & nOrder & "'")
            
            ' UPDATE FLAG HELPPROGRESS
            Dim nMinBaru As Integer
            Dim rsCari As New ADODB.Recordset
            nMinBaru = 0
             
            Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
            If Not rsCari.EOF Then
                nMinBaru = CDbl(rsCari("SELISIH"))
            Else
                nMinBaru = nSelisih
            End If
            rsCari.Close
            'Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 2, Trim(stallNumber))
        Else
            MsgBox "Stall Dipakai unit lain. " & vbCr & "Kendaraan dengan No. Polisi : " & Left(Trim(rsHlp("KETERANGAN")), InStr(Trim(rsHlp("KETERANGAN")), " |"))
        End If
        
        If oConn.Errors.Count = 0 Then
            oConn.CommitTrans
        Else
            oConn.RollbackTrans
            MsgBox "Proses tidak berhasil ! " & vbCr & "Refresh JPCB dan ulangi Proses"
        End If
    End If
End Function

Public Function startWashing(nOrder As String, nameMech As String, stallNumber As String, dReason As String, nChassis As String, nTangl As String, nDay As String, timeStart As String, nActRt As Double, timeFinish As String)
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim nFO As String
Dim rsHlp As New ADODB.Recordset
Dim okProses As Boolean
    
    okProses = True
     
    Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(nOrder) & "|%' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAG = 'W'")
    If Not rsHlp.EOF Then
        okProses = False
        MsgBox "Unit sudah diprogress User Lain ! "
    End If
    If rsHlp.State <> adStateClosed Then rsHlp.Close
    
     
    Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE KETERANGAN NOT LIKE '%|" & Trim(nOrder) & "|%' AND tanggal = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(stallNumber) & "' AND ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') > 0) or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') > 0 AND DATEDIFF(n,JAMSELESAI, '" & Format(timeFinish, "00:00") & "') < 0)) AND FLAG IN ('','S','P','N','F')")
    If Not rsHlp.EOF Then
        okProses = False
        MsgBox "Ada Unit lain yang akan tertutup !!" & vbCr & "Pindahkan unit Tersebut !"
    End If
    If rsHlp.State <> adStateClosed Then rsHlp.Close
    
    If okProses = False Then
        MsgBox "Stall Dipakai unit lain. " & vbCr & "Kendaraan dengan No. Polisi : " & Left(Trim(rsHlp("KETERANGAN")), InStr(Trim(rsHlp("KETERANGAN")), " |"))
        Exit Function
    End If
    If okProses Then
        oConn.Errors.Clear
        oConn.BeginTrans
        
        Set rs = New ADODB.Recordset
         
        Set rsHlp = oConn.Execute("SELECT * FROM STALL WITH(NOLOCK) WHERE NOSTALL = '" & Trim(stallNumber) & "' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAG = 'S'")
        If rsHlp.EOF Then
            Dim rsHlp3 As New ADODB.Recordset
             
            Set rsHlp3 = oConn.Execute("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(nameMech) & "'")
            If Not rsHlp3.EOF Then
                nFO = Trim(rsHlp3("GROUP"))
            End If
            rsHlp3.Close
            
            ' SIMPAN START
            oConn.Execute ("INSERT INTO Progress VALUES ('" & nOrder & "','" & nChassis & "','A','" & stallNumber & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "','" & gSA & "' , 'PROSES FINAL INSPECTION', '" & gSA & "')")
            oConn.Execute ("INSERT INTO tempProgress VALUES ('" & nOrder & "','" & Format(nTangl, "MM/DD/YYYY") & "','" & stallNumber & "', '" & Format(isnow, "HH:NN") & "', '','','" & gSA & "','" & Trim(nFO) & "')")
                        
            ' FLAG START
            oConn.Execute ("UPDATE STALL SET flag = 'A', RATE = " & Replace(CDbl(5 / 60), ",", ".") & ", JAMMULAI = '" & pembulatanJam(Format(isnow, "HH:NN")) & "', JAMSELESAI = '" & Left(hitungWaktu(CDbl(rateBaru), Format(isnow, "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "' WHERE noStall = '" & stallNumber & "' AND tanggal = '" & Format(nTangl, "MM/DD/YYYY") & "' AND keterangan like '%|" & nOrder & "|%'")
            oConn.Execute ("UPDATE jobdta SET RPR_COND = 'P' WHERE no_chassis = '" & nChassis & "' AND no_wo = '" & nOrder & "'")
            
            ' UPDATE FLAG HELPPROGRESS
'            Dim nMinBaru As Integer
'            Dim rsCari As New ADODB.Recordset
'            nMinBaru = 0
'
'            Set rsCari = oConn.Execute("SELECT DATEDIFF(n,START,PAUSE) AS SELISIH FROM TEMPPROGRESS WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0 AND TANGGAL = '" & Format(nTangl, "MM/DD/YYYY") & "' ORDER BY TANGGAL DESC, START DESC, PAUSE DESC")
'            If Not rsCari.EOF Then
'                nMinBaru = CDbl(rsCari("SELISIH"))
'            Else
'                nMinBaru = nSelisih
'            End If
'            rsCari.Close
'            Call saveTemp(nOrder, nMinBaru, Left(timeStart, 2) & ":" & Right(timeStart, 2), 2, Trim(stallNumber))
'        Else
'            MsgBox "Stall Dipakai unit lain. " & vbCr & "Kendaraan dengan No. Polisi : " & Left(Trim(rsHlp("KETERANGAN")), InStr(Trim(rsHlp("KETERANGAN")), " |"))
        End If
        
        If oConn.Errors.Count = 0 Then
            oConn.CommitTrans
        Else
            oConn.RollbackTrans
            MsgBox "Proses tidak berhasil ! " & vbCr & "Refresh JPCB dan ulangi Proses"
        End If
    End If
End Function


Public Function saveTemp(nWO As String, nSelisih As Integer, jMulai As String, idx As Integer, nStall As String)
    
    On Error GoTo errHandler
    If idx = 1 Then
        ' INSERT
        oConn.Execute ("INSERT INTO HELPPROGRESS SELECT NOSTALL, TANGGAL, '" & jMulai & "', '" & Replace(Round(nSelisih / 60, 1), ",", ".") & "', 'Q', KETERANGAN, MECH FROM STALL WITH(NOLOCK) WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND KETERANGAN LIKE '%|" & nWO & "|%'")
    ElseIf idx = 2 Then
        ' UPDATE
        oConn.Execute ("UPDATE HELPPROGRESS SET FLAG = 'X', RATE = '" & Replace(Round(nSelisih / 60, 2), ",", ".") & "' WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND KETERANGAN LIKE '%|" & nWO & "|%' AND NOSTALL = '" & nStall & "' AND JAMMULAI = '" & jMulai & "'")
    ElseIf idx = 3 Then
        ' UPDATE
        oConn.Execute ("UPDATE HELPPROGRESS SET FLAG = 'X', RATE = '" & Replace(Round(nSelisih / 60, 2), ",", ".") & "' WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND KETERANGAN LIKE '%|" & nWO & "|%' AND NOSTALL = '" & nStall & "' AND JAMMULAI = '" & jMulai & "'")
        ' INSERT
        oConn.Execute ("INSERT INTO HELPPROGRESS SELECT NOSTALL, TANGGAL, '" & jMulai & "', '" & Replace(Round(nSelisih / 60, 1), ",", ".") & "', 'Q', KETERANGAN, MECH FROM STALL WITH(NOLOCK) WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND KETERANGAN LIKE '%|" & nWO & "|%'")
    End If
    Exit Function

errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function newRate(nOrder As String) As Double
Dim rateBaru As Double
Dim totalRate As Integer
Dim sisaRate As Integer
Dim nSelisih As Integer
Dim rsActualRate As New ADODB.Recordset
    
    Set rs = New ADODB.Recordset
        
    ' UPDATE JAM TERPAKAI MEKANIK
     
    Set rs = oConn.Execute("SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
    If Not rs.EOF Then
        nSelisih = IIf(rs("SELISIH") > 0, rs("SELISIH"), 0)
        totalRate = CDbl(takeActualRate(nOrder)) * 60
        sisaRate = totalRate - CDbl(nSelisih)
        If sisaRate > 0 Then
            newRate = Round(CDbl(sisaRate) / 60, 1)
        Else
            newRate = Round(CDbl(totalRate) / 60, 1)
        End If
    End If
    rs.Close
End Function

Public Function takeActualRate(numberOrder As String) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT ACTRATE1+ACTRATE2+ACTRATE3+ACTRATE4+ACTRATE5+ACTRATE6+ACTRATE7+ACTRATE8 AS TOTaCT FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & numberOrder & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        takeActualRate = CDbl(rs("TOTACT"))
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchMechanic(mechanicName As String, nameDay As String, dateDay As String, hourDay As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "select A.MECH,[NAME],[GROUP], LEVMECH, MENITKERJA, TERPAKAI from mekanik A WITH(NOLOCK) LEFT JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH where A.MECH = '" & mechanicName & "' AND hari = '" & nameDay & "' AND TANGGAL = '" & dateDay & "' AND (MENITKERJA-TERPAKAI) > 0 AND MASUK = 1 AND JAMMULAI <= '" & hourDay & "' AND JAMSELESAI >= '" & hourDay & "' order by A.MECH"

    searchMechanic = False
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        searchMechanic = True
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

