VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsKPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

''=============================================================================================''
''ooooooooooo000ooooooooooo KPI TAM + P Anung by unit JT bln          ooooooooooo000ooooooooooo''
''=============================================================================================''

Public Function addFilter(Bln As Integer, Thn As Integer) As String
Dim txtFilter As String
        
    txtFilter = ""
    txtFilter = "select DISTINCT NORANGKA " & _
                "FROM M_Estimasi a WITH(NOLOCK) INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
                "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis WHERE a.JT_Tgl is Not Null AND " & _
                "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
                "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "' "
    txtFilter = txtFilter + "UNION Select DISTINCT NoRangka FROM " & Log.LogSPK & "..T_MRS a WITH(NOLOCK) LEFT JOIN " & Log.LogSPK & "..AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = 1"
    
    addFilter = txtFilter
    
    Call Add_Filter(txtFilter)
End Function


Private Sub Add_Filter(sqlString)
    On Error GoTo errHandler
        Dim cm As New Command
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        Call Delete_Filter("Delete FROM FilterKPI_1")
        
        cm.CommandText = "Insert Into FilterKPI_1 " & sqlString
        cm.CommandType = adCmdText
        Set cm.ActiveConnection = mConn
        cm.Execute
        Set cm = Nothing
        Exit Sub
        
errHandler:
        If Not cm Is Nothing Then Set cm = Nothing
            err.Raise err.Number, err.Source, err.Description
End Sub
Private Sub Delete_Filter(sqlString)
    On Error GoTo errHandler
        Dim cm As New Command
        
        cm.CommandText = sqlString
        cm.CommandType = adCmdText
        Set cm.ActiveConnection = mConn
        cm.Execute
        Set cm = Nothing
        Exit Sub
        
errHandler:
        If Not cm Is Nothing Then Set cm = Nothing
            err.Raise err.Number, err.Source, err.Description
End Sub


Public Function KPI_1(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim Ket As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    If Bln = 1 Then
        Ket = "MONTH(entryCr1) = '12' AND YEAR(entryCr1) = '" & CDbl(Thn - 1) & "'"
    Else
        Ket = "MONTH(entryCr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr1) = '" & CDbl(Thn) & "'"
    End If
    If a = True Then
        strSql = "Select Count(*) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO Where A.BATAL = 0 and " & Ket & ""
    Else
        strSql = "Select Count(*) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO Where OkCR1 = 1 and A.BATAL = 0 and " & Ket & ""
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_1 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_1c(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    If a = True Then
        strSql = "Select Count(*) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO Where Month(entryCR1)= '" & Bln & "' And Year(entryCR1)='" & Thn & "' and (Tgl_BPK > '" & Format(TglYYYYMMDD(GetTglAkhir(Bln, Thn, 0)), "MM/DD/YYYY") & "' OR Tgl_BPK < '" & Format(Bln, "00") & "/" & "01/" & Format(Thn, "0000") & "') And OkCR1 = 1 AND A.BATAL = 0"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_1c = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_2(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim Ket As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    If Bln = 1 Then
        Ket = "(MONTH(entryCr2) = '12' AND YEAR(entryCr2) = '" & CDbl(Thn - 1) & "')) and MONTH(entryCr1) = '12' AND YEAR(entryCr1) = '" & CDbl(Thn - 1) & "'"
    Else
        Ket = "(MONTH(entryCr2) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr2) = '" & CDbl(Thn) & "')) and MONTH(entryCr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr1) = '" & CDbl(Thn) & "'"
    End If
    If a = True Then
        strSql = "Select Count(*) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO Where OkCR1 = 1 AND A.BATAL = 0 AND ISNULL(TGLCR2,'') <> '' and ((Month(entryCR2)= '" & Bln & "' And Year(entryCR2)='" & Thn & "') or " & Ket
    Else
        strSql = "Select Count(*) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO Where OkCR1 = 1 AND A.BATAL = 0 AND ISNULL(TGLCR2,'') <> '' and Month(entryCR2)= '" & Bln & "' AND (Tgl_BPK > '" & Format(TglYYYYMMDD(GetTglAkhir(Bln, Thn, 0)), "MM/DD/YYYY") & "' OR Tgl_BPK < '" & Format(Bln, "00") & "/" & "01/" & Format(Thn, "0000") & "') And Year(entryCR2)='" & Thn & "'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_2 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_3(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String

    On Error GoTo errHandler
    If Bln = 1 Then
        txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
    Else
        txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
    End If
    
    If a = True Then
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        'strSQL = "Select Count(BPK_No) FROM BPK Where Month(Tgl_BPK)= '" & Bln & "' And Year(Tgl_BPK)='" & Thn & "'"
        strSql = "Select Count(*) FROM T_MRS A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = '1'"
    Else
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        strSql = "Select Count(DISTINCT LastNoWo) FROM M_ListReminder WITH(NOLOCK) Where ((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') or " & txtHelp & ") And Tipe = '1K' And " & _
                 "NoRangka in (Select NoRangka FROM " & Log.LogSPK & "..T_MRS A WITH(NOLOCK) INNER JOIN " & Log.LogSPK & "..AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = '1')"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If a = True Then
        rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Else
        rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    End If
    Set rs.ActiveConnection = Nothing
    Set KPI_3 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_3a(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String
    On Error GoTo errHandler
    If Bln = 1 Then
        txtHelp = "(MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "')"
    Else
        txtHelp = "(MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "')"
    End If
    If a = True Then
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        strSql = ""
    Else
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        strSql = "Select Count(DISTINCT LastNoWo) FROM M_ListReminderCall WITH(NOLOCK) Where ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') or " & txtHelp & ") And KmNextSvc = '1000' " & _
                 "And NoRangka in (Select NoRangka FROM " & Log.LogSPK & "..T_MRS A WITH(NOLOCK) INNER JOIN " & Log.LogSPK & "..AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = '1') " & _
                 "AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder WITH(NOLOCK) WHERE Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "' And Tipe = '1K')"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If a = True Then
        rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Else
        rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    End If
    Set rs.ActiveConnection = Nothing
    Set KPI_3a = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_3sales(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String
Dim Ket As String

    On Error GoTo errHandler
    If Bln = 1 Then
        Ket = "MONTH(entryCr1) = '12' AND YEAR(entryCr1) = '" & CDbl(Thn - 1) & "'"
        txtHelp = "month(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '12' " & _
                  "AND YEAR(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '" & CDbl(Thn - 1) & "'"
    Else
        Ket = "MONTH(entryCr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr1) = '" & CDbl(Thn) & "'"
        txtHelp = "month(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '" & CDbl(Bln - 1) & "' " & _
                  "AND YEAR(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '" & CDbl(Thn) & "'"
    End If
    
    If a = True Then
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        'strSQL = "Select Count(*) FROM T_MRS A INNER JOIN AFI B ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' "
'        strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO INNER JOIN SPK C ON A.SPK_NO = C.SPK_NO " & _
'                 "INNER JOIN INKATEGORY D ON C.KODE_TYPE = D.KODEKATEGORY Where A.BATAL = 0 And " & _
'                 "month(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END), tglcr1)) = '" & Bln & "' AND YEAR(dateadd(d,TglCR2,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END), tglcr1)) = '" & Thn & "' and " & ket & ""
        strSql = "Select count(*) FROM AFI A WITH(NOLOCK) left join summary b WITH(NOLOCK) on a.spk_no=b.no_bukti and b.tipetr='spk' and b.status=1 " & _
                 "left join inmst c WITH(NOLOCK) on b.norangka=c.kodeitem left join inkategory d WITH(NOLOCK) on c.kodekategori=d.kodekategory Where A.BATAL = 0 " & _
                 "And ((month(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '" & Bln & "' " & _
                 "AND YEAR(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '" & Thn & "') or (" & txtHelp & ") ) and " & Ket & ""
    Else
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        strSql = "Select Count(*) FROM T_MRS A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND ISNULL(REMINDERSALES,'0') = '1'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If a = True Then
        rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Else
        rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    End If
    Set rs.ActiveConnection = Nothing
    Set KPI_3sales = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_4(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String
Dim txtHelp2 As String

    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If a = True Then
        If Bln = 1 Then
            txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
            txtHelp2 = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
            txtHelp2 = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        End If
        strSql = "Select Count(DISTINCT Lastnowo) From M_ListRemindercall WITH(NOLOCK) Where K_DM = 1 And ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') Or (" & txtHelp & ")) and NoRangka in (" & addFilter(Bln, Thn) & ") and lastnowo in (Select DISTINCT Lastnowo From M_ListReminder WITH(NOLOCK) Where ((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') Or (" & txtHelp2 & ")) And Tipe = 'DM')"
    Else
        If Bln = 1 Then
            txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        End If
        strSql = "Select Count(DISTINCT Lastnowo) From M_ListReminder WITH(NOLOCK) Where ((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') Or (" & txtHelp & ")) And Tipe = 'DM' AND NoRangka IN (" & addFilter(Bln, Thn) & ")"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_4 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_5(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String

    On Error GoTo errHandler
    If Bln = 1 Then
        txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
    Else
        txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
    End If
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    Dim nFil As String
    nFil = addFilter(Bln, Thn)
    If a = True Then
        '''strSQL = "Select Count(DISTINCT Lastnowo) AS xx From M_ListRemindercall Where K_Telp = 1 And ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') Or (" & txtHelp & ")) And NoRangka in (" & addFilter(Bln, Thn) & ")"
        strSql = "Select Count(DISTINCT Lastnowo) AS xx From M_ListRemindercall WITH(NOLOCK) Where K_Telp = 1 And ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') Or (" & txtHelp & ")) And NoRangka in (Select NoRangka FROM FilterKPI_1)"
    Else
        '''strSQL = "Select Count(DISTINCT Lastnowo) AS xx From M_ListRemindercall Where ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') Or (" & txtHelp & ")) And NoRangka in (" & addFilter(Bln, Thn) & ")"
        strSql = "Select Count(DISTINCT Lastnowo) AS xx From M_ListRemindercall WITH(NOLOCK) Where ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') Or (" & txtHelp & ")) And NoRangka in (Select NoRangka FROM FilterKPI_1)"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_5 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_6(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String

    On Error GoTo errHandler
    If Bln = 1 Then
        txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
    Else
        txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
    End If
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If a = True Then
        strSql = "select COUNT(DISTINCT NORANGKA) " & _
                 "FROM M_Estimasi a WITH(NOLOCK) INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NOINV = B.NOINV AND B.FS_LOG_DEL = '' " & _
                 "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis WHERE a.JT_Tgl is Not Null AND " & _
                 "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
                 "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "' " & _
                 "AND NORANGKA NOT IN (SELECT DISTINCT NO_CHASSIS FROM JOBDTA X WITH(NOLOCK) INNER JOIN T_REGULAR Y WITH(NOLOCK) ON X.FS_RC = Y.FS_RC " & _
                 "LEFT JOIN M_ESTIMASI H WITH(NOLOCK) ON X.NO_CHASSIS = H.NORANGKA AND H.SBENEXT = Y.FS_NILAI LEFT JOIN M_LISTREMINDER W WITH(NOLOCK) ON X.NO_CHASSIS = W.NORANGKA AND H.KMNEXT = W.KMNEXTSVC " & _
                 "LEFT JOIN M_LISTREMINDERCALL Q WITH(NOLOCK) ON X.NO_CHASSIS = Q.NORANGKA AND H.KMNEXT = Q.KMNEXTSVC WHERE NO_CHASSIS IN ( " & _
                 "SELECT NORANGKA FROM M_Estimasi a WITH(NOLOCK) INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis WHERE a.JT_Tgl is Not Null AND " & _
                 "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
                 "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "' " & _
                 "AND DATEDIFF(d,X.DATE,(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)) > 0 " & _
                 ") AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND ISNULL(TGLPROSES,'') = '' AND ISNULL(Q.TGLFOLLOW,'') = '' AND FS_NILAI = A.SBENEXT) "
        
        If Log.RPLHistori = 1 Then
            strSql = UCase(strSql)
            strSql = Replace(strSql, "T_REGULAR", "[" & dbReplikasiHO & "]..T_REGULAR")
        End If
    Else
        strSql = "Select Count(DISTINCT LastNowo) From M_ListReminder WITH(NOLOCK) Where ((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') Or (" & txtHelp & ")) And NoRangka in (" & addFilter(Bln, Thn) & ")"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_6 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_6a(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler

    If a = True Then
    Call Get_FilterKPI_6a(Bln, Thn)
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        strSql = "Select Count(NoSPK) FROM T_MRS a  WITH(NOLOCK) LEFT JOIN AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = 1 " & _
                 "AND NORANGKA NOT IN ( SELECT DISTINCT NO_CHASSIS FROM " & Log.LogBengkel & "..JOBDTA X WITH(NOLOCK) INNER JOIN " & Log.LogBengkel & "..T_REGULAR Y WITH(NOLOCK) ON X.FS_RC = Y.FS_RC " & _
                 "LEFT JOIN " & Log.LogBengkel & "..M_LISTREMINDER W WITH(NOLOCK) ON X.NO_CHASSIS = W.NORANGKA AND Y.FS_NILAI = W.KMNEXTSVC " & _
                 "LEFT JOIN " & Log.LogBengkel & "..M_LISTREMINDERCALL Q WITH(NOLOCK) ON X.NO_CHASSIS = Q.NORANGKA AND Y.FS_NILAI = Q.KMNEXTSVC " & _
                 "WHERE NO_CHASSIS IN ( SELECT NORANGKA FROM T_MRS a WITH(NOLOCK) LEFT JOIN AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = 1 AND DATEDIFF(d,X.DATE,TGL1) > 0 " & _
                 ") AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND ISNULL(TGLPROSES,'') = '' AND ISNULL(Q.TGLFOLLOW,'') = '' AND FS_NILAI = '1000')"
                 
        strSql = "Select Count(NoSPK) FROM T_MRS a WITH(NOLOCK) LEFT JOIN AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = 1 " & _
                 "AND NORANGKA NOT IN (Select NoRangka FROM " & Log.LogBengkel & "..FilterKPI_6a )"
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_6a = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Private Sub Get_FilterKPI_6a(Bln As Integer, Thn As Integer)
    On Error GoTo errHandler
        Dim cm As New Command
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        Call Delete_FilterKPI_6a
        
        strSql = "Insert Into FilterKPI_6a (NoRangka) " & _
                 "(SELECT DISTINCT NO_CHASSIS FROM " & Log.LogBengkel & "..JOBDTA X WITH(NOLOCK) INNER JOIN " & Log.LogBengkel & "..T_REGULAR Y WITH(NOLOCK) ON X.FS_RC = Y.FS_RC " & _
                 "LEFT JOIN " & Log.LogBengkel & "..M_LISTREMINDER W WITH(NOLOCK) ON X.NO_CHASSIS = W.NORANGKA AND Y.FS_NILAI = W.KMNEXTSVC " & _
                 "LEFT JOIN " & Log.LogBengkel & "..M_LISTREMINDERCALL Q WITH(NOLOCK) ON X.NO_CHASSIS = Q.NORANGKA AND Y.FS_NILAI = Q.KMNEXTSVC " & _
                 "WHERE NO_CHASSIS IN ( SELECT NORANGKA FROM " & Log.LogSPK & "..T_MRS a WITH(NOLOCK) LEFT JOIN " & Log.LogSPK & "..AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = 1 AND DATEDIFF(d,X.DATE,TGL1) > 0 " & _
                 ") AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND ISNULL(TGLPROSES,'') = '' AND ISNULL(Q.TGLFOLLOW,'') = '' AND FS_NILAI = '1000')"
                 
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

Private Sub Delete_FilterKPI_6a()
    On Error GoTo errHandler
    
        Dim cm As New Command
                strSql = "Delete From FilterKPI_6a"
        
        cm.CommandText = strSql
        cm.CommandType = adCmdText
        Set cm.ActiveConnection = mConn
        cm.Execute
        Set cm = Nothing
        Exit Sub
        
errHandler:
        If Not cm Is Nothing Then Set cm = Nothing
            err.Raise err.Number, err.Source, err.Description
End Sub

Public Function KPI_6b(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String

    On Error GoTo errHandler
    If Bln = 1 Then
        txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
    Else
        txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
    End If
    
    
    
    'mConn.ConnectionTimeout = 3600
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    If a = True Then
''''        strSQL = "Select Count(DISTINCT LastNowo) From M_ListReminderCall Where ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)= '" & Thn & "') Or (" & txtHelp & ")) AND LastNoWo NOT IN (SELECT LastNoWo FROM M_LISTREMINDER) " & _
''''                 "And NoRangka in (" & addFilter(Bln, Thn) & ")"
'''         strSQL = "Select Count(DISTINCT LastNowo) From M_ListReminderCall Where ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)= '" & Thn & "') Or (" & txtHelp & ")) AND LastNoWo NOT IN (SELECT LastNoWo FROM M_LISTREMINDER Where ((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') Or (" & txtHelp & ")) And NoRangka in (" & addFilter(Bln, Thn) & ")) " & _
'''                 "And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ' 6 Agustus 2012
        
        Call Add_Temp_LastNoWO(Bln, Thn, txtHelp)
        Call Add_Temp_NoRangka1(Bln, Thn, txtHelp)
        Call Add_Temp_NoRangka2(Bln, Thn, txtHelp)
        
        strSql = "Select Count(DISTINCT LastNowo) From M_ListReminderCall WITH(NOLOCK) " & _
                 "Where ((Month(TglFollow)= '8' And Year(TglFollow)= '2012') " & _
                 "Or (MONTH(TglFollow) = '7' AND YEAR(TglFollow) = '2012')) " & _
                 "AND LastNoWo NOT IN " & _
                 "(SELECT LastNoWo from Temp_LastNoWo) " & _
                 "And NoRangka in " & _
                 "(Select NoRangka FROM Temp_NoRangka1 UNION Select NoRangka FROM Temp_NoRangka2)"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_6b = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

' 6 Agustus 2012
Private Sub Add_Temp_LastNoWO(Bln As Integer, Thn As Integer, param As String)
    On Error GoTo errHandler
        Call Delete_Temp(0)
        Dim cm As New Command
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        
        strSql = "Insert Into Temp_LastNoWo (LastNoWo) " & _
                 "(SELECT LastNoWo FROM M_LISTREMINDER Where " & _
                 "((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') Or (" & param & ")))"
                 
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

Private Sub Add_Temp_NoRangka1(Bln As Integer, Thn As Integer, param As String)
    On Error GoTo errHandler
        Call Delete_Temp(1)
        Dim cm As New Command
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        
        strSql = "Insert Into Temp_NoRangka1 (NoRangka) " & _
                 "(select DISTINCT NORANGKA FROM M_Estimasi a WITH(NOLOCK) INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
                 "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis WHERE a.JT_Tgl is Not Null AND " & _
                 "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
                 "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "')"
                 
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

Private Sub Add_Temp_NoRangka2(Bln As Integer, Thn As Integer, param As String)
    On Error GoTo errHandler
        Call Delete_Temp(2)
        Dim cm As New Command
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        
        strSql = "Insert Into Temp_NoRangka2 (NoRangka) " & _
                 "(Select DISTINCT NoRangka FROM " & Log.LogSPK & "..T_MRS a WITH(NOLOCK) LEFT JOIN " & Log.LogSPK & "..AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO " & _
                 "Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = 1)"
                 
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

Private Sub Delete_Temp(isDel As Integer)
    On Error GoTo errHandler
    
        Dim cm As New Command
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        
        Select Case isDel
            Case 0
                strSql = "Delete From Temp_LastNoWo"
            Case 1
                strSql = "Delete From Temp_NoRangka1"
            Case 2
                strSql = "Delete From Temp_NoRangka2"
        End Select
        
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

'============

Public Function KPI_8(Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String

    On Error GoTo errHandler
    If Bln = 1 Then
        txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
    Else
        txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
    End If
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    strSql = "SELECT COUNT(DISTINCT LASTNOWO) FROM BOOK A WITH(NOLOCK) INNER JOIN M_LISTREMINDERCALL B WITH(NOLOCK) ON A.NO_CHASSIS = B.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,TGLFOLLOW, DATE_BOOK) = 0 AND ((MONTH(TGLFOLLOW) = '" & Bln & "' AND YEAR(TGLFOLLOW) = '" & Thn & "') Or (" & txtHelp & ")) And NoRangka in (" & addFilter(Bln, Thn) & ") "
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_8 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_9(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String
    On Error GoTo errHandler
    
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    Dim nFil As String
    nFil = addFilter(Bln, Thn)
    If a Then
        If Bln = 1 Then
            txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
        End If
        ''' asli '''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -7 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -14 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        '''strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -21 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -30 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        strSql = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C WITH(NOLOCK) ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -21 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (Select NoRangka From FilterKPI_1) AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        
    Else
        If Bln = 1 Then
            txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        End If
        ''' asli '''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -7 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -14 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        '''strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -21 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -30 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        strSql = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C WITH(NOLOCK) ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -21 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (Select NoRangka FROM FilterKPI_1)"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_9 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_10(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    Dim txtHelp As String
    Dim txtHelp1 As String
    
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If Bln = 1 Then
        txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
        txtHelp1 = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
    Else
        txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
        txtHelp1 = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
    End If
    
    If a Then
'        strSQL = "SELECT COUNT(DISTINCT NOFAK) FROM BOOK A WHERE (NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDER WHERE ((MONTH(TGLPROSES) = '" & Bln & "' AND YEAR(TGLPROSES) = " & Thn & ") OR (" & txtHelp1 & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) <= 7  And NoRangka in (" & addFilter(Bln, Thn) & ")) " & _
'                 "OR NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDERCALL WHERE ((MONTH(TGLFOLLOW) = '" & Bln & "' AND YEAR(TGLFOLLOW) = " & Thn & ") OR (" & txtHelp & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) <= 7 And NoRangka in (" & addFilter(Bln, Thn) & "))) "
'        strSQL = "SELECT COUNT(DISTINCT NOFAK) FROM BOOK A WHERE (NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDER WHERE ((MONTH(TGLPROSES) = '" & Bln & "' AND YEAR(TGLPROSES) = " & Thn & ") OR (" & txtHelp1 & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) <= 14  And NoRangka in (" & addFilter(Bln, Thn) & ")) " & _
'                 "OR NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDERCALL WHERE ((MONTH(TGLFOLLOW) = '" & Bln & "' AND YEAR(TGLFOLLOW) = " & Thn & ") OR (" & txtHelp & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) <= 14 And NoRangka in (" & addFilter(Bln, Thn) & "))) "
        strSql = "SELECT COUNT(DISTINCT NOFAK) FROM BOOK A WITH(NOLOCK) WHERE (NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDER WITH(NOLOCK) WHERE ((MONTH(TGLPROSES) = '" & Bln & "' AND YEAR(TGLPROSES) = " & Thn & ") OR (" & txtHelp1 & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) <= 21  And NoRangka in (" & addFilter(Bln, Thn) & ")) " & _
                 "OR NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDERCALL WITH(NOLOCK) WHERE ((MONTH(TGLFOLLOW) = '" & Bln & "' AND YEAR(TGLFOLLOW) = " & Thn & ") OR (" & txtHelp & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) <= 21 And NoRangka in (" & addFilter(Bln, Thn) & "))) "
'        strSQL = "SELECT COUNT(DISTINCT NOFAK) FROM BOOK A WHERE (NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDER WHERE ((MONTH(TGLPROSES) = '" & Bln & "' AND YEAR(TGLPROSES) = " & Thn & ") OR (" & txtHelp1 & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) <= 30  And NoRangka in (" & addFilter(Bln, Thn) & ")) " & _
'                 "OR NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDERCALL WHERE ((MONTH(TGLFOLLOW) = '" & Bln & "' AND YEAR(TGLFOLLOW) = " & Thn & ") OR (" & txtHelp & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) <= 30 And NoRangka in (" & addFilter(Bln, Thn) & "))) "
    Else
        strSql = "SELECT COUNT(distinct nofak) FROM BOOK WITH(NOLOCK) WHERE MONTH(DATE_BOOK) = '" & Bln & "' AND YEAR(DATE_BOOK) = '" & Thn & "'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_10 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_11(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String

    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    'strSQL = "SELECT COUNT(DISTINCT A.NO_WO) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE DATANG IN ('CALL', 'SMS', 'MAIL') AND FS_LOG_DEL = '' AND MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND FS_RC <> '' AND FS_RC <> 'R001'"
    If a Then
        If Bln = 1 Then
            txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
        End If
        '''asli''''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -7 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -14 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        strSql = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C WITH(NOLOCK) ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -21 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -30 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
    Else
        If Bln = 1 Then
            txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        End If
        '''asli''''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -7 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -14 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        strSql = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C WITH(NOLOCK) ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -21 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -30 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_11 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_11b(Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    strSql = "SELECT COUNT(DISTINCT A.NO_WO) FROM JOBDTL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO WHERE FS_LOG_DEL = '' AND MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND FS_RC <> '' AND FS_RC <> 'R001'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_11b = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_12(Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim Ket As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
    If Bln = 1 Then
        Ket = "MONTH(entryCr1) = '12' AND YEAR(entryCr1) = '" & CDbl(Thn - 1) & "'"
    Else
        Ket = "MONTH(entryCr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr1) = '" & CDbl(Thn) & "'"
    End If
    strSql = "Select Count(distinct a.SPK_NO) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO INNER JOIN SPK C WITH(NOLOCK) ON A.SPK_NO = C.SPK_NO Where A.BATAL = 0 And SBI1000 = 1 and " & Ket & ""
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_12 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_12b(Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim Ket As String
Dim ket2 As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    If Bln = 1 Then
        Ket = "MONTH(entrycr1) = '12' AND YEAR(entrycr1) = '" & CDbl(Thn - 1) & "'"
        ket2 = "((MONTH(date) = '12' AND YEAR(date) = '" & CDbl(Thn - 1) & "') OR (MONTH(date) = '" & CDbl(Bln) & "' AND YEAR(date) = '" & CDbl(Thn) & "'))"
    Else
        Ket = "MONTH(entrycr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entrycr1) = '" & CDbl(Thn) & "'"
        ket2 = "((MONTH(DATE) = '" & CDbl(Bln - 1) & "' AND YEAR(DATE) = '" & CDbl(Thn) & "') OR (MONTH(DATE) = '" & CDbl(Bln) & "' AND YEAR(DATE) = '" & CDbl(Thn) & "'))"
    End If
'    strSQL = "SELECT COUNT(DISTINCT A.SPK_NO) FROM " & Log.LogSPK & "..BPK A INNER JOIN " & Log.LogSPK & "..SPK B ON A.SPK_NO = B.SPK_NO INNER JOIN " & Log.LogSPK & "..AFI D ON A.SPK_NO = D.SPK_NO " & _
'             "INNER JOIN " & Log.LogBengkel & "..JOBDTA C ON B.NORANGKA = C.NO_CHASSIS WHERE FS_RC = 'R001' AND A.BATAL = 0 AND C.DELIV <> 'BTL' AND FS_LOG_DEL = '' AND " & ket & " AND " & ket2
    strSql = "SELECT COUNT(DISTINCT A.SPK_NO) FROM " & Log.LogSPK & "..BPK A WITH(NOLOCK) INNER JOIN " & Log.LogSPK & "..AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO INNER JOIN " & Log.LogSPK & "..SPK C WITH(NOLOCK) ON A.SPK_NO = C.SPK_NO " & _
             "Where A.BATAL = 0 And SBI1000 = 1 and " & Ket & " AND NORANGKA IN (SELECT NO_CHASSIS FROM " & Log.LogBengkel & "..JOBDTA WITH(NOLOCK) WHERE FS_RC = 'R001' AND DELIV <> 'BTL' AND " & ket2 & ")"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_12b = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

'===============================================================================

Public Function GraphUnd(a As String, tg1 As Date, tg2 As Date, aMax As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
        strSql = "select CAST( a.SbeNExt AS VARCHAR(10)) as SBE, count(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln " & _
                  "WHEN JT_Tgl < JT_Bln then JT_Tgl " & _
                  "WHEN JT_Tgl = JT_Bln then  Jt_Bln END) as DateNextSv " & _
                  "FROM M_Estimasi a WITH(NOLOCK) INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
                  "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis WHERE (a.JT_Tgl is Not Null " & _
                  "And  a.SbeNExt < " & aMax
                  If a = "0000" Then
                     strSql = strSql & " And Right(CAST( a.SbeNExt AS VARCHAR(10)),4) = '0000'"
                  End If
                  strSql = strSql & ") And (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln " & _
                  "WHEN JT_Tgl < JT_Bln then JT_Tgl " & _
                  "WHEN JT_Tgl = JT_Bln then  Jt_Bln END) between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' " & _
                  "Group By a.SbeNext Order By a.SbeNExt Asc"
          
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GraphUnd = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function GraphUndSPK(tg1 As Date, tg2 As Date) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
        strSql = "select '1000' as Seribu, Count(*) as Jumlah FROM T_MRS WITH(NOLOCK) " & _
                 "Where Tgl1 between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' "
          
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GraphUndSPK = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function GraphDat(a As String, tg1 As Date, tg2 As Date, aMax As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      
    strSql = "Select b.FS_Nilai, Count(a.FS_RC) FROM T_Regular b WITH(NOLOCK) " & _
             "Inner Join JobDTA a WITH(NOLOCK) On b.FS_RC = a.FS_RC " & _
             "Where a.[date] between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' " & _
             "And  b.FS_Nilai < " & aMax
                If a = "0000" Then
                    strSql = strSql & " And (Right(CAST( b.FS_Nilai AS VARCHAR(10)),4) = '0000' Or Right(CAST( b.FS_Nilai AS VARCHAR(10)),4) = '1000')"
                End If
             strSql = strSql & " Group By b.FS_Nilai Order By b.Fs_Nilai"
    If Log.RPLHistori = 1 Then
        strSql = UCase(strSql)
        strSql = Replace(strSql, "T_REGULAR", "[" & dbReplikasiHO & "]..T_REGULAR")
        
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GraphDat = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function GraphRatio1000U(tg1 As Date, tg2 As Date) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
        strSql = "Select * FROM T_MRS WITH(NOLOCK) Where Tgl1 between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' "
          
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GraphRatio1000U = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function GraphRatio1000D(tg1 As Date, tg2 As Date) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
        strSql = "Select a.NoRangka, b.FS_RC " & _
                 "FROM " & Log.LogSPK & "..T_MRS a WITH(NOLOCK) Inner Join " & Log.LogBengkel & "..JobDTA b WITH(NOLOCK) On a.NoRangka = b.No_Chassis " & _
                 "Where b.FS_RC  = 'R001' And a.Tgl1 between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' "
          
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GraphRatio1000D = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function GraphRat(a As String, tg1 As Date, tg2 As Date, aMax As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
'        strSQL = "Select b.KM, Count(b.norangka) " & _
'                 "FROM M_Estimasi_Temp a Inner Join JobDta_Temp b " & _
'                 "On a.KM = b.KM And a.NoRangka = b.NoRangka " & _
'                 "Inner Join M_Estimasi c On a.KM = c.SBENext And a.NoRangka = c.NoRangka " & _
'                 "Where (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln " & _
'                 "WHEN JT_Tgl < JT_Bln then JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) between '" & TglNul(tg1) & "' And '" & TglNul(tg2) & "' " & _
'                 "And CAST(b.KM as NUMERIC) < " & aMax
'                 If a = "0000" Then
'                     strSQL = strSQL & " And Right(CAST( a.KM AS VARCHAR(10)),4) = '0000'"
'                 End If
'                 strSQL = strSQL & " Group By b.Km"

'    strSQL = "Select B.BSBE, Count(B.BRangka)  FROM (Select NoRangka As ARangka, SbeNext as ASBE " & _
'             "FROM M_ESTIMASI Where JT_Tgl is Not Null) A INNER JOIN " & _
'             "(Select a.No_Chassis as BRangka, b.FS_Nilai as BSBE FROM T_Regular b " & _
'             "Inner Join JobDTA a On b.FS_RC = a.FS_RC) B " & _
'             "On A.ASBE = B.BSBE And A.ARangka = B.BRangka " & _
'             "Inner Join M_Estimasi C On A.ASBE = C.SBENext And A.ARangka = C.NoRangka " & _
'             "Where (CASE WHEN C.JT_Tgl > C.JT_Bln then  C.JT_Bln " & _
'             "WHEN C.JT_Tgl < C.JT_Bln then C.JT_Tgl WHEN C.JT_Tgl = C.JT_Bln then  C.Jt_Bln END) " & _
'             "between '" & TglNul(tg1) & "' And '" & TglNul(tg2) & "' " & _
'             "And CAST(B.BSBE as NUMERIC) < " & aMax
'             If a = "0000" Then
'                strSQL = strSQL & " And Right(CAST( A.ASbe AS VARCHAR(10)),4) = '0000'"
'             End If
'             strSQL = strSQL & " Group By b.BSbe"
     
     strSql = "Select X.XSBE as SBE, Count(X.XSBE) as nSBE FROM( " & _
              "select a.NoRangka as XRANGKA, CAST( a.SbeNExt AS VARCHAR(10)) as XSBE " & _
              "FROM M_Estimasi a WITH(NOLOCK) " & _
              "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
              "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis " & _
              "WHERE (a.JT_Tgl is Not Null And  a.SbeNExt < " & aMax & " "
                If a = "0000" Then
                    strSql = strSql & " And Right(CAST( a.SbeNExt AS VARCHAR(10)),4) = '0000') "
                End If
              strSql = strSql & " And (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) " & _
              "Between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "') X " & _
              "Inner Join " & _
              "(Select B.BRangka AS YRANGKA, B.BSBE AS YSBE FROM (Select NoRangka As ARangka, SbeNext as ASBE " & _
              "FROM M_ESTIMASI WITH(NOLOCK) Where JT_Tgl is Not Null) A Inner Join " & _
              "(Select a.No_Chassis as BRangka, b.FS_Nilai as BSBE " & _
              "FROM T_Regular b WITH(NOLOCK) Inner Join JobDTA a WITH(NOLOCK) On b.FS_RC = a.FS_RC) B " & _
              "On A.ASBE = B.BSBE And A.ARangka = B.BRangka " & _
              "Inner Join M_Estimasi C WITH(NOLOCK) On A.ASBE = C.SBENext And A.ARangka = C.NoRangka " & _
              "Where (CASE WHEN C.JT_Tgl > C.JT_Bln then  C.JT_Bln " & _
              "WHEN C.JT_Tgl < C.JT_Bln then C.JT_Tgl " & _
              "WHEN C.JT_Tgl = C.JT_Bln then  C.Jt_Bln END) between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' " & _
              "And CAST(B.BSBE as NUMERIC) < " & aMax & " "
                If a = "0000" Then
                    strSql = strSql & " And Right(CAST( A.ASbe AS VARCHAR(10)),4) = '0000' "
                End If
              strSql = strSql & " ) Y On X.XRANGKA = Y.YRANGKA And X.XSBE = Y.YSBE " & _
              "Group By X.XSBE Order By CAST(ROUND(x.XSBE, 0) AS INT)"
    If Log.RPLHistori = 1 Then
        strSql = UCase(strSql)
        strSql = Replace(strSql, "T_REGULAR", "[" & dbReplikasiHO & "]..T_REGULAR")
    End If
              
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GraphRat = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function
