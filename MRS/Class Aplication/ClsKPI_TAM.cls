VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsKPI_TAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSQL As String

''===========================================================================================''
''ooooooooooo000ooooooooooo KPI TAM SETELAH MEETING KGW 13 OKT 2010 ooooooooooo000ooooooooooo''
''===========================================================================================''

Public Function KPI_1(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim ket As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    If Bln = 1 Then
        ket = "MONTH(entryCr1) = '12' AND YEAR(entryCr1) = '" & CDbl(Thn - 1) & "'"
    Else
        ket = "MONTH(entryCr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr1) = '" & CDbl(Thn) & "'"
    End If
    If a = True Then
'        'strSQL = "Select Count(TglCR1) FROM AFI Where Month(TglCR1)= '" & Bln & "' And Year(TglCR1)='" & Thn & "'"
'        ''strSQL = "Select Count(*) FROM BPK Where Month(Tgl_BPK)= '" & Bln & "' And Year(Tgl_Bpk)='" & Thn & "' AND BATAL = 0"
'        strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where Month(Tgl_BPK)= '" & Bln & "' And Year(Tgl_Bpk)='" & Thn & "' AND A.BATAL = 0"
        strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where A.BATAL = 0 and " & ket & ""
    Else
'        'strSQL = "Select Count(TglCR1) FROM AFI Where Month(TglCR1)= '" & Bln & "' And Year(TglCR1)='" & Thn & "' And OkCR1 = 1"
'        strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where Month(Tgl_BPK)= '" & Bln & "' And Year(Tgl_BPK)='" & Thn & "' And OkCR1 = 1 and Month(entryCR1)= '" & Bln & "' And Year(entryCR1)='" & Thn & "' AND A.BATAL = 0"
        strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where OkCR1 = 1 and A.BATAL = 0 and " & ket & ""
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
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
        'strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where Month(entryCR1)= '" & Bln & "' And Year(entryCR1)='" & Thn & "' And OkCR1 = 1 AND A.BATAL = 0"
        strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where Month(entryCR1)= '" & Bln & "' And Year(entryCR1)='" & Thn & "' and (Tgl_BPK > '" & Format(TglYYYYMMDD(GetTglAkhir(Bln, Thn, 0)), "MM/DD/YYYY") & "' OR Tgl_BPK < '" & Format(Bln, "00") & "/" & "01/" & Format(Thn, "0000") & "') And OkCR1 = 1 AND A.BATAL = 0"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_1c = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_2(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim ket As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    If Bln = 1 Then
        ket = "(MONTH(entryCr2) = '12' AND YEAR(entryCr2) = '" & CDbl(Thn - 1) & "')) and MONTH(entryCr1) = '12' AND YEAR(entryCr1) = '" & CDbl(Thn - 1) & "'"
    Else
        ket = "(MONTH(entryCr2) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr2) = '" & CDbl(Thn) & "')) and MONTH(entryCr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr1) = '" & CDbl(Thn) & "'"
    End If
    If a = True Then
'        'strSQL = "Select Count(NoSPK) FROM T_Transfer Where Month(Tgltransfer)= '" & Bln & "' And Year(TglTransfer)='" & Thn & "'"
'        strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where Month(Tgl_BPK)= '" & Bln & "' And Year(Tgl_BPK)='" & Thn & "' And OkCR1 = 1 AND A.BATAL = 0 AND ISNULL(TGLCR2,'') <> '' and Month(entryCR2)= '" & Bln & "' And Year(entryCR2)='" & Thn & "'"
        strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where OkCR1 = 1 AND A.BATAL = 0 AND ISNULL(TGLCR2,'') <> '' and ((Month(entryCR2)= '" & Bln & "' And Year(entryCR2)='" & Thn & "') or " & ket
    Else
        strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where OkCR1 = 1 AND A.BATAL = 0 AND ISNULL(TGLCR2,'') <> '' and Month(entryCR2)= '" & Bln & "' AND (Tgl_BPK > '" & Format(TglYYYYMMDD(GetTglAkhir(Bln, Thn, 0)), "MM/DD/YYYY") & "' OR Tgl_BPK < '" & Format(Bln, "00") & "/" & "01/" & Format(Thn, "0000") & "') And Year(entryCR2)='" & Thn & "'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_2 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_3(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If a = True Then
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        'strSQL = "Select Count(BPK_No) FROM BPK Where Month(Tgl_BPK)= '" & Bln & "' And Year(Tgl_BPK)='" & Thn & "'"
        strSQL = "Select Count(*) FROM T_MRS A INNER JOIN AFI B ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND SBI1000 = '1'"
    Else
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        strSQL = "Select Count(DISTINCT LastNoWo) FROM M_ListReminder Where Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "' And Tipe = '1K'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If a = True Then
        rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
    Else
        rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    End If
    Set rs.ActiveConnection = Nothing
    Set KPI_3 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_3sales(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If a = True Then
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        strSQL = "Select Count(*) FROM T_MRS A INNER JOIN AFI B ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' "
    Else
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        strSQL = "Select Count(*) FROM T_MRS A INNER JOIN AFI B ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND ISNULL(REMINDERSALES,'') = '1'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If a = True Then
        rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
    Else
        rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
    End If
    Set rs.ActiveConnection = Nothing
    Set KPI_3sales = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_4(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If a = True Then
        strSQL = "Select Count(DISTINCT Lastnowo) From M_ListRemindercall Where K_DM = 1 And Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "'"
    Else
        strSQL = "Select Count(DISTINCT Lastnowo) From M_ListReminder Where Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "' And Tipe = 'DM'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_4 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_5(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If a = True Then
        strSQL = "Select Count(DISTINCT Lastnowo) From M_ListRemindercall Where K_Telp = 1 And Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "'"
    Else
        strSQL = "Select Count(DISTINCT Lastnowo) From M_ListRemindercall Where Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_5 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_6(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If a = True Then
        strSQL = "select Count(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then " & _
                 "JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) as DateNextSv " & _
                 "FROM M_Estimasi a INNER JOIN History b ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
                 "INNER JOIN Custdat c On a.NoRangka = c.No_Chassis WHERE a.JT_Tgl is Not Null AND " & _
                 "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
                 "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "' "
    Else
        'strSQL = "Select Count(DISTINCT a.LastNowo) From M_ListReminder a Left Join M_ListReminderCall b on a.LastNowo = b.LastNowo Where Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "'"
        'strSQL = "Select Count(DISTINCT LastNowo) From M_ListReminder Where Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "'"
        strSQL = "Select Count(DISTINCT LastNowo) From M_ListReminder Where Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "' AND TIPE <> '1K'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_6 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_6a(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    If a = True Then
        strSQL = "Select Count(NoSPK) FROM T_MRS Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "'"
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_6a = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_6b(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If a = True Then
        strSQL = "Select Count(DISTINCT LastNowo) From M_ListReminderCall Where Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "' AND KMNEXTSVC <> '1000' AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder WHERE Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "')"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_6b = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function KPI_8(Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    'strSQL = "Select Count(DISTINCT Lastnowo) From M_ListRemindercall Where App = 1 And Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "'"
    strSQL = "SELECT COUNT(DISTINCT LASTNOWO) FROM BOOK A INNER JOIN M_LISTREMINDERCALL B ON A.NO_CHASSIS = B.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,TGLFOLLOW, DATE) >= 0 AND DATEDIFF(d,TGLFOLLOW, DATE) <=7 AND MONTH(TGLFOLLOW) = '" & Bln & "' AND YEAR(TGLFOLLOW) = '" & Thn & "' AND DELIV <> 'BTL'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_8 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function KPI_9(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    If a Then
    '    'strSQL = "SELECT COUNT(*) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO WHERE DATANG IN ('CALL','MAIL','SMS') AND MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = ''"
    '    strSQL = "SELECT COUNT(*) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -30 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = ''"
        strSQL = "SELECT COUNT(*) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -30 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder WHERE Month(TglProses)= '10' And Year(TglProses)='2010')"
    Else
        strSQL = "SELECT COUNT(*) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -30 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_9 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_10(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    If a Then
        'strSQL = "Select Count(Lastnowo) From M_ListRemindercall Where App = 1 And Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "'"
        strSQL = "SELECT COUNT(DISTINCT NOFAK) FROM BOOK A WHERE MONTH(A.DATE) = '" & Bln & "' AND YEAR(A.DATE) = '" & Thn & "' AND DELIV <> 'BTL' AND (NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDER WHERE DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE) <=7) " & _
                 "OR NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDERCALL WHERE DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE) <=7)) "
    Else
        strSQL = "SELECT COUNT(*) FROM BOOK WHERE MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND DELIV <> 'BTL'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_10 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_11(Bln As Integer, Thn As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    strSQL = "SELECT COUNT(DISTINCT A.NO_WO) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE DATANG IN ('CALL', 'SMS', 'MAIL') AND FS_LOG_DEL = '' AND MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND FS_RC <> '' AND FS_RC <> 'R001'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
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

    strSQL = "SELECT COUNT(DISTINCT A.NO_WO) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO WHERE FS_LOG_DEL = '' AND MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND FS_RC <> '' AND FS_RC <> 'R001'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_11b = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_12(Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim ket As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    If Bln = 1 Then
        ket = "MONTH(entrycr1) = '12' AND YEAR(entrycr1) = '" & CDbl(Thn - 1) & "'"
    Else
        ket = "MONTH(entrycr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entrycr1) = '" & CDbl(Thn) & "'"
    End If
    'strSQL = "SELECT COUNT(*) FROM BPK WHERE " & ket & " AND BATAL = 0"
    strSQL = "Select Count(*) FROM BPK A INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO Where A.BATAL = 0 and " & ket & ""
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_12 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_12b(Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim ket As String
Dim ket2 As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    If Bln = 1 Then
        ket = "MONTH(entrycr1) = '12' AND YEAR(entrycr1) = '" & CDbl(Thn - 1) & "'"
        ket2 = "((MONTH(date) = '12' AND YEAR(date) = '" & CDbl(Thn - 1) & "') OR (MONTH(date) = '" & CDbl(Bln) & "' AND YEAR(date) = '" & CDbl(Thn) & "'))"
    Else
        ket = "MONTH(entrycr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entrycr1) = '" & CDbl(Thn) & "'"
        ket2 = "((MONTH(DATE) = '" & CDbl(Bln - 1) & "' AND YEAR(DATE) = '" & CDbl(Thn) & "') OR (MONTH(DATE) = '" & CDbl(Bln) & "' AND YEAR(DATE) = '" & CDbl(Thn) & "'))"
    End If
    strSQL = "SELECT COUNT(DISTINCT A.SPK_NO) FROM " & Log.LogSPK & "..BPK A INNER JOIN " & Log.LogSPK & "..SPK B ON A.SPK_NO = B.SPK_NO INNER JOIN " & Log.LogSPK & "..AFI D ON A.SPK_NO = D.SPK_NO " & _
             "INNER JOIN " & Log.LogBengkel & "..JOBDTA C ON B.NORANGKA = C.NO_CHASSIS WHERE FS_RC = 'R001' AND A.BATAL = 0 AND FS_LOG_DEL = '' AND " & ket & " AND " & ket2
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
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
    
        strSQL = "select CAST( a.SbeNExt AS VARCHAR(10)) as SBE, count(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln " & _
                  "WHEN JT_Tgl < JT_Bln then JT_Tgl " & _
                  "WHEN JT_Tgl = JT_Bln then  Jt_Bln END) as DateNextSv " & _
                  "FROM M_Estimasi a INNER JOIN History b ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
                  "INNER JOIN Custdat c On a.NoRangka = c.No_Chassis WHERE (a.JT_Tgl is Not Null " & _
                  "And  a.SbeNExt < " & aMax
                  If a = "0000" Then
                     strSQL = strSQL & " And Right(CAST( a.SbeNExt AS VARCHAR(10)),4) = '0000'"
                  End If
                  strSQL = strSQL & ") And (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln " & _
                  "WHEN JT_Tgl < JT_Bln then JT_Tgl " & _
                  "WHEN JT_Tgl = JT_Bln then  Jt_Bln END) between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' " & _
                  "Group By a.SbeNext Order By a.SbeNExt Asc"
          
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
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
    
        strSQL = "select '1000' as Seribu, Count(*) as Jumlah FROM T_MRS " & _
                 "Where Tgl1 between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' "
          
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
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
      
    strSQL = "Select b.FS_Nilai, Count(a.FS_RC) FROM T_Regular b " & _
             "Inner Join JobDTA a On b.FS_RC = a.FS_RC " & _
             "Where a.[date] between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' " & _
             "And  b.FS_Nilai < " & aMax
                If a = "0000" Then
                    strSQL = strSQL & " And (Right(CAST( b.FS_Nilai AS VARCHAR(10)),4) = '0000' Or Right(CAST( b.FS_Nilai AS VARCHAR(10)),4) = '1000')"
                End If
             strSQL = strSQL & " Group By b.FS_Nilai Order By b.Fs_Nilai"
                 
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
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
    
        strSQL = "Select * FROM T_MRS Where Tgl1 between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' "
          
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
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
    
        strSQL = "Select a.NoRangka, b.FS_RC " & _
                 "FROM " & Log.LogSPK & "..T_MRS a Inner Join " & Log.LogBengkel & "..JobDTA b On a.NoRangka = b.No_Chassis " & _
                 "Where b.FS_RC  = 'R001' And a.Tgl1 between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' "
          
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLConSPK, adOpenStatic, adLockReadOnly, adCmdText
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
     
     strSQL = "Select X.XSBE as SBE, Count(X.XSBE) as nSBE FROM( " & _
              "select a.NoRangka as XRANGKA, CAST( a.SbeNExt AS VARCHAR(10)) as XSBE " & _
              "FROM M_Estimasi a " & _
              "INNER JOIN History b ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
              "INNER JOIN Custdat c On a.NoRangka = c.No_Chassis " & _
              "WHERE (a.JT_Tgl is Not Null And  a.SbeNExt < " & aMax & " "
                If a = "0000" Then
                    strSQL = strSQL & " And Right(CAST( a.SbeNExt AS VARCHAR(10)),4) = '0000') "
                End If
              strSQL = strSQL & " And (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) " & _
              "Between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "') X " & _
              "Inner Join " & _
              "(Select B.BRangka AS YRANGKA, B.BSBE AS YSBE FROM (Select NoRangka As ARangka, SbeNext as ASBE " & _
              "FROM M_ESTIMASI Where JT_Tgl is Not Null) A Inner Join " & _
              "(Select a.No_Chassis as BRangka, b.FS_Nilai as BSBE " & _
              "FROM T_Regular b Inner Join JobDTA a On b.FS_RC = a.FS_RC) B " & _
              "On A.ASBE = B.BSBE And A.ARangka = B.BRangka " & _
              "Inner Join M_Estimasi C On A.ASBE = C.SBENext And A.ARangka = C.NoRangka " & _
              "Where (CASE WHEN C.JT_Tgl > C.JT_Bln then  C.JT_Bln " & _
              "WHEN C.JT_Tgl < C.JT_Bln then C.JT_Tgl " & _
              "WHEN C.JT_Tgl = C.JT_Bln then  C.Jt_Bln END) between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' " & _
              "And CAST(B.BSBE as NUMERIC) < " & aMax & " "
                If a = "0000" Then
                    strSQL = strSQL & " And Right(CAST( A.ASbe AS VARCHAR(10)),4) = '0000' "
                End If
              strSQL = strSQL & " ) Y On X.XRANGKA = Y.YRANGKA And X.XSBE = Y.YSBE " & _
              "Group By X.XSBE Order By CAST(ROUND(x.XSBE, 0) AS INT)"
              
              
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GraphRat = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


