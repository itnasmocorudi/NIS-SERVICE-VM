VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsKPICEK"
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
Dim sSQLReminder As String

    txtFilter = ""
   ' txtFilter = "select DISTINCT A.NORANGKA FROM M_Estimasi a " & _
                "INNER JOIN History b ON a.NoWo = b.No_Wo AND A.NOINV = B.NOINV AND B.FS_LOG_DEL = '' " & _
                "INNER JOIN Custdat c On a.NoRangka = c.No_Chassis " & _
                "INNER JOIN T_REGULAR D ON D.FS_NILAI = A.SBENEXT " & _
                "LEFT JOIN JOBDTA E ON A.NORANGKA = E.NO_CHASSIS AND D.FS_RC = E.FS_RC " & _
                "LEFT JOIN M_LISTREMINDER F ON F.LASTNOWO = A.NOWO AND F.NORANGKA = A.NORANGKA " & _
                "LEFT JOIN M_LISTREMINDERCALL G ON G.LASTNOWO = A.NOWO AND G.NORANGKA = A.NORANGKA " & _
                "WHERE a.JT_Tgl is Not Null AND ISNULL(TGLSVCNOW,'') <> '' AND " & _
                "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
                "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "' " & _
                "AND (CASE WHEN ISNULL(E.DATE,0) <> 0 AND ISNULL(F.TGLPROSES,0) = 0 AND ISNULL(G.TGLFOLLOW,0) = 0 THEN 0 ELSE 1 END) > 0 "
                
'     txtFilter = "select DISTINCT A.NORANGKA FROM M_Estimasi a " & _
'                "INNER JOIN T_REGULAR D ON D.FS_NILAI = A.SBENEXT " & _
'                "LEFT JOIN JOBDTA E ON A.NORANGKA = E.NO_CHASSIS AND D.FS_RC = E.FS_RC " & _
'                "LEFT JOIN M_LISTREMINDER F ON F.LASTNOWO = A.NOWO AND F.NORANGKA = A.NORANGKA " & _
'                "LEFT JOIN M_LISTREMINDERCALL G ON G.LASTNOWO = A.NOWO AND G.NORANGKA = A.NORANGKA " & _
'                "WHERE a.JT_Tgl is Not Null AND ISNULL(TGLSVCNOW,'') <> '' AND " & _
'                "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
'                "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "' " & _
'                "AND (CASE WHEN ISNULL(E.DATE,0) <> 0 AND ISNULL(F.TGLPROSES,0) = 0 AND ISNULL(G.TGLFOLLOW,0) = 0 THEN 0 ELSE 1 END) > 0 "
                    
    ''2017-10-02 perubahan filter data
    
'    txtFilter = "select DISTINCT A.NORANGKA FROM M_Estimasi a " & _
'                "INNER JOIN T_REGULAR D ON D.FS_NILAI = A.SBENEXT " & _
'                "LEFT JOIN JOBDTA E ON A.NORANGKA = E.NO_CHASSIS AND D.FS_RC = E.FS_RC " & _
'                "LEFT JOIN M_LISTREMINDER F ON F.LASTNOWO = A.NOWO AND F.NORANGKA = A.NORANGKA " & _
'                "LEFT JOIN M_LISTREMINDERCALL G ON G.LASTNOWO = A.NOWO AND G.NORANGKA = A.NORANGKA " & _
'                "WHERE a.JT_Tgl is Not Null AND ISNULL(TGLSVCNOW,'') <> '' AND " & _
'                "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
'                "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "' " & _
'      "AND ISNULL(E.DATE,'01/01/1900') = '01/01/1900' AND ISNULL(F.TGLPROSES,'01/01/1900') <> '01/01/1900' AND ISNULL(G.TGLFOLLOW,'01/01/1900') <> '01/01/1900' "
                
    ''20180219
    
    If Log.RPLHistori = 0 Then
        txtFilter = "select DISTINCT A.NORANGKA FROM M_Estimasi a WITH (NOLOCK) " & _
                "INNER JOIN T_REGULAR D WITH (NOLOCK) ON D.FS_NILAI = A.SBENEXT " & _
                "LEFT JOIN JOBDTA E WITH (NOLOCK) ON A.NORANGKA = E.NO_CHASSIS AND D.FS_RC = E.FS_RC " & _
                "LEFT JOIN M_LISTREMINDER F WITH (NOLOCK) ON F.LASTNOWO = A.NOWO AND F.NORANGKA = A.NORANGKA " & _
                "LEFT JOIN M_LISTREMINDERCALL G WITH (NOLOCK) ON G.LASTNOWO = A.NOWO AND G.NORANGKA = A.NORANGKA " & _
                "WHERE a.JT_Tgl is Not Null AND ISNULL(TGLSVCNOW,'') <> '' AND " & _
                "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
                "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "' "
        If Log.hansaRenkei = 1 Then
            sSQLReminder = "select reminderMRA, * from settingHansaRenkei"
            txtFilter = txtFilter & " and (ISNULL(e.date,'') = '' or e.DATE >= DATEADD(d," & objO.FetchData_Integer(sSQLReminder) & ",(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END)))"
        Else
            txtFilter = txtFilter & " and (ISNULL(e.date,'') = '' or e.DATE >= DATEADD(d,-14,(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END)))"
        End If
                ''"AND (CASE WHEN ISNULL(E.DATE,'01/01/1900') <> '01/01/1900' AND ISNULL(F.TGLPROSES,'01/01/1900') = '01/01/1900' AND ISNULL(G.TGLFOLLOW,'01/01/1900') = '01/01/1900' THEN 0 ELSE 1 END) > 0 "
                
    Else
        txtFilter = "select DISTINCT A.NORANGKA FROM M_Estimasi a WITH (NOLOCK) " & _
                "INNER JOIN [" & dbReplikasiHO & "]..T_REGULAR D WITH (NOLOCK) ON D.FS_NILAI = A.SBENEXT " & _
                "LEFT JOIN JOBDTA E WITH (NOLOCK) ON A.NORANGKA = E.NO_CHASSIS AND D.FS_RC = E.FS_RC " & _
                "LEFT JOIN M_LISTREMINDER F WITH (NOLOCK) ON F.LASTNOWO = A.NOWO AND F.NORANGKA = A.NORANGKA " & _
                "LEFT JOIN M_LISTREMINDERCALL G WITH (NOLOCK) ON G.LASTNOWO = A.NOWO AND G.NORANGKA = A.NORANGKA " & _
                "WHERE a.JT_Tgl is Not Null AND ISNULL(TGLSVCNOW,'') <> '' AND " & _
                "Month(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Bln & "' " & _
                "And Year(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & Thn & "' "
     
        If Log.hansaRenkei = 1 Then
            sSQLReminder = "select reminderMRA, * from settingHansaRenkei"
            txtFilter = txtFilter & " and (ISNULL(e.date,'') = '' or e.DATE >= DATEADD(d," & objO.FetchData_Integer(sSQLReminder) & ",(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END)))"
        Else
            txtFilter = txtFilter & " and (ISNULL(e.date,'') = '' or e.DATE >= DATEADD(d,-14,(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END)))"
        End If
                ''"AND (CASE WHEN ISNULL(E.DATE,'01/01/1900') <> '01/01/1900' AND ISNULL(F.TGLPROSES,'01/01/1900') = '01/01/1900' AND ISNULL(G.TGLFOLLOW,'01/01/1900') = '01/01/1900' THEN 0 ELSE 1 END) > 0 "
    End If
                    
    txtFilter = txtFilter + "UNION " + addFilter2(Bln, Thn)
    
    If OnKPI10 = True Then
        Dim nData As String
        Dim rsD As ADODB.Recordset
        Set rsD = objO.FetchData(txtFilter)
        If rsD.RecordCount > 0 Then
            rsD.MoveFirst
            Do Until rsD.EOF
                addFilter = addFilter & Trim(rsD.Fields(0).value) & "','"
                rsD.MoveNext
            Loop
            addFilter = "'" & Left(addFilter, Len(addFilter) - 2)
        Else
            addFilter = ""
        End If
    Else
        addFilter = txtFilter
    End If
End Function

Public Function addFilter2(Bln As Integer, Thn As Integer) As String
      Dim txtFilter2 As String
      Dim sSQLReminder As String

      '    txtFilter2 = "Select DISTINCT A.NORANGKA FROM " & Log.LogSPK & "..T_MRS a WITH(NOLOCK) " & _
      '                "LEFT JOIN " & Log.LogSPK & "..AFI B  WITH(NOLOCK) ON A.NOSPK = B.SPK_NO " & _
      '                "LEFT JOIN JOBDTA C WITH(NOLOCK) ON A.NORANGKA = C.NO_CHASSIS AND C.FS_RC = 'R001' " & _
      '                "LEFT JOIN M_LISTREMINDER D WITH(NOLOCK) ON D.LASTNOWO = A.NOSPK AND D.NORANGKA = A.NORANGKA " & _
      '                "LEFT JOIN M_LISTREMINDERCALL E WITH(NOLOCK) ON E.LASTNOWO = A.NOSPK AND E.NORANGKA = A.NORANGKA " & _
      '                "Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)= '" & Thn & "' AND SBI1000 = 1 " & _
      '                "AND (CASE WHEN ISNULL(C.DATE,0) <> 0 AND ISNULL(D.TGLPROSES,0) = 0 AND ISNULL(E.TGLFOLLOW,0) = 0 THEN 0 ELSE 1 END) > 0 "
          ''2017-10-02 || perubahan filter data
          
      '    txtFilter2 = "Select DISTINCT A.NORANGKA FROM " & Log.LogSPK & "..T_MRS a WITH(NOLOCK) " & _
      '                "LEFT JOIN " & Log.LogSPK & "..AFI B  WITH(NOLOCK) ON A.NOSPK = B.SPK_NO " & _
      '                "LEFT JOIN JOBDTA C WITH(NOLOCK) ON A.NORANGKA = C.NO_CHASSIS AND C.FS_RC = 'R001' " & _
      '                "LEFT JOIN M_LISTREMINDER D WITH(NOLOCK) ON D.LASTNOWO = A.NOSPK AND D.NORANGKA = A.NORANGKA " & _
      '                "LEFT JOIN M_LISTREMINDERCALL E WITH(NOLOCK) ON E.LASTNOWO = A.NOSPK AND E.NORANGKA = A.NORANGKA " & _
      '                "Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)= '" & Thn & "' AND SBI1000 = 1 " & _
      '                "AND ISNULL(C.DATE,'01/01/1900') = '01/01/1900' AND ISNULL(D.TGLPROSES,'01/01/1900') <> '01/01/1900' AND ISNULL(E.TGLFOLLOW, '01/01/1900') <> '01/01/1900' "
          
1         txtFilter2 = "Select DISTINCT A.NORANGKA FROM " & Log.LogSPK & "..T_MRS a WITH(NOLOCK) " & _
                      "LEFT JOIN " & Log.LogSPK & "..AFI B  WITH(NOLOCK) ON A.NOSPK = B.SPK_NO " & _
                      "LEFT JOIN JOBDTA C WITH(NOLOCK) ON A.NORANGKA = C.NO_CHASSIS AND C.FS_RC = 'R001' " & _
                      "LEFT JOIN M_LISTREMINDER D WITH(NOLOCK) ON D.LASTNOWO = A.NOSPK AND D.NORANGKA = A.NORANGKA " & _
                      "LEFT JOIN M_LISTREMINDERCALL E WITH(NOLOCK) ON E.LASTNOWO = A.NOSPK AND E.NORANGKA = A.NORANGKA " & _
                      "Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)= '" & Thn & "' AND SBI1000 = 1 "
                      ''20210224 | req mb thania menutup query
                      ''"AND ISNULL(C.DATE,'01/01/1900') = '01/01/1900' AND ISNULL(D.TGLPROSES,'01/01/1900') <> '01/01/1900' AND ISNULL(E.TGLFOLLOW, '01/01/1900') <> '01/01/1900' "
         
2         If Log.hansaRenkei = 1 Then
3             sSQLReminder = "select reminderMRA, * from settingHansaRenkei"
4             txtFilter2 = txtFilter2 & " and (ISNULL(c.date,'') = '' or c.DATE >= DATEADD(d," & objO.FetchData_Integer(sSQLReminder) & ",a.tgl1))"
5         Else
6             txtFilter2 = txtFilter2 & " and (ISNULL(c.date,'') = '' or c.DATE >= DATEADD(d,-7,a.tgl1))"
7         End If
          
8         addFilter2 = txtFilter2
End Function

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

Public Function KPI_1new(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim Ket As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    If Bln = 1 Then
        Ket = "MONTH(C.tgltransfer) = '12' AND YEAR(C.tgltransfer) = '" & CDbl(Thn - 1) & "'"
    Else
        Ket = "MONTH(C.tgltransfer) = '" & CDbl(Bln - 1) & "' AND YEAR(C.tgltransfer) = '" & CDbl(Thn) & "'"
    End If
    If a = True Then
        strSql = "Select Count(A.SPK_NO) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO " & _
                 "Inner Join t_transferCR1 C WITH(NOLOCK) on A.SPK_NO = C.nospk Where A.BATAL = 0 and " & Ket & ""
    Else
        strSql = "Select Count(A.SPK_NO) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO " & _
                 "Inner Join t_transferCR1 C WITH(NOLOCK) on A.SPK_NO = C.nospk Where C.OkCR1 = 1 and A.BATAL = 0 and " & Ket & ""
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_1new = rs
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

Public Function KPI_2new(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim Ket As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    If Bln = 1 Then
        Ket = "((MONTH(C.tgltransfer) = '12' AND YEAR(C.tgltransfer) = '" & CDbl(Thn - 1) & "')) and MONTH(D.tgltransfer) = '12' AND YEAR(D.tgltransfer) = '" & CDbl(Thn - 1) & "')"
    Else
        Ket = "((MONTH(C.tgltransfer) = '" & CDbl(Bln - 1) & "' AND YEAR(C.tgltransfer) = '" & CDbl(Thn) & "')) and MONTH(D.tgltransfer) = '" & CDbl(Bln - 1) & "' AND YEAR(D.tgltransfer) = '" & CDbl(Thn) & "')"
    End If
    If a = True Then
        strSql = "Select Count(A.SPK_NO) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO " & _
                 "Inner Join T_Transfer C WITH(NOLOCK) On A.SPK_NO = C.nospk Inner Join T_TransferCR1 D WITH(NOLOCK) On A.SPK_NO = D.nospk " & _
                 "Where D.OkCR1 = 1 AND A.BATAL = 0 AND ISNULL(B.TGLCR2,'') <> '' and ((Month(C.tgltransfer)= '" & Bln & "' And Year(C.tgltransfer)='" & Thn & "') or " & Ket
    Else
        strSql = "Select Count(A.SPK_NO) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO " & _
                 "Inner Join T_Transfer C WITH(NOLOCK) On A.SPK_NO = C.nospk Inner Join T_TransferCR1 D WITH(NOLOCK) On A.SPK_NO = D.nospk " & _
                 "Where D.OkCR1 = 1 AND A.BATAL = 0 AND ISNULL(B.TGLCR2,'') <> '' and Month(C.tgltransfer)= '" & Bln & "' AND (Tgl_BPK > '" & Format(TglYYYYMMDD(GetTglAkhir(Bln, Thn, 0)), "MM/DD/YYYY") & "' OR Tgl_BPK < '" & Format(Bln, "00") & "/" & "01/" & Format(Thn, "0000") & "') And Year(C.tgltransfer)='" & Thn & "'"
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_2new = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_3(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String
Dim txtHelp2 As String

    On Error GoTo errHandler
    If Bln = 1 Then
        txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        txtHelp2 = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
    Else
        txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        txtHelp2 = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
    End If
    
    If a = True Then
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        strSql = "SELECT COUNT(DISTINCT NORANGKA) AS xx FROM (" & addFilter2(Bln, Thn) & ") AS X"
    Else
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        strSql = "Select Count(DISTINCT LastNoWo) AS xx FROM ( " & _
                 "Select DISTINCT LastNoWo FROM M_ListReminder WITH(NOLOCK) " & _
                 "Where ((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') or " & txtHelp & ") And Tipe = '1K' And " & _
                 "NoRangka in (" & addFilter2(Bln, Thn) & ") UNION " & _
                 "Select DISTINCT LastNoWo FROM M_ListReminderCall WITH(NOLOCK) " & _
                 "Where ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') or " & txtHelp2 & ") And KmNextSvc = '1000' And " & _
                 "NoRangka in (" & addFilter2(Bln, Thn) & ")) as Z"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If a = True Then
        rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
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
        strSql = "SELECT 0"
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
Dim ket2 As String

    On Error GoTo errHandler
    If Bln = 1 Then
        ket2 = "((MONTH(entryCr2) = '12' AND YEAR(entryCr2) = '" & CDbl(Thn - 1) & "')) and MONTH(entryCr1) = '12' AND YEAR(entryCr1) = '" & CDbl(Thn - 1) & "')"
        Ket = "MONTH(entryCr1) = '12' AND YEAR(entryCr1) = '" & CDbl(Thn - 1) & "'"
        txtHelp = "month(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '12' " & _
                  "AND YEAR(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '" & CDbl(Thn - 1) & "'"
    Else
        ket2 = "((MONTH(entryCr2) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr2) = '" & CDbl(Thn) & "')) and MONTH(entryCr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr1) = '" & CDbl(Thn) & "')"
        Ket = "MONTH(entryCr1) = '" & CDbl(Bln - 1) & "' AND YEAR(entryCr1) = '" & CDbl(Thn) & "'"
        txtHelp = "month(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '" & CDbl(Bln - 1) & "' " & _
                  "AND YEAR(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),TGLCR2)) = '" & CDbl(Thn) & "'"
    End If
    
    If a = True Then
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        strSql = "Select Count(*) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO Where OkCR1 = 1 AND A.BATAL = 0 AND REMINDERSALES = 1 AND ISNULL(TGLCR2,'') <> '' and ((Month(entryCR2)= '" & Bln & "' And Year(entryCR2)='" & Thn & "') or " & ket2
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

Public Function KPI_3salesnew(a As Boolean, Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim txtHelp As String
Dim Ket As String
Dim ket2 As String

    On Error GoTo errHandler
    If Bln = 1 Then
        ket2 = "((MONTH(D.TglTransfer) = '12' AND YEAR(D.TglTransfer) = '" & CDbl(Thn - 1) & "')) and MONTH(C.TglTransfer) = '12' AND YEAR(C.TglTransfer) = '" & CDbl(Thn - 1) & "')"
        Ket = "MONTH(C.TglTransfer) = '12' AND YEAR(C.TglTransfer) = '" & CDbl(Thn - 1) & "'"
        txtHelp = "month(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),B.TGLCR2)) = '12' " & _
                  "AND YEAR(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),B.TGLCR2)) = '" & CDbl(Thn - 1) & "'"
    Else
        ket2 = "((MONTH(D.TglTransfer) = '" & CDbl(Bln - 1) & "' AND YEAR(D.TglTransfer) = '" & CDbl(Thn) & "')) and MONTH(C.TglTransfer) = '" & CDbl(Bln - 1) & "' AND YEAR(C.TglTransfer) = '" & CDbl(Thn) & "')"
        Ket = "MONTH(C.TglTransfer) = '" & CDbl(Bln - 1) & "' AND YEAR(C.TglTransfer) = '" & CDbl(Thn) & "'"
        txtHelp = "month(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),B.TGLCR2)) = '" & CDbl(Bln - 1) & "' " & _
                  "AND YEAR(dateadd(d,(CASE WHEN (LEFT(KodeKategory,2) = 'WU' OR LEFT(KodeKategory,2) = 'BU') THEN 21 ELSE 7 END),B.TGLCR2)) = '" & CDbl(Thn) & "'"
    End If
    
    If a = True Then
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        strSql = "Select Count(A.SPK_NO) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO " & _
                 "Inner Join T_TransferCR1 C WITH(NOLOCK) on A.SPK_NO = C.Nospk Inner Join T_Transfer D WITH(NOLOCK) on A.SPK_NO = D.Nospk " & _
                 "Where C.OkCR1 = 1 AND A.BATAL = 0 AND D.REMINDERSALES = 1 AND ISNULL(B.TGLCR2,'') <> '' and ((Month(D.TglTransfer)= '" & Bln & "' And Year(D.TglTransfer)='" & Thn & "') or " & ket2
    Else
        If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
        strSql = "Select Count(A.No_SPK) FROM T_MRS A WITH(NOLOCK) INNER JOIN AFI B WITH(NOLOCK) ON A.NOSPK = B.SPK_NO Where Month(Tgl1)= '" & Bln & "' And Year(Tgl1)='" & Thn & "' AND ISNULL(D.REMINDERSALES,'0') = '1'"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If a = True Then
        rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Else
        rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    End If
    Set rs.ActiveConnection = Nothing
    Set KPI_3salesnew = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function KPI_4(a As Boolean, Bln As Integer, Thn As Integer, tipe As String) As ADODB.Recordset
Dim txtHelp As String
Dim txtHelp2 As String
Dim txtCriteria As String

    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    txtCriteria = ""
    txtHelp = ""
    txtHelp2 = ""
    If tipe = "E" Then
        ' DM SBE
        txtCriteria = " AND LASTNOWO not IN (SELECT SPK_NO FROM " & Log.LogSPK & "..SPK)"
    ElseIf tipe = "I" Then
        ' DM SBI
        txtCriteria = " AND LASTNOWO IN (SELECT SPK_NO FROM " & Log.LogSPK & "..SPK)"
    End If
    
    If a = True Then
        If Bln = 1 Then
            txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
            txtHelp2 = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
            txtHelp2 = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        End If
        strSql = "Select Count(DISTINCT Lastnowo) AS xx From M_ListRemindercall WITH(NOLOCK) Where K_DM = 1 And ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') Or (" & txtHelp & ")) and NoRangka in (" & addFilter(Bln, Thn) & ") and lastnowo in (Select DISTINCT Lastnowo From M_ListReminder WITH(NOLOCK) Where ((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') Or (" & txtHelp2 & ")) And Tipe = 'DM')" & txtCriteria
    Else
        If Bln = 1 Then
            txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        End If
        strSql = "Select Count(DISTINCT Lastnowo) AS xx From M_ListReminder WITH(NOLOCK) Where ((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') Or (" & txtHelp & ")) And Tipe = 'DM' AND NoRangka IN (" & addFilter(Bln, Thn) & ")" & txtCriteria
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

Public Function KPI_5(a As Boolean, Bln As Integer, Thn As Integer, tipe As String) As ADODB.Recordset
Dim txtHelp As String
Dim txtCriteria As String

    On Error GoTo errHandler
    txtCriteria = ""
    If Trim(tipe) = "E" Then
        ' TELP SBE
        txtCriteria = " and kmnextsvc <> 1000"
    ElseIf Trim(tipe) = "I" Then
        ' TELP SBI
        txtCriteria = " and kmnextsvc = 1000"
    End If
    If Bln = 1 Then
        txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
    Else
        txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
    End If
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If a = True Then
        strSql = "Select Count(DISTINCT Lastnowo) AS xx From M_ListRemindercall WITH(NOLOCK) Where K_Telp = 1 And ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') Or (" & txtHelp & ")) And NoRangka in (" & addFilter(Bln, Thn) & ")" & txtCriteria
    Else
        strSql = "Select Count(DISTINCT Lastnowo) AS xx From M_ListRemindercall WITH(NOLOCK) Where ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)='" & Thn & "') Or (" & txtHelp & ")) And NoRangka in (" & addFilter(Bln, Thn) & ")" & txtCriteria
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
Dim txtHelp2 As String

    On Error GoTo errHandler
    If Bln = 1 Then
        txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        txtHelp2 = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
    Else
        txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        txtHelp2 = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
    End If
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If a = True Then
        strSql = "SELECT COUNT(DISTINCT NORANGKA) AS xx FROM (" & addFilter(Bln, Thn) & ") AS X"
    Else
        strSql = "SELECT COUNT(DISTINCT LASTNOWO) AS xx FROM (" & _
                 "Select DISTINCT LastNowo " & _
                 "From M_ListReminder WITH(NOLOCK) Where ((Month(TglProses)= '" & Bln & "' And Year(TglProses)='" & Thn & "') Or (" & txtHelp & ")) " & _
                 "And NoRangka in (" & addFilter(Bln, Thn) & ") " & _
                 "UNION " & _
                 "Select DISTINCT LastNowo From M_ListReminderCall WITH(NOLOCK) " & _
                 "Where ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)= '" & Thn & "') Or (" & txtHelp2 & ")) " & _
                 "And NoRangka in (" & addFilter(Bln, Thn) & ")) AS X"
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
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    If a = True Then
        strSql = "SELECT 0"
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
'        strSQL = "Select Count(DISTINCT LastNowo) From M_ListReminderCall Where ((Month(TglFollow)= '" & Bln & "' And Year(TglFollow)= '" & Thn & "') Or (" & txtHelp & ")) AND LastNoWo NOT IN (SELECT LastNoWo FROM M_LISTREMINDER) " & _
'                 "And NoRangka in (" & addFilter(Bln, Thn) & ")"
         strSql = "SELECT 0"
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
    If a Then
        If Bln = 1 Then
            txtHelp = "MONTH(TglFollow) = '12' AND YEAR(TglFollow) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglFollow) = '" & CDbl(Bln - 1) & "' AND YEAR(TglFollow) = '" & CDbl(Thn) & "'"
        End If
        ''' asli '''
        ''.:  7 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -7 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.: 14 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -14 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.: 21 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -21 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.: 30 HARI :.''
        strSql = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C WITH(NOLOCK) ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -30 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.: 90 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -90 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.: 150 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -150 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
    Else
        If Bln = 1 Then
            txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        End If
        ''' asli '''
        ''.:  7 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -7 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 14 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -14 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 21 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -21 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 30 HARI :.''
        strSql = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C WITH(NOLOCK) ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -30 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 90 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -90 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 150 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -150 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And NoRangka in (" & addFilter(Bln, Thn) & ")"
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
'        strSQL = "SELECT COUNT(DISTINCT NOFAK) FROM BOOK A WHERE (NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDER WHERE ((MONTH(TGLPROSES) = '" & Bln & "' AND YEAR(TGLPROSES) = " & Thn & ") OR (" & txtHelp1 & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLPROSES, 101),A.DATE_BOOK) <= 21  And NoRangka in (" & addFilter(Bln, Thn) & ")) " & _
'                 "OR NO_CHASSIS IN (SELECT NORANGKA FROM M_LISTREMINDERCALL WHERE ((MONTH(TGLFOLLOW) = '" & Bln & "' AND YEAR(TGLFOLLOW) = " & Thn & ") OR (" & txtHelp & ")) AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) >= 0 AND DATEDIFF(d,CONVERT(varchar(10), TGLFOLLOW, 101),A.DATE_BOOK) <= 21 And NoRangka in (" & addFilter(Bln, Thn) & "))) "
        strSql = "SELECT count(distinct y.nofak) FROM ( " & _
                 "SELECT DISTINCT NOFAK, A.NO_CHASSIS FROM BOOK A WITH(NOLOCK) LEFT JOIN M_LISTREMINDER B WITH(NOLOCK) ON A.NO_CHASSIS = B.NORANGKA " & _
                 "Where DateDiff(d, Convert(VarChar(10), B.TGLPROSES, 101), a.DATE_BOOK) >= 0 And DateDiff(d, Convert(VarChar(10), B.TGLPROSES, 101), a.DATE_BOOK) <= 30 " & _
                 "AND ((MONTH(TGLPROSES) = '" & Bln & "' AND YEAR(TGLPROSES) = " & Thn & ") OR (" & txtHelp1 & ")) " & _
                 "Union " & _
                 "SELECT DISTINCT NOFAK, A.NO_CHASSIS FROM BOOK A WITH(NOLOCK) LEFT JOIN M_LISTREMINDERCALL B WITH(NOLOCK) ON A.NO_CHASSIS = B.NORANGKA " & _
                 "Where DateDiff(d, Convert(VarChar(10), B.TglFollow, 101), a.DATE_BOOK) >= 0 And DateDiff(d, Convert(VarChar(10), B.TglFollow, 101), a.DATE_BOOK) <= 30 " & _
                 "AND ((MONTH(TGLFOLLOW) = '" & Bln & "' AND YEAR(TGLFOLLOW) = " & Thn & ") OR (" & txtHelp & "))) AS Y where y.No_Chassis in (" & addFilter(Bln, Thn) & ")"
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
        ''.:  7 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -7 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.: 14 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -14 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.: 21 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -21 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.:  30 HARI :.''
        strSql = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -30 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.:  90 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -90 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
        ''.:  150 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListRemindercall C ON B.NO_CHASSIS = C.NORANGKA WHERE K_TELP = 1 AND DATEDIFF(d,B.DATE,TGLFOLLOW) >= -150 AND DATEDIFF(d,B.DATE,TGLFOLLOW) <= 0 AND ((MONTH(tglFollow) = '" & Bln & "' AND YEAR(tglFollow) = '" & Thn & "') Or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' And NoRangka in (" & addFilter(Bln, Thn) & ") AND LASTNOWO NOT IN (Select LASTNOWO From M_ListReminder)"
    Else
        If Bln = 1 Then
            txtHelp = "MONTH(TglProses) = '12' AND YEAR(TglProses) = '" & CDbl(Thn - 1) & "'"
        Else
            txtHelp = "MONTH(TglProses) = '" & CDbl(Bln - 1) & "' AND YEAR(TglProses) = '" & CDbl(Thn) & "'"
        End If
        '''asli''''
        ''.:  7 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -7 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 14 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -14 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 21 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -21 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 30 HARI :.''
        strSql = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C WITH(NOLOCK) ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -30 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 90 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -90 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
        ''.: 150 HARI :.''
        'strSQL = "SELECT COUNT(distinct b.no_Chassis) FROM JOBDTL A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO INNER JOIN M_ListReminder C ON B.NO_CHASSIS = C.NORANGKA WHERE DATEDIFF(d,B.DATE,TGLPROSES) >= -150 AND DATEDIFF(d,B.DATE,TGLPROSES) <= 0 AND ((MONTH(TglProses) = '" & Bln & "' AND YEAR(TglProses) = '" & Thn & "') or (" & txtHelp & ")) AND DELIV <> 'BTL' AND FS_LOG_DEL = '' and FS_RC <> 'R001' and FS_RC <> '' AND TIPE <> '1K' And NoRangka in (" & addFilter(Bln, Thn) & ")"
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

Public Function KPI_12new(Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim Ket As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
    If Bln = 1 Then
        Ket = "MONTH(X.TglTransfer) = '12' AND YEAR(X.TglTransfer) = '" & CDbl(Thn - 1) & "'"
    Else
        Ket = "MONTH(X.TglTransfer) = '" & CDbl(Bln - 1) & "' AND YEAR(X.TglTransfer) = '" & CDbl(Thn) & "'"
    End If
    strSql = "Select Count(distinct a.SPK_NO) FROM BPK A WITH(NOLOCK) INNER JOIN AFI B ON A.SPK_NO = B.SPK_NO " & _
             "INNER JOIN SPK C WITH(NOLOCK) ON A.SPK_NO = C.SPK_NO " & _
             "INNER JOIN T_TransferCR1 X WITH(NOLOCK) On A.SPK_NO = X.NOSPK INNER JOIN T_Transfer Y WITH(NOLOCK) On A.SPK_NO = Y.NOSPK Where A.BATAL = 0 And Y.SBI1000 = 1 and " & Ket & ""
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_12new = rs
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

Public Function KPI_12bnew(Bln As Integer, Thn As Integer) As ADODB.Recordset
Dim Ket As String
Dim ket2 As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    If Bln = 1 Then
        Ket = "MONTH(X.TglTransfer) = '12' AND YEAR(X.TglTransfer) = '" & CDbl(Thn - 1) & "'"
        ket2 = "((MONTH(date) = '12' AND YEAR(date) = '" & CDbl(Thn - 1) & "') OR (MONTH(date) = '" & CDbl(Bln) & "' AND YEAR(date) = '" & CDbl(Thn) & "'))"
    Else
        Ket = "MONTH(X.TglTransfer) = '" & CDbl(Bln - 1) & "' AND YEAR(X.TglTransfer) = '" & CDbl(Thn) & "'"
        ket2 = "((MONTH(DATE) = '" & CDbl(Bln - 1) & "' AND YEAR(DATE) = '" & CDbl(Thn) & "') OR (MONTH(DATE) = '" & CDbl(Bln) & "' AND YEAR(DATE) = '" & CDbl(Thn) & "'))"
    End If
    strSql = "SELECT COUNT(DISTINCT A.SPK_NO) FROM " & Log.LogSPK & "..BPK A WITH(NOLOCK) INNER JOIN " & Log.LogSPK & "..AFI B WITH(NOLOCK) ON A.SPK_NO = B.SPK_NO INNER JOIN " & Log.LogSPK & "..SPK C WITH(NOLOCK) ON A.SPK_NO = C.SPK_NO " & _
             "INNER JOIN " & Log.LogSPK & "..T_TransferCR1 X WITH(NOLOCK) ON A.SPK_NO = X.NoSPK INNER JOIN " & Log.LogSPK & "..T_Transfer Y WITH(NOLOCK) ON A.SPK_NO = Y.NoSPK Where A.BATAL = 0 And Y.SBI1000 = 1 and " & Ket & " AND NORANGKA IN (SELECT NO_CHASSIS FROM " & Log.LogBengkel & "..JOBDTA WITH(NOLOCK) WHERE FS_RC = 'R001' AND DELIV <> 'BTL' AND " & ket2 & ")"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set KPI_12bnew = rs
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
      
    If Log.RPLHistori = 0 Then
        strSql = "Select b.FS_Nilai, Count(a.FS_RC) FROM T_Regular b WITH(NOLOCK) " & _
             "Inner Join JobDTA a WITH(NOLOCK) On b.FS_RC = a.FS_RC " & _
             "Where a.[date] between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' " & _
             "And  b.FS_Nilai < " & aMax
                If a = "0000" Then
                    strSql = strSql & " And (Right(CAST( b.FS_Nilai AS VARCHAR(10)),4) = '0000' Or Right(CAST( b.FS_Nilai AS VARCHAR(10)),4) = '1000')"
                End If
             strSql = strSql & " Group By b.FS_Nilai Order By b.Fs_Nilai"
    Else
        strSql = "Select b.FS_Nilai, Count(a.FS_RC) FROM [" & dbReplikasiHO & "]..T_Regular b WITH(NOLOCK) " & _
             "Inner Join JobDTA a WITH(NOLOCK) On b.FS_RC = a.FS_RC " & _
             "Where a.[date] between '" & TglNul2(tg1) & "' And '" & TglNul2(tg2) & "' " & _
             "And  b.FS_Nilai < " & aMax
                If a = "0000" Then
                    strSql = strSql & " And (Right(CAST( b.FS_Nilai AS VARCHAR(10)),4) = '0000' Or Right(CAST( b.FS_Nilai AS VARCHAR(10)),4) = '1000')"
                End If
             strSql = strSql & " Group By b.FS_Nilai Order By b.Fs_Nilai"
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
1         On Error GoTo errHandler
2         If mConn.State = adStateClosed Then mConn.Open (SQLCon)
          
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
           
3          strSql = "Select X.XSBE as SBE, Count(X.XSBE) as nSBE FROM( " & _
                    "select a.NoRangka as XRANGKA, CAST( a.SbeNExt AS VARCHAR(10)) as XSBE " & _
                    "FROM M_Estimasi a WITH(NOLOCK) " & _
                    "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
                    "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis " & _
                    "WHERE (a.JT_Tgl is Not Null And  a.SbeNExt < " & aMax & " "
4                     If a = "0000" Then
5                         strSql = strSql & " And Right(CAST( a.SbeNExt AS VARCHAR(10)),4) = '0000') "
6                     End If
7                   strSql = strSql & " And (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) " & _
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
8                     If a = "0000" Then
9                         strSql = strSql & " And Right(CAST( A.ASbe AS VARCHAR(10)),4) = '0000' "
10                    End If
11                  strSql = strSql & " ) Y On X.XRANGKA = Y.YRANGKA And X.XSBE = Y.YSBE " & _
                    "Group By X.XSBE Order By CAST(ROUND(x.XSBE, 0) AS INT)"
                    
12        If Log.RPLHistori = 1 Then
13            strSql = UCase(strSql)
14            strSql = Replace(strSql, "T_REGULAR", "[" & dbReplikasiHO & "]..T_REGULAR")
              
15        End If
                    
16        Set rs = New ADODB.Recordset
17        rs.CursorLocation = adUseClient
18        rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
19        Set rs.ActiveConnection = Nothing
20        Set GraphRat = rs
21        If mConn.State = adStateOpen Then mConn.Close
              
22        Exit Function
errHandler:
23    err.Raise err.Number, err.Source, err.Description
End Function


