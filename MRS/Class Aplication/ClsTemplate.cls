VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSQL As String

Public Function FetchDataTemplate() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "SELECT * FROM M_Template WITH(NOLOCK) Order By ID ASC"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataTemplate = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_TemplateSMS(mNo As Integer, mTypeSms As String, mPesan1 As String)
   On Error GoTo ErrAddTemplateSMS
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSQL = "Insert Into M_Template " & _
               "Values( " & mNo & " , '" & mTypeSms & "' , '" & mPesan1 & "')"
                                  
      cm.CommandText = strSQL
      cm.CommandType = adCmdText
      
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddTemplateSMS:
      If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub Update_TemplateSMS(mNo As Integer, mTypeSms As String, mPesan1 As String)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "Update M_Template set TypeSMS='" & mTypeSms & "', " & _
                                    "Pesan1='" & mPesan1 & "' " & _
                                    "Where Id = " & mNo & " "
             
    cm.CommandText = strSQL
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

Public Sub Delete_Item_TemplateSMS(mNo As Integer)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "DELETE FROM M_Template Where Id = " & mNo & " "
             
    cm.CommandText = strSQL
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


'=============================================================================================
'=============================================================================================

Public Function FetchDataInbox() As ADODB.Recordset
    On Error GoTo errHandler
    'SMS... If mConn.State = adStateClosed Then mConn.Open (AccessCon)
    'SMS... strSQL = "SELECT tbinbox.ibx_id  as Id, tbinbox.ibx_pesan  as pesan, " & _
             "IIF(isnull(tbperson.psn_nama) = true ,tbinbox.ibx_sender, " & _
             "tbperson.psn_nama)  as sender, tbinbox.ibx_waktu as waktu, " & _
             "tbinbox.ibx_milik as milik, " & _
             "IIF( tbinbox.ibx_status = 0, 'Belum Dibaca', 'Sudah Dibaca' ) as status " & _
             "FROM tbperson right outer Join tbinbox " & _
             "On tbperson.psn_notel = tbinbox.ibx_sender " & _
             "order by tbinbox.ibx_waktu desc"
    If mConn.State = adStateClosed Then mConn.Open (SQLConSMS)
    strSQL = "SELECT [id] as ID, pesan, nomor as sender, Tanggal as Waktu, 'Inbox' as milik, case when status = 1 then 'Belum Dibaca' else 'Sudah Dibaca' end  as status " & _
             "FROM INBOX WITH(NOLOCK) Order by Tanggal desc"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    'SMS... rs.Open strSQL, AccessCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    rs.Open strSQL, SQLConSMS, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataInbox = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataInboxRange(Tgl1 As Date, Tgl2 As Date) As ADODB.Recordset
    On Error GoTo errHandler
    'SMS... If mConn.State = adStateClosed Then mConn.Open (AccessCon)
    'SMS... strSQL = "SELECT ibx_id as Id, ibx_pesan as pesan, ibx_sender as sender, ibx_waktu as waktu, ibx_milik as milik, IIF( ibx_status = 0, 'Belum Dibaca', 'Sudah Dibaca' ) as status " & _
             "FROM tbInbox WHERE ibx_waktu Between #" & FormatTgl(Tgl1) & "# And #" & FormatTgl(DateAdd("d", 1, Tgl2)) & "# Order By ibx_waktu DESC"
    
    If mConn.State = adStateClosed Then mConn.Open (SQLConSMS)
    strSQL = "SELECT [id] as ID, pesan, nomor as sender, Tanggal as Waktu, 'Inbox' as milik, case when status = 1 then 'Belum Dibaca' else 'Sudah Dibaca' end  as status " & _
             "FROM INBOX WITH(NOLOCK) Where Tanggal Between '" & TglNul2(Tgl1) & " 00:00:00.001' And '" & TglNul2(Tgl2) & " 23:59:59.999' Order by Tanggal desc"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    'SMS... rs.Open strSQL, AccessCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    rs.Open strSQL, SQLConSMS, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataInboxRange = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub Update_ReadSMS(mNo As Integer)
On Error GoTo errHandler
    Dim cm As New Command
    'SMS... If mConn.State = adStateClosed Then mConn.Open (AccessCon)
    'SMS... IstrSQL = "Update tbinbox set ibx_status = 1 Where ibx_Id = " & mNo
    
    If mConn.State = adStateClosed Then mConn.Open (SQLConSMS)
    strSQL = "update Inbox set Status = 0 Where id = " & mNo
    
             
    cm.CommandText = strSQL
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
'=============================================================================================
'=============================================================================================

Public Function FetchDataSendItem() As ADODB.Recordset
    On Error GoTo errHandler
    'SMS... If mConn.State = adStateClosed Then mConn.Open (AccessCon)
    'SMS... strSQL = "SELECT tbSendItem.si_id as Id, tbSendItem.si_pesan as Pesan, " & _
             "IIF(isnull(tbperson.psn_nama) = true ,tbSendItem.si_notuj, tbperson.psn_nama)  as NoTujuan, " & _
             "tbSendItem.si_waktu as waktu, tbSendItem.si_milik as Milik " & _
             "FROM tbperson right outer Join tbSendItem " & _
             "On tbperson.psn_notel = tbSendItem.si_notuj " & _
             "Order By tbSendItem.si_waktu DESC"
    
    If mConn.State = adStateClosed Then mConn.Open (SQLConSMS)
    strSQL = "Select [id] as id, Pesan, Tujuan as NoTujuan, Tanggal as Waktu, Modul as Milik FROM SendItem WITH(NOLOCK) Order By Tanggal desc"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    'SMS... rs.Open strSQL, AccessCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    rs.Open strSQL, SQLConSMS, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataSendItem = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataSendItemRange(Tgl1 As Date, Tgl2 As Date) As ADODB.Recordset
    On Error GoTo errHandler
    'SMS... If mConn.State = adStateClosed Then mConn.Open (AccessCon)
    'SMS... strSQL = "SELECT si_id as Id, si_pesan as Pesan, si_notuj as NoTujuan, si_waktu as waktu, si_milik as Milik " & _
             "FROM tbSendItem WHERE si_waktu Between #" & FormatTgl(Tgl1) & "# And #" & FormatTgl(DateAdd("d", 1, Tgl2)) & "# Order By si_waktu DESC"
    
    If mConn.State = adStateClosed Then mConn.Open (SQLConSMS)
    strSQL = "Select [id] as id, Pesan, Tujuan as NoTujuan, Tanggal as waktu, Modul as Milik FROM SendItem WITH(NOLOCK) " & _
             "WHERE Tanggal Between '" & TglNul2(Tgl1) & " 00:00:00.001' And '" & TglNul2(Tgl2) & " 23:59:59.999' Order By Tanggal desc"
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    'SMS... rs.Open strSQL, AccessCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    rs.Open strSQL, SQLConSMS, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataSendItemRange = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

'=============================================================================================
'=============================================================================================

Public Function FetchDataStatistik() As ADODB.Recordset
    On Error GoTo errHandler
    
    'SMS... If mConn.State = adStateClosed Then mConn.Open (AccessCon)
    'SMS... strSQL = "TRANSFORM Count(tbsenditem.si_id) AS CountOfsi_id SELECT tbsenditem.si_milik as Users, " & _
             "Count(tbsenditem.si_id) AS [Total] FROM tbsenditem GROUP BY tbsenditem.si_milik " & _
             "PIVOT Format([si_waktu],'mm') In ('01','02','03','04','05','06','07','08','09','10','11','12')"
             '"PIVOT Format([si_waktu],'mmm') In ('Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agust','Sep','Okt','noP','Des')"

    If mConn.State = adStateClosed Then mConn.Open (SQLConSMS)
    strSQL = "SELECT Modul as Users,  Month(Tanggal) ,Count(id) AS [Total] FROM senditem WITH(NOLOCK) GROUP BY Modul, Month(Tanggal)"
             

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    'SMS... rs.Open strSQL, AccessCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    rs.Open strSQL, SQLConSMS, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataStatistik = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchDataStatistikGrafik() As ADODB.Recordset
    On Error GoTo errHandler
    'SMS... If mConn.State = adStateClosed Then mConn.Open (AccessCon)
    'SMS... strSQL = "TRANSFORM Count(tbsenditem.si_id) AS CountOfsi_id SELECT " & _
             "Count(tbsenditem.si_id) AS [Total] FROM tbsenditem GROUP BY tbsenditem.si_milik " & _
             "PIVOT Format([si_waktu],'mmm') In ('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')"
    
    If mConn.State = adStateClosed Then mConn.Open (SQLConSMS)
    strSQL = "TRANSFORM Count([id]) AS CountOfsi_id SELECT Modul as Users, " & _
             "Count([id]) AS [Total] FROM senditem WITH(NOLOCK) GROUP BY senditem.Modul " & _
             "PIVOT Format([Tanggal],'mmm') In ('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    'SMS... rs.Open strSQL, AccessCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    rs.Open strSQL, SQLConSMS, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataStatistikGrafik = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchPulsa() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "SELECT * FROM M_Pulsa WITH(NOLOCK) "

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchPulsa = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub Update_TipePulsa(tipe As Integer)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "Update M_Pulsa Set Tipe = " & tipe
             
    cm.CommandText = strSQL
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

Public Sub Update_NilaiPulsa(Nilai As Double, Tarif As Double)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSQL = "Update M_Pulsa Set Pulsa = " & Nilai & ", Tarif = " & Tarif
             
    cm.CommandText = strSQL
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
Public Function GetIrr_SBI(a As Boolean) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    If a Then 'Data SMS
        strSQL = "Select Count(*) From " & Log.LogSPK & "..T_MRS WITH(NOLOCK) WHERE NoRangka Not In (SELECT NoRangka FROM " & Log.LogBengkel & "..M_LISTREMINDER WITH(NOLOCK) WHERE TIPE = '1K') AND TGL1 BETWEEN '02/01/2010' AND '" & TglNul2(DateAdd("d", 6, Format(isnow, "DD/MM/YYYY"))) & "'"
    Else ' Data Call
        strSQL = "Select Count(*) From " & Log.LogSPK & "..T_MRS WITH(NOLOCK) WHERE NoRangka Not In (SELECT NoRangka FROM " & Log.LogBengkel & "..M_LISTREMINDERCALL WITH(NOLOCK) WHERE LastNoWO In (SELECT NoSpk FROM " & Log.LogSPK & "..T_MRS WITH(NOLOCK))) AND TGL1 BETWEEN '02/01/2010' AND '" & TglNul2(DateAdd("d", 2, Format(isnow, "DD/MM/YYYY"))) & "'"
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    GetIrr_SBI = rs.Fields(0).value
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
    GetIrr_SBI = 0
    If mConn.State = adStateOpen Then mConn.Close
End Function

Public Function GetIrr_SBE(a As Boolean) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    Dim TglNext0 As Date
    Dim TglNext1 As Date
    
    TglNext1 = TglNul2(DateAdd("d", IIf(a, 7, 6), Format(isnow, "DD/MM/YYYY")))
    TglNext0 = DateAdd("m", -30, isnow)
    If a Then 'Data SMS
        strSQL = "sELECT Count(X.NoChassis) fROM (SELECT a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
             "(CASE WHEN A.JT_Tgl > A.JT_Bln then A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then A.Jt_Bln END) as DateNextSvc, " & _
             "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, b.No_Engine as NoEngine, " & _
             "a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, " & _
             "b.FS_RC FROM M_ESTIMASI A WITH(NOLOCK) Inner Join " & _
             "(SELECT G.NORANGKA, (SELECT MAX(SBENEXT) FROM m_estimasi WITH(NOLOCK) WHERE NORANGKA=G.NORANGKA AND JT_TGL=MAX(G.JT_TGL)) S, MAX(G.JT_TGL) AS JT_TGL " & _
             "FROM M_ESTIMASI G WITH(NOLOCK) Where G.JT_TGL Is Not Null GROUP BY G.NORANGKA) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S " & _
             "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
             "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis Where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "') X " & _
             "WHERE X.NoChassis NOT IN (SELECT NoRangka FROM M_ListReminder WITH(NOLOCK) WHERE TIPE = 'SM')"
      
    Else ' Data Call
        strSQL = "sELECT Count(X.NoChassis) fROM (SELECT a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
             "(CASE WHEN A.JT_Tgl > A.JT_Bln then A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then A.Jt_Bln END) as DateNextSvc, " & _
             "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, b.No_Engine as NoEngine, " & _
             "a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, " & _
             "b.FS_RC FROM M_ESTIMASI A WITH(NOLOCK) Inner Join " & _
             "(SELECT G.NORANGKA, (SELECT MAX(SBENEXT) FROM m_estimasi WITH(NOLOCK) WHERE NORANGKA=G.NORANGKA AND JT_TGL=MAX(G.JT_TGL)) S, MAX(G.JT_TGL) AS JT_TGL " & _
             "FROM M_ESTIMASI G WITH(NOLOCK) Where G.JT_TGL Is Not Null GROUP BY G.NORANGKA) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S " & _
             "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' " & _
             "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis Where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "') X " & _
             "WHERE X.NoChassis NOT IN (SELECT NoRangka FROM M_ListReminderCall)"
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    GetIrr_SBE = rs.Fields(0).value
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
    GetIrr_SBE = 0
    If mConn.State = adStateOpen Then mConn.Close
End Function
