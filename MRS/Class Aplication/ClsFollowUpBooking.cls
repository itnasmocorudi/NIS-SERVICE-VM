VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFollowUpBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Public Function FetchDataFollowUpHari(isTgl As Date, all As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    '' CAST(a.[Date] as CHAR(40))
    If all Then
        strSql = "Select a.Deliv,(Cast(DateDiff(d,a.Date_book,a.[date]) as CHAR(3))) + ((Case When c.NoFak Is Null then '--' Else " & _
                 "(Case When c.Proses = 0 then 'OK' When c.Proses = 1 then 'RS' When c.Proses = 2 then 'BT' When c.Proses = 3 then 'CA' END)End)) as FU, " & _
                 "a.NoFAK, CONVERT(char(12), a.[Date_book], 5) + ' ' + a.time_d as orderBook, CONVERT(char(12), a.[Date], 5) as Tanggal, CAST(a.[Time] as CHAR(40)) as Jam, a.No_Polisi , a.No_Chassis , " & _
                 "a.Customer , a.CUST_ADRES, a.Telpon, a.GSM, a.CDMA, CATATAN1+CATATAN2 AS CAT, A.FMDL " & _
                 "FROM BOOK a WITH(NOLOCK) Left Outer Join M_FuBook c WITH(NOLOCK) On c.NoFak = a.NoFak " & _
                 "WHERE A.NOFAK NOT IN (SELECT NOFAK FROM JOBDTA WITH(NOLOCK) WHERE FS_LOG_DEL = '') And a.[Date] = '" & FormatTgl(DateAdd("d", 1, isTgl)) & "' " & _
                 "AND A.NOFAK NOT IN (SELECT NOBOK FROM " & Log.LogQueue & "..DATABOOKING WITH(NOLOCK) WHERE AKTIV <> 0)"

    Else
        strSql = "Select a.Deliv,(Cast(DateDiff(d,a.Date_book,a.[date]) as CHAR(3))) + ((Case When c.NoFak Is Null then '--' Else " & _
                "(Case When c.Proses = 0 then 'OK' When c.Proses = 1 then 'RS' When c.Proses = 2 then 'BT' When c.Proses = 3 then 'CA' END)End)) as FU, " & _
                "a.NoFAK, CONVERT(char(12), a.[Date_book], 5) + ' ' + a.time_d as orderBook, CONVERT(char(12), a.[Date], 5) as Tanggal, CAST(a.[Time] as CHAR(40)) as Jam, a.No_Polisi , a.No_Chassis , " & _
                "a.Customer , a.CUST_ADRES, a.Telpon, a.GSM, a.CDMA, CATATAN1+CATATAN2 AS CAT, A.FMDL " & _
                "FROM BOOK a WITH(NOLOCK) Left Outer Join M_FuBook c WITH(NOLOCK) On c.NoFak = a.NoFak " & _
                "WHERE A.NOFAK NOT IN (SELECT NOFAK FROM JOBDTA WITH(NOLOCK) WHERE FS_LOG_DEL = '') And a.[Date] = '" & FormatTgl(DateAdd("d", 1, isTgl)) & "' And A.[DATE_BOOK] < '" & FormatTgl(DateAdd("d", 0, isTgl)) & "' " & _
                "AND A.NOFAK NOT IN (SELECT NOBOK FROM " & Log.LogQueue & "..DATABOOKING WITH(NOLOCK) WHERE AKTIV <> 0)"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataFollowUpHari = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataFollowUpJam(isTgl As Date, JamMin As Integer, JamMax As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select a.Deliv,(Cast(DateDiff(d,a.Date_book,a.[date]) as CHAR(3)))  + ((Case When c.NoFak Is Null then '--' Else 'OK' End)) as FU, " & _
             "a.NoFAK, CONVERT(char(12), a.[Date_book], 5) + ' ' + a.time_d as orderBook, CONVERT(char(12), a.[Date], 5) as Tanggal, CAST(a.[Time] as CHAR(40)) as Jam, a.No_Polisi , a.No_Chassis , " & _
             "a.Customer , a.CUST_ADRES, a.Telpon, a.GSM, a.CDMA, CATATAN1+CATATAN2 AS CAT, A.FMDL " & _
             "FROM BOOK a WITH(NOLOCK) Left Outer Join M_FuBook c WITH(NOLOCK) On c.NoFak = a.NoFak " & _
             "WHERE A.NOFAK NOT IN (SELECT NOFAK FROM JOBDTA WITH(NOLOCK) WHERE FS_LOG_DEL = '') AND A.NOFAK NOT IN (SELECT NOBOK FROM " & Log.LogQueue & "..DATABOOKING WITH(NOLOCK) WHERE AKTIV <> 0) " & _
             "And a.[Date] = '" & FormatTgl(isTgl) & "' And CAST((Left(a.[Time],2) + Right(a.[Time],2)) as INT) Between " & JamMin & " And " & JamMax
    
'    strSQL = "Select a.Deliv,(Cast(DateDiff(d,a.Date_book,a.[date]) as CHAR(3)))  + ((Case When c.NoFak Is Null then '--' Else 'OK' End)) as FU, " & _
'             "a.NoFAK, CONVERT(char(12), a.[Date], 5) as Tanggal, CAST(a.[Time] as CHAR(40)) as Jam, a.No_Polisi , a.No_Chassis , " & _
'             "a.Customer , a.CUST_ADRES, a.Telpon, a.GSM, a.CDMA, CATATAN1+CATATAN2 AS CAT, A.FMDL " & _
'             "FROM BOOK a Left Outer Join M_FuBook c On c.NoFak = a.NoFak Left Outer Join JobDTA b On a.NoFak = b.NoFak " & _
'             "WHERE IsNull(b.NoFak,'') = '' And a.[Date] = '" & FormatTgl(isTgl) & "' And CAST((Left(a.[Time],2) + Right(a.[Time],2)) as INT) Between " & JamMin & " And " & JamMax
        
'    strSQL = "Select NoFAK, CAST([Date] as CHAR(40)) as Tanggal, CAST([Time] as CHAR(40)) as Jam, No_Polisi, No_Chassis, Customer, CUST_ADRES, Telpon, GSM, CDMA FROM BOOK " & _
'             "WHERE NoFak <> 'BOOKED' And Deliv <> 'BTL' And [Date] = '" & FormatTgl(isTgl) & "' And CAST((Left([Time],2) + Right([Time],2)) as INT) Between " & JamMin & " And " & JamMax
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataFollowUpJam = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataPart(NoBooking As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    Dim xNisDb As String
    xNisDb = NisDb
'    strSQL = "Select Bengkel..Book_Part.KodeItem, Bengkel..Book_Part.NamaBarang,Bengkel..Book_Part.Qty as Qty_Pesan, IsNull([" & xNisDb & "]..SOORD.Qty,0) as Qty_SP, " & _
'             "(Case When [" & xNisDb & "]..SOORD.KodeItem Is Null then 'Tidak Tersedia' Else 'Tersedia' End) as Keterangan, " & _
'             "IsNull((Case When [" & xNisDb & "]..SOORD.Qty = Bengkel..Book_Part.Qty then 'Sesuai' " & _
'                          "When [" & xNisDb & "]..SOORD.Qty < Bengkel..Book_Part.Qty then 'Kurang' " & _
'                          "When [" & xNisDb & "]..SOORD.Qty > Bengkel..Book_Part.Qty then 'Salah' End),' ') as Kelengkapan " & _
'             "FROM Bengkel..Book INNER JOIN Bengkel..Book_Part On Bengkel..Book.NoFak = Bengkel..Book_Part.NoFak " & _
'             "Left Join [" & xNisDb & "]..SOORD On Bengkel..Book_Part.KodeItem = [" & xNisDb & "]..SOORD.kodeItem " & _
'             "INNER JOIN [" & xNisDb & "]..SOHDR ON Bengkel..Book.NoFak = [" & xNisDb & "]..SOHDR.NoPo " & _
'             "WHERE Bengkel..Book.NoFak = '" & NoBooking & "'"

    strSql = "Select " & gDb & "..Book_Part.KodeItem, " & gDb & "..Book_Part.NamaBarang," & gDb & "..Book_Part.Qty as Qty_Pesan, IsNull([" & xNisDb & "]..SOORD.Qty,0) as Qty_SP, " & _
             "(Case When [" & xNisDb & "]..SOORD.KodeItem Is Null then 'Tidak Tersedia' Else 'Tersedia' End) as Keterangan, " & _
             "IsNull((Case When [" & xNisDb & "]..SOORD.Qty = " & gDb & "..Book_Part.Qty then 'Sesuai' " & _
                          "When [" & xNisDb & "]..SOORD.Qty < " & gDb & "..Book_Part.Qty then 'Kurang' " & _
                          "When [" & xNisDb & "]..SOORD.Qty > " & gDb & "..Book_Part.Qty then 'Salah' End),' ') as Kelengkapan " & _
             "FROM " & gDb & "..Book WITH(NOLOCK) INNER JOIN " & gDb & "..Book_Part WITH(NOLOCK) On " & gDb & "..Book.NoFak = " & gDb & "..Book_Part.NoFak " & _
             "Left Join [" & xNisDb & "]..SOORD WITH(NOLOCK) On " & gDb & "..Book_Part.KodeItem = [" & xNisDb & "]..SOORD.kodeItem " & _
             "INNER JOIN [" & xNisDb & "]..SOHDR WITH(NOLOCK) ON " & gDb & "..Book.NoFak = [" & xNisDb & "]..SOHDR.NoPo " & _
             "WHERE " & gDb & "..Book.NoFak = '" & NoBooking & "'"
   
   
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataPart = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function Get_FU_Booking(NoBooking As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select * FROM M_FUBOOK WITH(NOLOCK) WHERE NoFak = '" & NoBooking & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set Get_FU_Booking = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchItemSPK(noRangka As String, noPol As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
    strSql = "Select SPK.SPK_No, BPK.Tgl_BPK as TglDO, T_MRS.TglCR1, T_MRS.KmCR1, T_MRS.TglCR2, T_MRS.KmCR2, " & _
             "SPK.SNama as UseName, SPK.SAlamat as UseAddr, SPK.SKota as UseCity, SPK.SPropinsi as UsePrpinsi,  SPK.SKodePos as UsePos, " & _
             "AFI.NoPol as NoPolisi, InMst.NoMesin as NoEngine, SPK.NoRangka as NoChassis,SPK.STelp as Telpon, " & _
             "SPK.SGsm as GSM, SPK.SCdma as CDMA, InKategory.KodeKategory as KodeType, " & _
             "InKategory.NamaKategory as NamaType,InWarnaExt.KodeWarna,  InWarnaExt.NamaEnglish as NamaWarna, " & _
             "ArSales.NamaSales as Wiraniaga, SPK.Nama_Faktur as NamaSTNK, " & _
             "(AFI.Alamat1 + AFI.Alamat2 + AFI.Alamat3 + AFI.Alamat4 + AFI.Alamat5) As AlamatSTNK, " & _
             "SPK.Kota as KotaSTNK, AFI.KodePos as KodePosSTNK, AFI.TmptLahir as TempatLahir, " & _
             "AFI.TglLahir as TanggalLahir, AFI.Agama, AFI.Pekerjaan, AFI.Hoby, " & _
             "SPK.NoRangka as No_Chassis, Spk.Kode_Lgn, AFI.Tahun " & _
             "FROM AFI WITH(NOLOCK) INNER JOIN SPK WITH(NOLOCK) ON AFI.SPK_No = SPK.SPK_No INNER JOIN " & _
                  "INMerek WITH(NOLOCK) ON SPK.Kode_Merk = INMerek.KodeMerek INNER JOIN " & _
                  "InKategory WITH(NOLOCK) ON SPK.Kode_Type = InKategory.KodeKategory INNER JOIN " & _
                  "InWarnaExt ON SPK.Kode_Warna = InWarnaExt.KodeWarna INNER JOIN " & _
                  "BPK WITH(NOLOCK) ON AFI.SPK_No = BPK.SPK_No INNER JOIN " & _
                  "ArSales WITH(NOLOCK) ON SPK.Kode_Sales = ArSales.KodeSales INNER JOIN " & _
                  "INMst WITH(NOLOCK) ON SPK.NoRangka = InMst.kodeItem INNER JOIN " & _
                  "T_MRS WITH(NOLOCK) ON SPK.SPK_No = T_MRS.NoSPK " & _
             "Where SPK.NoRangka = '" & noRangka & "' Or AFI.NoPol = '" & noPol & "'"
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchItemSPK = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_FUBOOK(mNoBooking As String, mProses As Integer, mAlasan As String)
    Dim cm As New Command
    On Error GoTo exErr
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
        strSql = "Insert Into M_FUBOOK " & _
                           "Values( '" & mNoBooking & "' , " & mProses & " , GetDate(), '" & mAlasan & "') "

        cm.CommandText = strSql
        cm.CommandType = adCmdText
        
        Set cm.ActiveConnection = mConn
        cm.Execute
        Set cm = Nothing
        If mConn.State = adStateOpen Then mConn.Close

    Exit Sub
exErr:
    If Not cm Is Nothing Then Set cm = Nothing
    If err.Number = -2147217873 Then
        strSql = "Update M_FUBOOK Set Proses = " & mProses & ", Waktu = GetDate(), Alasan = '" & mAlasan & "'" & _
                           "Where NoFak = '" & mNoBooking & "'"

        cm.CommandText = strSql
        cm.CommandType = adCmdText
        
        Set cm.ActiveConnection = mConn
        cm.Execute
        Set cm = Nothing
        If mConn.State = adStateOpen Then mConn.Close
        Exit Sub
    End If
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub Update_DelivBOOK(Index As Integer, mNoBooking As String)
    Dim cm As New Command
    On Error GoTo exErr
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      Select Case Index
         Case 3
            If mConn.State = adStateOpen Then mConn.Close
            Exit Sub
         Case 2
            strSql = "Update BOOK Set Deliv = 'BTL', FS_LOG_DEL = '" & gSA & "', FD_TGL_DEL = getdate() Where NoFak = '" & mNoBooking & "'"
            
            Dim rsBooK As New ADODB.Recordset
            Set rsBooK = mConn.Execute("SELECT * fROM BOOK WITH(NOLOCK) WHERE NOFAK = '" & mNoBooking & "'")
            If Not rsBooK.EOF Then
                Call updateAntrian(Trim(rsBooK("NO_POLISI")), Trim(rsBooK("FMDL")), Trim(rsBooK("CUSTOMER")), Trim(rsBooK("DATE")), Trim(rsBooK("TIME")), 3, mNoBooking)
            End If
            If rsBooK.State <> adStateClosed Then rsBooK.Close
         Case 1
            Dim rsDeliv As ADODB.Recordset
            Set rsDeliv = Get_DelivBOOK(mNoBooking)
            Dim Flg As String
            Dim Jml As Integer
            If rsDeliv.Fields(0).value = "BOK" Or rsDeliv.Fields(0).value = "BTL" Then
               Flg = "RES1"
            Else
               Jml = Val(Right(rsDeliv.Fields(0).value, 1))
               Flg = Trim("RES" & Val(Jml + 1))
            End If
            strSql = "Update BOOK Set Deliv = '" & Flg & "', FS_LOG_ADD = '" & gSA & "' Where NoFak = '" & mNoBooking & "'"
         Case 0
            If mConn.State = adStateOpen Then mConn.Close
            Exit Sub
      End Select
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        cm.CommandText = strSql
        cm.CommandType = adCmdText
        
        Set cm.ActiveConnection = mConn
        cm.Execute
        Set cm = Nothing
        If Index = 2 Then
            Call Update_StallDelivBOOK(mNoBooking)
        End If
        If mConn.State = adStateOpen Then mConn.Close
        
      Exit Sub
exErr:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub Update_StallDelivBOOK(mNoBooking As String)
    Dim cm As New Command
    On Error GoTo exErr
        strSql = "UPDATE STALL SET FLAG = 'X' WHERE KETERANGAN LIKE  '%|" & mNoBooking & "|%'"
        If mConn.State = adStateClosed Then mConn.Open (SQLCon)
        cm.CommandText = strSql
        cm.CommandType = adCmdText
        
        Set cm.ActiveConnection = mConn
        cm.Execute
        Set cm = Nothing
        If mConn.State <> adStateClosed Then mConn.Close
      Exit Sub
exErr:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function Get_DelivBOOK(mNoBooking As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select Deliv FROM BOOK WITH(NOLOCK) WHERE NoFak = '" & mNoBooking & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set Get_DelivBOOK = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

