VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDataPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private mSConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Private mNoSPK  As String

Private mKodeLGN As String

Private mNama As String
Private mAlamat As String
Private mKodePos As String
Private mKota As String
Private mPropinsi As String
Private mNoPolisi As String
Private mRumah As String
Private mGSM As String
Private mCDMA As String

Public Property Let Nama(value As String)
   mNama = value
End Property
Public Property Let Alamat(value As String)
   mAlamat = value
End Property
Public Property Let KodePos(value As String)
   mKodePos = value
End Property
Public Property Let Kota(value As String)
   mKota = value
End Property
Public Property Let Propinsi(value As String)
   mPropinsi = value
End Property
Public Property Let NoPolisi(value As String)
   mNoPolisi = value
End Property
Public Property Let Rumah(value As String)
   mRumah = value
End Property
Public Property Let GSM(value As String)
   mGSM = value
End Property
Public Property Let CDMA(value As String)
   mCDMA = value
End Property
'-----------------------------------------------



Public Property Let NoSPK(value As String)
    mNoSPK = value
End Property


Public Property Let kodeLanggan(value As String)
    mKodeLGN = value
End Property


Public Function FetchDataPelanggan() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    Dim keteranganHide As String
    If Log.hideService = 1 Then
        keteranganHide = " where A.FS_KM_AKHIR > 10 "
    End If
    strSql = "Select FS_TH_Mobil as Thn_Mobil, UseName as Nama_Customer, Useaddr as Alamat_Customer, Usecity as Kota, usepos as KodePos, usepstn as Telpon, usegsm as GSM, usecdma as CDMA, " & _
             "No_Polisi, No_Engine, No_Chassis, MDL, Model, FMDL, Color, USEAGAMA AS agama, FS_KD_Dealer as Kode_Dealer, FS_NM_Dealer as Nama_Dealer, " & _
             "Job, JobCD, FS_KM_Akhir as KM_Akhir, USEJobDSC as Job_Desc,FS_KD_PLG as Kode_Langgan, (SELECT COUNT(*) FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS) AS JMLHISTORY, " & _
             "ISNULL((SELECT TOP 1 CASE ISNULL(DATEDIFF(day, [DATE], GETDATE()),0) WHEN '' THEN 0 ELSE ISNULL(DATEDIFF(day, [DATE], GETDATE()),0) END FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS ORDER BY [DATE] DESC),0) AS LASTSVC, " & _
             "ISNULL((SELECT TOP 1 [DATE] FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS ORDER BY [DATE] DESC),0) AS TGLLASTSVC, " & _
             "ISNULL((SELECT TOP 1 [DATE] FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS AND FS_RC <> '' ORDER BY [DATE] DESC),0) AS LASTSBE, " & _
             "CONVERT(float, FS_KM_AKHIR) AS KMAKHIR FROM CustDat A WITH(NOLOCK) Order By FS_KD_PLG desc"
    ''20200925 || request penambahan baru, utk menu data Pelanggan di MRA
    ''Ketika export ke excel/CSV untuk pengambilan data custnya mirip dengan mapping retention, dengan link ke SID
    ''Sehingga data alamat bisa didapatkan field pemisahan alamat, kelurahan, kecamatan, kota, propinsi.

    strSql = "Select a.FS_TH_Mobil as Thn_Mobil, a.UseName as Nama_Customer,isnull(c.Propinsi,'') as propinsi,ISNULL(c.Kota,'') kota,ISNULL(c.kecamatan,'') as kecamatan,ISNULL(c.Kelurahan,'') AS Kelurahan, ISNULL(c.Alamat,'') as Alamat_Customer, a.Usecity as Kota, a.usepos as KodePos, a.usepstn as Telpon, a.usegsm as GSM, a.usecdma as CDMA, "
    strSql = strSql & " a.No_Polisi, a.No_Engine, a.No_Chassis, a.MDL, a.Model, a.FMDL, a.Color, a.USEAGAMA AS agama, a.FS_KD_Dealer as Kode_Dealer, a.FS_NM_Dealer as Nama_Dealer,"
    strSql = strSql & " a.Job, a.JobCD, a.FS_KM_Akhir as KM_Akhir, USEJobDSC as Job_Desc,FS_KD_PLG as Kode_Langgan, (SELECT COUNT(*) FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS) AS JMLHISTORY,"
    strSql = strSql & " ISNULL((SELECT TOP 1 CASE ISNULL(DATEDIFF(day, [DATE], GETDATE()),0) WHEN '' THEN 0 ELSE ISNULL(DATEDIFF(day, [DATE], GETDATE()),0) END FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS ORDER BY [DATE] DESC),0) AS LASTSVC,"
    strSql = strSql & " ISNULL((SELECT TOP 1 [DATE] FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS ORDER BY [DATE] DESC),0) AS TGLLASTSVC,"
    strSql = strSql & " ISNULL((SELECT TOP 1 [DATE] FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS AND FS_RC <> '' ORDER BY [DATE] DESC),0) AS LASTSBE,"
    strSql = strSql & " CONVERT(Float, a.FS_KM_AKHIR) As KMAKHIR"
    strSql = strSql & " FROM CustDat A WITH(NOLOCK)"
    strSql = strSql & " LEFT JOIN SingleIDOri..UnitCustomer b WITH(NOLOCK) on b.NoRangka = a.NO_CHASSIS"
    strSql = strSql & " left join SingleIDOri..NewCustomer c WITH(NOLOCK) on c.IDCust = case when b.IDCustService <> '' then b.IDCustService when b.IDCustSales <> '' then b.IDCustSales when b.IDCustSV <> '' then b.IDCustSV end"
    strSql = strSql & keteranganHide & " Order By FS_KD_PLG desc"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenDynamic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataPelanggan = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


'==========================================================================================
'======================= TRANSFER =========================================================
'==========================================================================================

Public Function GetAllTransfer() As ADODB.Recordset
        Dim cm As New Command
On Error GoTo errHandler

    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
'    strSQL = "Select AFI.SPK_No, SPK.NoRangka, SPK.SNama, SPK.SAlamat, SPK.SKodePos, " & _
'         "SPK.SKota, SPK.SPropinsi, SPK.STelp, " & _
'         "SPK.SGsm, SPK.SCdma, InKategory.kodeKategory as KodeType, InKategory.NamaKategory as NamaType, " & _
'         "BPK.BPK_No as NoDo, Cast(BPK.Tgl_BPK as Char(12)) as TglDO, Cast(T_MRS.TglCR1 as Char(12))as TglCr1, " & _
'         "CAST(T_MRS.KmCR1 as Char(12)) as KmCR1 , Cast(T_MRS.TglCR2 as Char(12)) as TglCR2, Cast(T_MRS.KmCR2 as Char(12)) as KmCR2, " & _
'         "('Sales : ' + AFI.KetSalesCr2 + ' - CRC : ' + AFI.KetCrcCr2)as StatusCR2, " & _
'         "Cast(T_Transfer.TglTransfer as Char(12)) As Tanggal_Transfer, DATEDIFF(day,T_MRS.TglCR1,T_MRS.TglCR2) AS CR2_CR1, DATEDIFF(day,T_MRS.TglCR2,T_Transfer.TglTransfer) AS CR2_TRANSF " & _
'         "FROM BPK INNER JOIN " & _
'         "AFI On BPK.SPK_No = AFI.SPK_No INNER Join " & _
'         "T_MRS On AFI.SPK_No = T_MRS.NoSPK INNER Join " & _
'         "T_Transfer ON AFI.SPK_No = T_Transfer.NoSPK INNER JOIN " & _
'         "SPK ON AFI.SPK_No = SPK.SPK_No INNER JOIN " & _
'         "InKategory ON SPK.Kode_Type = InKategory.KodeKategory " & _
'         "Where BPK.Batal = 0 And Afi.User_Proses <> 'Null' " & _
'         "Order By T_Transfer.TglTransfer ASC"
    strSql = "Select top 500 AFI.SPK_No, SPK.NoRangka, SPK.SNama, SPK.SAlamat, SPK.SKodePos, " & _
         "SPK.SKota, SPK.SPropinsi, SPK.STelp, " & _
         "SPK.SGsm, SPK.SCdma, InKategory.kodeKategory as KodeType, InKategory.NamaKategory as NamaType, " & _
         "BPK.BPK_No as NoDo, BPK.Tgl_BPK as TglDO, T_MRS.TglCR1 as TglCr1, " & _
         "CAST(T_MRS.KmCR1 as Char(12)) as KmCR1 , T_MRS.TglCR2  as TglCR2, Cast(T_MRS.KmCR2 as Char(12)) as KmCR2, " & _
         "('Sales : ' + AFI.KetSalesCr2 + ' - CRC : ' + AFI.KetCrcCr2)as StatusCR2, " & _
         "T_Transfer.TglTransfer  As Tanggal_Transfer, DATEDIFF(day,T_MRS.TglCR1,T_MRS.TglCR2) AS CR2_CR1, DATEDIFF(day,T_MRS.TglCR2,T_Transfer.TglTransfer) AS CR2_TRANSF, TGL1, NAMAIND, AFI.NOPOL as nopol " & _
         "FROM BPK WITH(NOLOCK) INNER JOIN " & _
         "AFI WITH(NOLOCK) On BPK.SPK_No = AFI.SPK_No INNER Join " & _
         "T_MRS WITH(NOLOCK) On AFI.SPK_No = T_MRS.NoSPK INNER Join " & _
         "T_Transfer WITH(NOLOCK) ON AFI.SPK_No = T_Transfer.NoSPK INNER JOIN " & _
         "SPK WITH(NOLOCK) ON AFI.SPK_No = SPK.SPK_No INNER JOIN " & _
         "InKategory WITH(NOLOCK) ON SPK.Kode_Type = InKategory.KodeKategory " & _
         "INNER JOIN InMst WITH(NOLOCK) ON InMst.KodeItem = SPK.NoRangka " & _
         "INNER JOIN InWarnaExt WITH(NOLOCK) ON InWarnaExt.KodeWarna = InMst.KodeWarna " & _
         "Where BPK.Batal = 0 " & _
         "Order By T_Transfer.TglTransfer desc"
         'And Afi.User_Proses <> 'Null' "  >> dihapuskan, krn kirim afi ke HO tdk ada hubgn dgn clm
    cm.CommandText = strSql
    cm.CommandType = adCmdText

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set cm.ActiveConnection = mConn
    rs.Open cm, , adOpenDynamic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set GetAllTransfer = rs
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Function

Public Function GetAllTransferRange(Tgl1 As Date, Tgl2 As Date, ChInc As Boolean) As ADODB.Recordset
        Dim cm As New Command
On Error GoTo errHandler

    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    If ChInc = False Then
'    strSQL = "Select   AFI.SPK_No, SPK.NoRangka, SPK.SNama, SPK.SAlamat, SPK.SKodePos, " & _
'         "SPK.SKota, SPK.SPropinsi, SPK.STelp, " & _
'         "SPK.SGsm, SPK.SCdma, InKategory.kodeKategory as KodeType, InKategory.NamaKategory as NamaType, " & _
'         "BPK.BPK_No as NoDo, Cast(BPK.Tgl_BPK as Char(12)) as TglDO, Cast(T_MRS.TglCR1 as Char(12))as TglCr1, " & _
'         "CAST(T_MRS.KmCR1 as Char(12)) as KmCR1 , Cast(T_MRS.TglCR2 as Char(12)) as TglCR2, Cast(T_MRS.KmCR2 as Char(12)) as KmCR2, " & _
'         "('Sales : ' + AFI.KetSalesCr2 + ' - CRC : ' + AFI.KetCrcCr2)as StatusCR2, " & _
'         "Cast(T_Transfer.TglTransfer as Char(12)) As Tanggal_Transfer " & _
'         "FROM BPK INNER JOIN " & _
'         "AFI On BPK.SPK_No = AFI.SPK_No INNER Join " & _
'         "T_MRS On AFI.SPK_No = T_MRS.NoSPK INNER Join " & _
'         "T_Transfer ON AFI.SPK_No = T_Transfer.NoSPK INNER JOIN " & _
'         "SPK ON AFI.SPK_No = SPK.SPK_No INNER JOIN " & _
'         "InKategory ON SPK.Kode_Type = InKategory.KodeKategory " & _
'         "Where BPK.Batal = 0 And Afi.User_Proses <> 'Null' " & _
'         "AND T_Transfer.TglTransfer Between '" & FormatTgl(tgl1) & "' AND '" & FormatTgl(tgl2) & "' " & _
'         "AND T_MRS.TglCR2 Between '" & FormatTgl(tgl1) & "' AND '" & FormatTgl(tgl2) & "' " & _
'         "Order By T_Transfer.TglTransfer ASC"
    strSql = "Select   AFI.SPK_No, SPK.NoRangka, SPK.SNama, SPK.SAlamat, SPK.SKodePos, " & _
         "SPK.SKota, SPK.SPropinsi, SPK.STelp, " & _
         "SPK.SGsm, SPK.SCdma, InKategory.kodeKategory as KodeType, InKategory.NamaKategory as NamaType, " & _
         "BPK.BPK_No as NoDo, BPK.Tgl_BPK as TglDO, T_MRS.TglCR1 as TglCr1, " & _
         "CAST(T_MRS.KmCR1 as Char(12)) as KmCR1 , T_MRS.TglCR2 as TglCR2, Cast(T_MRS.KmCR2 as Char(12)) as KmCR2, " & _
         "('Sales : ' + AFI.KetSalesCr2 + ' - CRC : ' + AFI.KetCrcCr2)as StatusCR2, " & _
         "T_TransferCR1.TglTransfer  As Tanggal_Transfer, tgl1, namaind " & _
         "FROM BPK WITH(NOLOCK) INNER JOIN " & _
         "AFI WITH(NOLOCK) On BPK.SPK_No = AFI.SPK_No INNER Join " & _
         "T_MRS WITH(NOLOCK) On AFI.SPK_No = T_MRS.NoSPK INNER Join " & _
         "T_TransferCR1 WITH(NOLOCK) ON AFI.SPK_No = T_TransferCR1.NoSPK INNER JOIN " & _
         "SPK WITH(NOLOCK) ON AFI.SPK_No = SPK.SPK_No INNER JOIN " & _
         "InKategory WITH(NOLOCK) ON SPK.Kode_Type = InKategory.KodeKategory " & _
         "INNER JOIN InMst WITH(NOLOCK) ON InMst.KodeItem = SPK.NoRangka " & _
         "INNER JOIN InWarnaExt WITH(NOLOCK) ON InWarnaExt.KodeWarna = InMst.KodeWarna " & _
         "Where BPK.Batal = 0 And Afi.User_Proses <> 'Null' " & _
         "AND T_TransferCR1.TglTransfer Between '" & FormatTgl(Tgl1) & "' AND '" & FormatTgl(Tgl2) & "' "

'         "AND T_MRS.TglCR2 Between '" & FormatTgl(Tgl1) & "' AND '" & FormatTgl(Tgl2) & "' "
         strSql = strSql & " Order By T_TransferCR1.TglTransfer DESC"
    Else
'    strSQL = "Select   AFI.SPK_No, SPK.NoRangka, SPK.SNama, SPK.SAlamat, SPK.SKodePos, " & _
'         "SPK.SKota, SPK.SPropinsi, SPK.STelp, " & _
'         "SPK.SGsm, SPK.SCdma, InKategory.kodeKategory as KodeType, InKategory.NamaKategory as NamaType, " & _
'         "BPK.BPK_No as NoDo, Cast(BPK.Tgl_BPK as Char(12)) as TglDO, Cast(T_MRS.TglCR1 as Char(12))as TglCr1, " & _
'         "CAST(T_MRS.KmCR1 as Char(12)) as KmCR1 , Cast(T_MRS.TglCR2 as Char(12)) as TglCR2, Cast(T_MRS.KmCR2 as Char(12)) as KmCR2, " & _
'         "('Sales : ' + AFI.KetSalesCr2 + ' - CRC : ' + AFI.KetCrcCr2)as StatusCR2, " & _
'         "Cast(T_Transfer.TglTransfer as Char(12)) As Tanggal_Transfer " & _
'         "FROM BPK INNER JOIN " & _
'         "AFI On BPK.SPK_No = AFI.SPK_No INNER Join " & _
'         "T_MRS On AFI.SPK_No = T_MRS.NoSPK INNER Join " & _
'         "T_Transfer ON AFI.SPK_No = T_Transfer.NoSPK INNER JOIN " & _
'         "SPK ON AFI.SPK_No = SPK.SPK_No INNER JOIN " & _
'         "InKategory ON SPK.Kode_Type = InKategory.KodeKategory " & _
'         "Where BPK.Batal = 0 And Afi.User_Proses <> 'Null' " & _
'         "AND T_Transfer.TglTransfer Between '" & FormatTgl(tgl1) & "' AND '" & FormatTgl(tgl2) & "' " & _
'         "Order By T_Transfer.TglTransfer ASC"
    strSql = "Select   AFI.SPK_No, SPK.NoRangka, SPK.SNama, SPK.SAlamat, SPK.SKodePos, " & _
         "SPK.SKota, SPK.SPropinsi, SPK.STelp, " & _
         "SPK.SGsm, SPK.SCdma, InKategory.kodeKategory as KodeType, InKategory.NamaKategory as NamaType, " & _
         "BPK.BPK_No as NoDo, BPK.Tgl_BPK  as TglDO, T_MRS.TglCR1 as TglCr1, " & _
         "CAST(T_MRS.KmCR1 as Char(12)) as KmCR1 , T_MRS.TglCR2 as TglCR2, Cast(T_MRS.KmCR2 as Char(12)) as KmCR2, " & _
         "('Sales : ' + AFI.KetSalesCr2 + ' - CRC : ' + AFI.KetCrcCr2)as StatusCR2, " & _
         "T_TransferCR1.TglTransfer As Tanggal_Transfer, tgl1, namaind " & _
         "FROM BPK WITH(NOLOCK) INNER JOIN " & _
         "AFI WITH(NOLOCK) On BPK.SPK_No = AFI.SPK_No  INNER Join " & _
         "T_MRS WITH(NOLOCK) On AFI.SPK_No = T_MRS.NoSPK INNER Join " & _
         "T_TransferCR1 WITH(NOLOCK) ON AFI.SPK_No = T_TransferCR1.NoSPK INNER JOIN " & _
         "SPK WITH(NOLOCK) ON AFI.SPK_No = SPK.SPK_No INNER JOIN " & _
         "InKategory WITH(NOLOCK) ON SPK.Kode_Type = InKategory.KodeKategory " & _
         "INNER JOIN InMst WITH(NOLOCK) ON InMst.KodeItem = SPK.NoRangka " & _
         "INNER JOIN InWarnaExt WITH(NOLOCK) ON InWarnaExt.KodeWarna = InMst.KodeWarna " & _
         "Where BPK.Batal = 0 And Afi.User_Proses <> 'Null' " & _
         "AND T_TransferCR1.TglTransfer Between '" & FormatTgl(Tgl1) & "' AND '" & FormatTgl(Tgl2) & "' " & _
         "Order By T_TransferCR1.TglTransfer DESC"
    End If
    cm.CommandText = strSql
    cm.CommandType = adCmdText

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set cm.ActiveConnection = mConn
    rs.Open cm, , adOpenDynamic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set GetAllTransferRange = rs
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Function

'==========================================================================================
'======================= INVALId CR =======================================================
'==========================================================================================


Public Function GetNotValidCR(param As String) As ADODB.Recordset
        Dim cm As New Command
On Error GoTo errHandler

    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
    strSql = "SELECT T_MRS.KodeLGN, T_MRS.NoSPK, T_MRS.NoRangka, T_MRS.TglCr1, T_MRS.KmCr1, " & _
             "T_MRS.TglCr2, isNull(T_MRS.KmCr2,0) as KmCr2, T_MRS.Tgl1 as Tgl_1000,T_MRS.Km1,T_MRS.Ket1, AFI.NoPol, SPK.Nama_Faktur, " & _
             "(AFI.Alamat1 + AFI.Alamat2 + AFI.Alamat3 + AFI.Alamat4 + AFI.Alamat5) As AlamatFaktur, " & _
             "SPK.Kota , SPK.Kode_Pos, T_Transfer.TglTransfer " & _
             "FROM T_MRS WITH(NOLOCK) INNER JOIN SPK WITH(NOLOCK) ON T_MRS.NOSPK = SPK.SPK_No " & _
             "INNER JOIN T_Transfer WITH(NOLOCK) On T_MRS.noSPK = T_Transfer.NoSPK " & _
             "INNER JOIN AFI WITH(NOLOCK) ON SPK.SPK_No = AFI.SPK_No " & _
             "Where " & param & " " & _
             "Order By T_MRS.Tgl1 desc"
    
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set cm.ActiveConnection = mConn
    rs.Open cm, , adOpenDynamic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set GetNotValidCR = rs
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchSearchTransfer(SearchBy As String, isTgl As Boolean, fill1 As String, fill2 As String, param As String) As ADODB.Recordset
        Dim cm As New Command
On Error GoTo errHandler

    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    strSql = "SELECT T_MRS.KodeLGN, T_MRS.NoSPK, T_MRS.NoRangka, T_MRS.TglCr1, T_MRS.KmCr1, T_MRS.TglCr2, isNull(T_MRS.KmCr2,0) as KmCr2, T_MRS.Tgl1 as Tgl_1000,T_MRS.Km1,T_MRS.Ket1, AFI.NoPol, SPK.Nama_Faktur, " & _
             "(AFI.Alamat1 + AFI.Alamat2 + AFI.Alamat3 + AFI.Alamat4 + AFI.Alamat5) As AlamatFaktur, " & _
             "SPK.Kota , SPK.Kode_Pos, T_Transfer.TglTransfer " & _
             "FROM T_MRS WITH(NOLOCK) INNER JOIN SPK WITH(NOLOCK) ON T_MRS.NOSPK = SPK.SPK_No " & _
             "INNER JOIN T_Transfer WITH(NOLOCK) On T_MRS.noSPK = T_Transfer.NoSPK " & _
             "INNER JOIN AFI WITH(NOLOCK) ON SPK.SPK_No = AFI.SPK_No " & _
             "Where (" & param & ") "

    strSql = strSql & " And " & SearchBy & " Between '" & Format(TglNul(fill1), "MM/DD/YYYY") & "' and '" & Format(TglNul(fill2), "MM/DD/YYYY") & "' Order By T_MRS.Tgl1 desc"
              
        cm.CommandText = strSql
        cm.CommandType = adCmdText

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set cm.ActiveConnection = mConn
    rs.Open cm, , adOpenDynamic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set FetchSearchTransfer = rs
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchItemCustomer(noRangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mSConn.State = adStateClosed Then mSConn.Open (SQLConSPK)
    
    strSql = "SELECT T_MRS.KodeLGN, spk.sNama, spk.sAlamat, Spk.sKodePos, " & _
             "spk.sKota, spk.sPropinsi, " & _
             "SPK.STelp , SPK.sGsm, SPK.sCdma, AFI.noPol,AFI.SPK_No " & _
             "FROM T_MRS WITH(NOLOCK) INNER JOIN " & _
             "SPK WITH(NOLOCK) ON T_MRS.NoSpk = spk.SPK_No INNER JOIN " & _
             "AFI WITH(NOLOCK) ON SPK.Spk_No = AFI.SPK_No " & _
             "WHERE SPK.NoRangka = '" & noRangka & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchItemCustomer = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub UpdateItemCustomer(noRangka As String, n As Integer)
    Dim cm As New Command
On Error GoTo errHandler
    If mSConn.State = adStateClosed Then mSConn.Open (SQLConSPK)
    
    If n = 0 Then
    strSql = "Update SPK set " & _
             "spk.sNama='" & mNama & "',  spk.sAlamat='" & mAlamat & "',  spk.sKodePos='" & mKodePos & "',  spk.sKota='" & mKota & "', " & _
             "spk.sPropinsi='" & mPropinsi & "',  spk.STelp='" & mRumah & "',  spk.sGsm='" & mGSM & "',  spk.sCdma='" & mCDMA & "' " & _
             "WHERE SPK.NoRangka ='" & noRangka & "'"
    Else
    strSql = "Update SPK set " & _
             "spk.sNama='" & mNama & "',  spk.sAlamat='" & mAlamat & "',  spk.sKodePos='" & mKodePos & "',  spk.sKota='" & mKota & "', " & _
             "spk.STelp='" & mRumah & "',  spk.sGsm='" & mGSM & "',  spk.sCdma='" & mCDMA & "' " & _
             "WHERE SPK.NoRangka ='" & noRangka & "'"
    End If
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    
    Set cm.ActiveConnection = mSConn
    cm.Execute
    
    Set cm = Nothing
    mSConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub
Public Sub UpdateItemCustomerNoPol(NoSPK As String, nopol As String)
    Dim cm As New Command
On Error GoTo errHandler
    If mSConn.State = adStateClosed Then mSConn.Open (SQLConSPK)
    
    strSql = "Update AFI set " & _
             "NoPol = '" & nopol & "' " & _
             "WHERE AFI.SPK_No ='" & NoSPK & "'"
   
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    
    Set cm.ActiveConnection = mSConn
    cm.Execute
    
    Set cm = Nothing
    mSConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function FetchItemCustomerBengkel(noRangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    strSql = "select FS_KD_PLG,UseName, UseAddr, UsePos, UseCity, UseCity, " & _
             "UsePSTN, UseGSM, UseCDMA, No_Polisi FROM CUSTDAT WITH(NOLOCK) " & _
             "Where No_Chassis = '" & noRangka & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchItemCustomerBengkel = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub UpdateItemCustomerBengkel(noRangka As String)
    Dim cm As New Command
On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Update Custdat set " & _
             "UseName='" & mNama & "',  UseAddr='" & mAlamat & "',  UsePos='" & mKodePos & "',  UseCity='" & mKota & "', " & _
             "UsePSTN='" & mRumah & "',  UseGSM='" & mGSM & "',  UseCDMA='" & mCDMA & "', " & _
             "No_Polisi='" & mNoPolisi & "' WHERE No_Chassis ='" & noRangka & "'"
   
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    
    Set cm.ActiveConnection = mConn
    cm.Execute
    
    Set cm = Nothing
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Sub
    
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub AddNew_SCust(Rangka As String, Nama As String, Alamat As String, KodePos As String, Kota As String, Propinsi As String, Rumah As String, GSM As String, CDMA As String, NoPolisi As String)
   On Error GoTo ErrAddLog
      
      Dim cm As New Command
      If mConn.State <> adStateClosed Then mConn.Close
      If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
      strSql = "Insert Into SCust (SRangka, SNama, SAlamat, SKodePos, SKota, SPropinsi, STelp, SGsm, SCdma, UserCs, EntryCs) " & _
               "Values ( '" & Rangka & "' , '" & Nama & "' ,'" & Alamat & "' , '" & KodePos & "' ,'" & Kota & "' , '" & Propinsi & "' ,'" & Rumah & "' , '" & GSM & "' ,'" & CDMA & "' , '" & gSA & "' , GetDate())"
                                  
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddLog:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function FetchSearchListSBI(SearchBy As String, param As String) As ADODB.Recordset
Dim cm As New Command

    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    strSql = "SELECT T_MRS.KodeLGN, T_MRS.NoSPK, T_MRS.NoRangka, TGL_BPK, T_MRS.TglCr1, T_MRS.KmCr1, T_MRS.TglCr2, T_MRS.KmCr2, T_MRS.Tgl1 as Tgl_1000,T_MRS.Km1,T_MRS.Ket1, AFI.NoPol, " & _
             "SNama, SAlamat, SKota, SKodePos, T_Transfer.TglTransfer, STelp, SGsm, SCdma, " & _
             "isNull((Select top 1 TglProses from " & Log.LogBengkel & "..M_listReminder WITH(NOLOCK) Where NoRangka = T_MRS.NoRangka and kmnextsvc = 1000 order by TglProses desc),'') as SMS, " & _
             "isNull((Select top 1 tglEntry from " & Log.LogBengkel & "..M_listReminderCall WITH(NOLOCK) Where NoRangka = T_MRS.NoRangka and kmnextsvc = 1000 order by TglEntry desc),'') as CALL " & _
             "FROM T_MRS WITH(NOLOCK) INNER JOIN SPK WITH(NOLOCK) ON T_MRS.NOSPK = SPK.SPK_No " & _
             "INNER JOIN T_Transfer WITH(NOLOCK) On T_MRS.noSPK = T_Transfer.NoSPK INNER JOIN AFI WITH(NOLOCK) ON SPK.SPK_No = AFI.SPK_No INNER JOIN BPK WITH(NOLOCK) ON " & _
             "SPK.SPK_No = BPK.SPK_No  Where (" & param & ") AND " & SearchBy
              
    cm.CommandText = strSql
    cm.CommandType = adCmdText

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set cm.ActiveConnection = mConn
    rs.Open cm, , adOpenStatic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set FetchSearchListSBI = rs
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchDataLockAR_() As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    If OnTopSV = 0 Then
    strSql = "Select Count(X.KodeLgn) FROM ARMST X WITH(NOLOCK) Inner Join " & _
             "(Select A.* FROM (Select KodeLgn, Sum(JumlahTrn) as AR From ARTRN WITH(NOLOCK) Group By KodeLgn) A " & _
             "Where A.AR > 0) Y On X.KodeLgn = Y.Kodelgn " & _
             "Where y.AR > x.BatasKredit And x.BatasKredit > 0"
    
    'UNTUK FEBUARI
    Else
    strSql = "Select Count(X.KodeLgn) FROM ARMST X WITH(NOLOCK) Inner Join " & _
         "(Select A.* FROM (Select KodeLgn, Sum(JumlahTrn) as AR From ARTRN WITH(NOLOCK) Where left(NoFaktur,2) = 'E-' Group By KodeLgn) A " & _
         "Where A.AR > 0) Y On X.KodeLgn = Y.Kodelgn " & _
         "Where y.AR > x.BatasGRP And x.BatasGRP > 0"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    FetchDataLockAR_ = Val(rs.Fields(0).value)
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
    FetchDataLockAR_ = 0
End Function

Public Function FetchDataCust_OnAR_Overdue_(nRange As Double) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    If OnTopSV = 0 Then
'    strSql = "Select Count(Distinct(Y.KodeLgn)) FROM ( " & _
'             "Select  X.NilaiAR, X.KodeLgn FROM ( " & _
'             "Select Sum(JumlahTRn) as NilaiAR, KodeLgn from ARTRN WITH(NOLOCK) Where KodeLgn in (select Distinct(KodeLgn) as KodeLg from ARTRN WITH(NOLOCK) Where KodeLgn in (Select KodeLgn FROM ARMST WITH(NOLOCK) Where SyaratBayar not in ('D0',''))) " & _
'             "And TglJthTmp < '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "' And NoFaktur Not in ( " & _
'             "Select NoFaktur From ARTRN WITH(NOLOCK) " & _
'             "Where KodeLgn  in (select Distinct(KodeLgn) as KodeLg from ARTRN WITH(NOLOCK) Where KodeLgn in " & _
'             "(Select KodeLgn FROM ARMST WITH(NOLOCK) Where SyaratBayar not in ('D0',''))) and TglJthTmp > '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "') Group By NoFaktur, KodeLgn) X Where x.NilaiAR > 0 " & _
'             ") Y"
   
    strSql = "Select  COUNT(distinct X.KodeLgn)" & _
             " FROM (" & _
             " Select Sum(JumlahTRn) as NilaiAR, a.KodeLgn" & _
             " from ARTRN a WITH(NOLOCK) inner join ARMST b WITH(NOLOCK) on a.KodeLgn = b.KodeLgn" & _
             " Where SyaratBayar not in ('D0','') And TglJthTmp < '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "'" & _
             " Group By NoFaktur, a.KodeLgn) X Where x.NilaiAR > 0"
    Else
    'UNTUK FEBUARI
'    strSql = "Select Count(Distinct(Y.KodeLgn)) FROM ( " & _
'             "Select  X.NilaiAR, X.KodeLgn FROM ( " & _
'             "Select Sum(JumlahTRn) as NilaiAR, KodeLgn from ARTRN WITH(NOLOCK) Where left(NoFaktur,2) = 'E-' AND KodeLgn in (select Distinct(KodeLgn) as KodeLg from ARTRN WITH(NOLOCK) Where KodeLgn in (Select KodeLgn FROM ARMST WITH(NOLOCK) Where TOPGRP not in ('D0',''))) " & _
'             "And TglJthTmp < '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "' And NoFaktur Not in ( " & _
'             "Select NoFaktur From ARTRN WITH(NOLOCK) " & _
'             "Where KodeLgn  in (select Distinct(KodeLgn) as KodeLg from ARTRN WITH(NOLOCK) Where left(NoFaktur,2) = 'E-' AND KodeLgn in " & _
'             "(Select KodeLgn FROM ARMST WITH(NOLOCK) Where TOPGRP not in ('D0',''))) and TglJthTmp > '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "') Group By NoFaktur, KodeLgn) X Where x.NilaiAR > 0 " & _
'             ") Y"
    strSql = "Select  COUNT(distinct X.KodeLgn)" & _
             " FROM (" & _
             " Select Sum(JumlahTRn) as NilaiAR, a.KodeLgn" & _
             " from ARTRN a WITH(NOLOCK) inner join ARMST b WITH(NOLOCK) on a.KodeLgn = b.KodeLgn" & _
             " Where left(NoFaktur,2) = 'E-' AND TOPGRP not in ('D0','') And TglJthTmp < '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "'" & _
             " Group By NoFaktur, a.KodeLgn) X Where x.NilaiAR > 0"
             
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    FetchDataCust_OnAR_Overdue_ = Val(rs.Fields(0).value)
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
    FetchDataCust_OnAR_Overdue_ = 0
End Function

Public Function FetchDataLockAR_DATA() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    If OnTopSV = 0 Then
    strSql = "Select X.KodeLgn, X.NamaLgn as Nama, X.Alamat1 + ' ' + X.Alamat2 as Alamat,X.BatasKredit, Y.AR, (Y.AR/X.BatasKredit) as Persen  FROM ARMST X WITH(NOLOCK) Inner Join " & _
             "(Select A.* FROM (Select KodeLgn, Sum(JumlahTrn) as AR From ARTRN WITH(NOLOCK) Group By KodeLgn) A " & _
             "Where A.AR > 0) Y On X.KodeLgn = Y.Kodelgn " & _
             "Where y.AR > x.BatasKredit And x.BatasKredit > 0"
    Else
    'UNTUK FEBUARI
    strSql = "Select X.KodeLgn, X.NamaLgn as Nama, X.Alamat1 + ' ' + X.Alamat2 as Alamat,X.BatasGRP AS BatasKredit, Y.AR, (Y.AR/X.BatasGRP) as Persen  FROM ARMST X WITH(NOLOCK) Inner Join " & _
             "(Select A.* FROM (Select KodeLgn, Sum(JumlahTrn) as AR From ARTRN WITH(NOLOCK) Where left(NoFaktur,2) = 'E-' Group By KodeLgn) A " & _
             "Where A.AR > 0) Y On X.KodeLgn = Y.Kodelgn " & _
             "Where y.AR > x.BatasGRP And x.BatasGRP > 0"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataLockAR_DATA = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
    'err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataCust_OnAR_Overdue_DATA(nRange As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    If OnTopSV = 0 Then
'    strSql = "Select C.KodeLgn, C.NamaLgn as Nama, C.Alamat1 + ' ' + C.Alamat2 as Alamat, C.SyaratBayar FROm ARMST C WITH(NOLOCK) Where " & _
'             "C.KodeLgn in (Select Distinct(Y.KodeLgn) as KodeLgn FROM ( " & _
'             "Select  X.NilaiAR, X.KodeLgn FROM ( " & _
'             "Select Sum(JumlahTRn) as NilaiAR, KodeLgn from ARTRN WITH(NOLOCK) Where KodeLgn in (select Distinct(KodeLgn) as KodeLg from ARTRN WITH(NOLOCK) Where KodeLgn in (Select KodeLgn FROM ARMST WITH(NOLOCK) Where SyaratBayar not in ('D0',''))) " & _
'             "And TglJthTmp < '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "' And NoFaktur Not in ( " & _
'             "Select NoFaktur From ARTRN WITH(NOLOCK) " & _
'             "Where KodeLgn  in (select Distinct(KodeLgn) as KodeLg from ARTRN WITH(NOLOCK) Where KodeLgn in " & _
'             "(Select KodeLgn FROM ARMST WITH(NOLOCK) Where SyaratBayar not in ('D0',''))) and TglJthTmp > '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "') Group By NoFaktur, KodeLgn) X Where x.NilaiAR > 0 " & _
'             ") Y)"
'
    strSql = "Select distinct b.KodeLgn, b.NamaLgn as Nama, b.Alamat1 + ' ' + b.Alamat2 as Alamat, b.SyaratBayar " & _
             " from ARTRN a WITH(NOLOCK) inner join ARMST b WITH(NOLOCK) on a.KodeLgn = b.KodeLgn " & _
             " Where SyaratBayar not in ('D0','') And TglJthTmp < '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "' " & _
             " Group By NoFaktur, b.KodeLgn, b.NamaLgn, b.Alamat1, b.Alamat2 , b.SyaratBayar " & _
             " Having Sum(JumlahTRn) > 0 "

    Else
    'UNTUK FEBUARI - TOP PELANGGAN
'    strSql = "Select C.KodeLgn, C.NamaLgn as Nama, C.Alamat1 + ' ' + C.Alamat2 as Alamat, C.TOPGRP as SyaratBayar FROm ARMST C WITH(NOLOCK) Where " & _
'             "C.KodeLgn in (Select Distinct(Y.KodeLgn) as KodeLgn FROM ( " & _
'             "Select  X.NilaiAR, X.KodeLgn FROM ( " & _
'             "Select Sum(JumlahTRn) as NilaiAR, KodeLgn from ARTRN WITH(NOLOCK) Where left(NoFaktur,2) = 'E-' AND KodeLgn in (select Distinct(KodeLgn) as KodeLg from ARTRN WITH(NOLOCK) Where left(NoFaktur,2) = 'E-' AND KodeLgn in (Select KodeLgn FROM ARMST WITH(NOLOCK) Where TOPGRP not in ('D0',''))) " & _
'             "And TglJthTmp < '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "' And NoFaktur Not in ( " & _
'             "Select NoFaktur From ARTRN WITH(NOLOCK) " & _
'             "Where left(NoFaktur,2) = 'E-' AND KodeLgn  in (select Distinct(KodeLgn) as KodeLg from ARTRN WITH(NOLOCK) Where left(NoFaktur,2) = 'E-' AND KodeLgn in " & _
'             "(Select KodeLgn FROM ARMST WITH(NOLOCK) Where TOPGRP not in ('D0',''))) and TglJthTmp > '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "') Group By NoFaktur, KodeLgn) X Where x.NilaiAR > 0 " & _
'             ") Y)"
     strSql = "Select distinct b.KodeLgn, b.NamaLgn as Nama, b.Alamat1 + ' ' + b.Alamat2 as Alamat, b.SyaratBayar " & _
             " from ARTRN a WITH(NOLOCK) inner join ARMST b WITH(NOLOCK) on a.KodeLgn = b.KodeLgn " & _
             " Where left(NoFaktur,2) = 'E-' AND TOPGRP not in ('D0','') And TglJthTmp < '" & Format(DateAdd("D", nRange, isnow), "mm/dd/YYYY") & "' " & _
             " Group By NoFaktur, b.KodeLgn, b.NamaLgn, b.Alamat1, b.Alamat2 , b.SyaratBayar " & _
             " Having Sum(JumlahTRn) > 0 "
             
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataCust_OnAR_Overdue_DATA = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
    'err.Raise err.Number, err.Source, err.Description
End Function



