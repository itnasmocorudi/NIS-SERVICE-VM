Attribute VB_Name = "MainModul"
Option Explicit

Private Declare Function GetVolumeInformation Lib _
   "kernel32.dll" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, _
   ByVal lpVolumeNameBuffer As String, _
   ByVal nVolumeNameSize As Integer, _
   lpVolumeSerialNumber As Long, _
   lpMaximumComponentLength As Long, _
   lpFileSystemFlags As Long, _
   ByVal lpFileSystemNameBuffer As String, _
   ByVal nFileSystemNameSize As Long) As Long

Global gManPart As Boolean
Global gTglManAw As Date
Global gTglManAk As Date

Global gSudahInv As Boolean
Global gGudang As String
Global gGenerasi As String
Global gFiskal As String
Global gKodeRetur As String
Global gBlnBedaPart As String

Global gTglPart As String
Global gUkuranOrder As Boolean
Global gUkuranInvoice As Boolean
Global gUkuranRetur As Boolean
Global gPreviewCetak As Boolean
Global gIERC As Boolean

Global gBank As String
Global gAcc As String
Global gRefreshInfo As Integer

Global gInvNo As String
Global gInvTgl As String
Global gInvCust As String
Global gInvAlmt As String
Global gInvKota As String
Global gInvNPWP As String
Global gInvNoPolisi As String
Global gInvModel As String
Global gInvNoMesin As String
Global gInvNoChas As String
Global gInvNoOrder As String
Global gInvSA As String
Global gInvStatus As String
Global gInvMech As String
Global gInvPrint As String
Global gInvBilling As String
Global gInvUraian1 As String
Global gInvUraian2 As String
Global gInvUraian3 As String
Global gInvUraian4 As String
Global gInvUraian5 As String
Global gInvUraian6 As String
Global gInvUraian7 As String
Global gInvRate1 As String
Global gInvRate2 As String
Global gInvRate3 As String
Global gInvRate4 As String
Global gInvRate5 As String
Global gInvRate6 As String
Global gInvRate7 As String
Global gInvRp1 As String
Global gInvRp2 As String
Global gInvRp3 As String
Global gInvRp4 As String
Global gInvRp5 As String
Global gInvRp6 As String
Global gInvRp7 As String
Global gGInvRp As String
Global gInvTRp1 As String
Global gInvTRp2 As String
Global gInvTRp3 As String
Global gInvTRp4 As String
Global gInvTRp5 As String
Global gInvTRp6 As String
Global gInvTRp7 As String
Global gGInvTRp As String
Global gInvTerbilang As String
Global gInvSS As String
Global gInvCatatan As String
Global gInvSM As String

Global gNoOrder As String
Global gNoInvoice As String
Global gKetSA As String
Global gCust As String
Global gAlmt As String
Global gKota As String
Global gTelpon As String
Global gCatatan As String
Global gNoPolisi As String
Global gModel As String
Global gWarna As String
Global gTglMasuk As String
Global gJamMasuk As String
Global gDealer As String
Global gDD As String
Global gMDLOrder As String
Global gFRM As String
Global gEG As String
Global gKeluhan(5) As String
Global gOrder(7) As String
Global gPartKd(7) As String
Global gPart(7) As String
Global gPartQty(7) As String
Global gKM As String
Global gPKhusus As String
Global gSAdvisor As String
Global gPSelesai As String
Global gTahun As String
Global gKodePlg As String
Global tPesan As String
Global gKdCbg As String
Global gNmCbg As String

Global gVersi As String
Global gCopyright As String
Global gSecmanager As String
Global gNpwp As String
Global gPkp As String
Global gTelp As String
Global gNoPolis As String
Global gGroupModel As String
Global gServerDb As String
Global gDb As String

Global LogonStr As String
Global LogonSrNis As String

Global oConn As New ADODB.Connection
Global oConnAcces As New ADODB.Connection
Global oConnAcces2 As New ADODB.Connection
Global oConnNis As New ADODB.Connection
Global KonDbs As String
Global KonDbNis As String
Global KonDbAcces As String
Global KonDbAcces2 As String
Global adoCon As String
Global adoConAcces As String
Global adoConAcces2 As String
Global adoConNis As String

Global gsSql1(90) As String
Global gsSql2(90) As String
Global oRs As New ADODB.Recordset
Global oRs1 As New ADODB.Recordset
Global oRs2 As New ADODB.Recordset
Global oRs3 As New ADODB.Recordset
Global oRsNis As New ADODB.Recordset
Global sSQL As String
Global gPrinter As String
Global gMainKeyCari As String
Global gBantuan As String
Global gBantuan2 As String
Global gStall As String
Global gSendTab As Boolean
Global gSearchOption As String

Global gTglDefault As String
Global gMDL As String
Global gRateCOM As Double
Global gRatePAS As Double
Global gRateIRC As Double
Global gTglInv As String

Global gTanggal As String
Global gTglServer As Date
Global gSA As String
Global gDiv As String
Global gNama_SA As String
Global gKd_Sama As String

Global gPerush As String
Global gAlamat As String
Global gLoket As String
Global gNoAntri As Integer
Global gTpAntri As String
Global gCabang As String
Global gTglTemp As String
Global gTglTemp1 As String
Global gTglTemp2 As String
Global gKetTemp As String
Global gReportConn As String
Global gUser As String
Global gPwd As String
Global Encry As New acak.Class
Global gReturn As Integer
Global gKode As String
Global gKodeFaktur As String
Global gConnecAAS As Boolean
Global gTag As String

'variabel antrian
Public antrian As String
Public sisaBook As Integer
Public sisaNonBook As Integer

'Tabel Stall
Public MyArray As New XArrayDB

'.:: ======================================== .:: BEGIN ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: M R S ::. ========================================= ::.
Private Const PathMrs32 As String = "C:\WINDOWS\system32\Mrs32.Ant"
Public SQLCon        As String
Public SQLConSPK     As String

Type LoginServer
     LogServ         As String
     LogUser         As String
     LogPass         As String
     LogNas          As String
     LogAlamat       As String
     LogBengkel      As String
     LogSPK          As String
End Type
Public Log          As LoginServer

Public Const AccessCon As String = "Provider = Microsoft.JET.OLEDB.4.0;" & _
"Data Source =  C:\Inetpub\wwwroot\sms\dbSMS.mdb"

Public IsNow As Date
'.:: ======================================== .:: M R S ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: END ::. ========================================= ::.


Public Function Koma2Titik(Nilai As Variant) As Double
    On Error GoTo Keluar
    Dim temp As String
    Dim i As Integer
    
    temp = str(Nilai)
    For i = 1 To Len(temp)
        If Mid(temp, i, 1) = "," Then
            temp = Left(temp, i - 1) + "." + Right(temp, Len(temp) - i)
        End If
    Next
    Koma2Titik = Val(temp)
    Exit Function
Keluar:
    Koma2Titik = 0
End Function
Public Function NmFieldRetur() As String
    NmFieldRetur = "CUSTOMER, CUST_ADRES, KOTA, TELPON, NO_POLISI, NO_ENGINE, NO_CHASSIS, MDL, MODEL, FMDL, COLOR, FS_KD_DEALER, FS_NM_DEALER," _
        & " DDATE, CODE, NPWP, PPN, DISC, FRONT, [GROUP], MECH, KM, [DATE], TIME, TIME_M, TIME_S, TIME_A, PART * -1, MTRL * -1, OIL * -1, MET * -1, SBLT * -1," _
        & " LABOR * -1, BODY * -1, TOTAL * -1, JOB, REQ1, REQ2, REQ3, REQ4, REQ5, REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, SLIP, RTE1 * -1, RTE2 * -1," _
        & " RTE3 * -1, RTE4 * -1, RTE5 * -1, RTE6 * -1, RTE7 * -1, RTE8 * -1, COND1, COND2, COND3, COND4, COND5, COND6, COND7, convert(char(10),FD_RETUR,120), NO_WO, FS_NO_RETUR, STATUS, PAID," _
        & " NOFAK, REMARK1, SUGGEST, DELIV, SPK, RPR_COND, MEKANIK1, MEKANIK2, MEKANIK3, MEKANIK4, MEKANIK5, MEKANIK6, MEKANIK7," _
        & " MEKANIK8, GRP1, GRP2, GRP3, GRP4, GRP5, GRP6, GRP7, GRP8, ACTRATE1 * -1, ACTRATE2 * -1, ACTRATE3 * -1, ACTRATE4 * -1, ACTRATE5 * -1, ACTRATE6 * -1," _
        & " ACTRATE7 * -1, ACTRATE8 * -1, OFF1 * -1, OFF2 * -1, OFF3 * -1, OFF4 * -1, OFF5 * -1, OFF6 * -1, OFF7 * -1, OFF8 * -1, KLPKERJA1, KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5," _
        & " KLPKERJA6, KLPKERJA7, KLPKERJA8, FS_KREDIT, FN_PAJAK * -1, FS_TH_MOBIL, FN_RETURN, FS_CATATAN, FS_LOG_ADD, FS_LOG_DEL, FN_JML_ORDER * -1," _
        & " FN_JML_RATE * -1, ALASAN_INV, FS_RC, FS_KD_PLG, FS_KD_WIL, FD_TGL_PART, FS_NO_RETUR, FD_RETUR, FS_KD_RETUR"
End Function

Public Sub BuatTabelRepo(vNamaTabel As String)
    sSQL = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
         & " drop table [dbo].[" & vNamaTabel & "]"
    oConn.Execute sSQL
    
    sSQL = "CREATE TABLE [dbo].[" & vNamaTabel & "] ( " _
         & "[NO_CHASSIS] [varchar] (25) NOT NULL DEFAULT '', " _
         & "[NO_POLISI][varchar](11) NOT NULL DEFAULT '', " _
         & "[FS_KD_PLG][varchar](13) NOT NULL DEFAULT '', " _
         & "[FS_KD_WIL][char](3) NOT NULL DEFAULT '', " _
         & "[CUSTOMER][varchar](30) NOT NULL DEFAULT '', " _
         & "[MDL][varchar](15) NOT NULL DEFAULT '', " _
         & "[FMDL][varchar](25) NOT NULL DEFAULT '', " _
         & "[TGLSVC][datetime] NOT NULL DEFAULT '01/01/3000', " _
         & "[NO_WO][char](8) NOT NULL DEFAULT '', " _
         & "[SA][char](5) NOT NULL DEFAULT '', "
          
    sSQL = sSQL + "[KLPKERJA][char](5) NOT NULL DEFAULT '', " _
         & "[REPS][varchar](50) NOT NULL DEFAULT '', " _
         & "[MECHS][char](5) NOT NULL DEFAULT '', " _
         & "[ACTRATE][DECIMAL](18,2) NOT NULL DEFAULT '0', " _
         & "[DISC][numeric](9) NOT NULL DEFAULT '0', " _
         & "[OFFS][DECIMAL](18,2) NOT NULL DEFAULT '0', " _
         & "[TPSVC][char](1) NOT NULL DEFAULT '', " _
         & "[THS][char](1) NOT NULL DEFAULT '', " _
         & "[FN_RETURN][char](1) NOT NULL DEFAULT '', " _
         & "[PROSES][char](1) NOT NULL DEFAULT '', " _
         & "[NOFAK][char](8) NOT NULL DEFAULT '', " _
         & "[NOSTALL][varchar](10) NOT NULL DEFAULT '', " _
         & "[FS_KD_DEALER][char](5) NOT NULL DEFAULT '', " _
         & "[FS_NM_DEALER][varchar](30) NOT NULL DEFAULT '') ON [PRIMARY]"
    oConn.Execute sSQL
End Sub

Public Sub BuatTabelInv(vNamaTabel As String)
    sSQL = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
         & " drop table [dbo].[" & vNamaTabel & "]"
    oConn.Execute sSQL
    
    sSQL = "CREATE TABLE [dbo].[" & vNamaTabel & "] (" _
         & " [CUSTOMER] [varchar] (30) NOT NULL ," _
         & " [CUST_ADRES] [varchar] (40) NOT NULL ," _
         & " [KOTA] [varchar] (30) NOT NULL ," _
         & " [TELPON] [varchar] (30) NOT NULL ," _
         & " [NO_POLISI] [varchar] (12) NOT NULL ," _
         & " [NO_ENGINE] [varchar] (15) NOT NULL ," _
         & " [NO_CHASSIS] [varchar] (25) NOT NULL ," _
         & " [MDL] [varchar] (15) NOT NULL ," _
         & " [MODEL] [varchar] (4) NOT NULL ," _
         & " [FMDL] [varchar] (25) NOT NULL ," _
         & " [COLOR] [varchar] (20) NOT NULL ," _
         & " [FS_KD_DEALER] [varchar] (15) NOT NULL ," _
         & " [FS_NM_DEALER] [varchar] (30) NOT NULL ," _
         & " [DDATE] [varchar] (10) NOT NULL ," _
         & " [CODE] [varchar] (9) NOT NULL ," _
         & " [NPWP] [varchar] (15) NOT NULL ," _
         & " [PPN] [decimal](18, 2) NOT NULL ," _
         & " [DISC] [decimal](18, 2) NOT NULL ,"
 
    sSQL = sSQL + " [FRONT] [varchar] (4) NOT NULL ," _
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
    
    sSQL = sSQL + " [REP1] [varchar] (25) NOT NULL ," _
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
    sSQL = sSQL + " [COND7] [varchar] (8) NOT NULL ," _
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
     
    sSQL = sSQL + " [GRP2] [varchar] (4) NOT NULL ," _
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
       
    sSQL = sSQL + " [OFF8] [decimal](18, 2) NOT NULL ," _
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
         
    sSQL = sSQL + " [ALASAN_INV] [VarChar](150) NOT NULL, " _
         & " [FS_RC] [VarChar](4) NOT NULL, " _
         & " [FS_KD_PLG] [varchar] (13) NOT NULL , " _
         & " [FS_KD_WIL] [varchar] (3) NOT NULL , " _
         & " [FD_TGL_PART] [datetime] NOT NULL , " _
         & " [FS_NO_RETUR] [varchar] (11) NOT NULL , " _
         & " [FD_RETUR] [datetime] NOT NULL , " _
         & " [FS_KD_RETUR] [VarChar](8) NOT NULL " _
         & " ) ON [PRIMARY]"
    oConn.Execute sSQL
End Sub

Public Function SimpanKoma2Titik(Nilai As Variant) As String
    On Error GoTo Keluar
    Dim temp As String
    Dim i As Integer
    
    If Nilai <> "" Then
        temp = str(Nilai)
        For i = 1 To Len(temp)
            If Mid(temp, i, 1) = "," Then
                temp = Left(temp, i - 1) + "." + Right(temp, Len(temp) - i)
            End If
        Next
        SimpanKoma2Titik = temp
    Else
        SimpanKoma2Titik = "0"
    End If
    Exit Function
Keluar:
    SimpanKoma2Titik = "0"
End Function

Public Function UbahKoma2Titik(Nilai As String) As String
    On Error GoTo Keluar
    Dim temp As String
    Dim i As Integer
    
    If Nilai <> "" Then
        temp = Nilai
        For i = 1 To Len(temp)
            If Mid(temp, i, 1) = "," Then
                temp = Left(temp, i - 1) + "." + Right(temp, Len(temp) - i)
            End If
        Next
        UbahKoma2Titik = temp
    Else
        UbahKoma2Titik = "0"
    End If
    Exit Function
Keluar:
    UbahKoma2Titik = ""
End Function

Public Sub CekTombol(Akey As Integer)
    Select Case Akey
    Case vbKeyF1
        SendKeys ("0")
    Case vbKeyF2
        SendKeys ("00")
    Case vbKeyF3
        SendKeys ("000")
    Case vbKeyF4
        SendKeys ("0000")
    Case 38
        SendKeys ("+{TAB}")
    Case 40, 13
        SendKeys ("{TAB}")
    End Select
End Sub

Public Sub CekTombolCombo(Akey As Integer)
    Select Case Akey
    Case vbKeyF1
        SendKeys ("0")
    Case vbKeyF2
        SendKeys ("00")
    Case vbKeyF3
        SendKeys ("000")
    Case vbKeyF4
        SendKeys ("0000")
    Case 13
        SendKeys ("{TAB}")
    End Select
End Sub

Public Function CekNumerik(Akey As Integer) As Integer
    Select Case Akey
        Case 44, 48 To 57, 8
            CekNumerik = Akey
        Case Else
            CekNumerik = 0
            SendKeys ("")
    End Select
End Function

Public Function Angka2Huruf(vNilai As Double, Jenis As Byte) As String
    Dim xNilai As String: Dim xPjgNilai As Byte: Dim xProses As String
    Dim xHasil As String: Dim xHasil1 As String: Dim i As Integer: Dim PerTiga As Byte
    Dim Koma As Byte: Dim KomaPosisi As Byte: Dim xxNilai As String: Dim xHasil2 As String
    Dim Sen As Boolean
    Dim DiKoma As Integer, Ambil As Integer
    Dim j As Integer, k As Integer
    Dim Spasi As Integer
    
    xxNilai = Trim(str(vNilai))
    Koma = 1
    KomaPosisi = 0
    For i = 1 To Len(xxNilai)
        If Mid(xxNilai, i, 1) = "." Then
            Koma = 2
            KomaPosisi = i
        End If
    Next
    
    DiKoma = 0
    Sen = False
    If KomaPosisi <> 0 Then
        If Len(xxNilai) - KomaPosisi <= 2 Then
            If Len(xxNilai) - KomaPosisi <> 0 Then
                DiKoma = 2
                Sen = True
            End If
        Else
            DiKoma = Len(xxNilai) - KomaPosisi
        End If
    End If
    
    Dim Formating As String
    Formating = "#0"
    If DiKoma <> 0 Then
        Formating = Formating + "."
        For i = 1 To DiKoma
            Formating = Formating + "0"
        Next
    End If
    xxNilai = Trim(Format(vNilai, Formating))
    
    
    For j = 1 To Koma
        xHasil = ""
        xHasil1 = ""
        If j = 1 Then
            If KomaPosisi = 0 Then
                xNilai = xxNilai
            Else
                xNilai = Mid(xxNilai, 1, KomaPosisi - 1)
            End If
        Else
            xNilai = Mid(xxNilai, KomaPosisi + 1, Len(xxNilai))
        End If
        
        PerTiga = 0
        If vNilai <> "0" Then
            Do While xNilai <> ""
               xProses = Right(xNilai, 3)
               xNilai = Left(xNilai, Len(xNilai) - Len(xProses))
               If xProses <> "000" Then
                    Select Case PerTiga
                    Case 1
                         xHasil = " ribu" + xHasil
                    Case 2
                         xHasil = " juta" + xHasil
                    Case 3
                         xHasil = " miliar" + xHasil
                    End Select
               End If
               
               For i = Len(xProses) To 1 Step -1
                   If Mid(xProses, i, 1) <> "0" Then
                        Select Case Len(xProses) - i
                        Case 2
                             xHasil = " ratus" + xHasil
                        Case 1
                             If Mid(xProses, i, 1) <> "1" Then
                                xHasil = " puluh" + xHasil
                             Else
                                If Mid(xProses, Len(xProses), 1) = "0" Then
                                    xHasil = " sepuluh" + xHasil
                                Else
                                    Spasi = 20
                                    For k = 2 To 10
                                        If Mid(xHasil, k, 1) = " " Then
                                            Spasi = k
                                            Exit For
                                        End If
                                    Next
                                    If Spasi = 20 Then
                                        xHasil = xHasil + " belas"
                                    Else
                                        xHasil = Left(xHasil, (k - 1)) + " belas" + Right(xHasil, Len(xHasil) - (k - 1))
                                    End If
                                    If Left(xHasil, 5) = " satu" Then
                                        xHasil = " se" + Right(xHasil, Len(xHasil) - 6)
                                    End If
                                End If
                             End If
                        End Select
                   End If
                   
                   Select Case Mid(xProses, i, 1)
                   Case "1"
                        If Len(xProses) - i = 0 Then
                            If PerTiga = 1 And Len(xProses) = 1 Then
                                xHasil = " se" + xHasil
                            Else
                                xHasil = " satu" + xHasil
                            End If
                        Else
                            If Len(xProses) - i <> 1 Then
                                xHasil = " se" + xHasil
                            End If
                        End If
                   Case "2"
                        xHasil = " dua" + xHasil
                   Case "3"
                        xHasil = " tiga" + xHasil
                   Case "4"
                        xHasil = " empat" + xHasil
                   Case "5"
                        xHasil = " lima" + xHasil
                   Case "6"
                        xHasil = " enam" + xHasil
                   Case "7"
                        xHasil = " tujuh" + xHasil
                   Case "8"
                        xHasil = " delapan" + xHasil
                   Case "9"
                        xHasil = " sembilan" + xHasil
                   End Select
               Next
               PerTiga = PerTiga + 1
            Loop
        Else
            xHasil = " Nol"
        End If
        Ambil = 1
        For i = 1 To Len(xHasil) - 3
            If Mid(xHasil, i, 3) = "se " Then
                xHasil1 = xHasil1 + Mid(xHasil, Ambil, i + 2 - Ambil)
                Ambil = i + 3
            End If
        Next
        xHasil1 = xHasil1 + Mid(xHasil, Ambil, Len(xHasil) - (Ambil - 1))
        xHasil1 = UCase(Mid(xHasil1, 2, 1)) + Right(xHasil1, Len(xHasil1) - 2)
        If Jenis = 1 Then
            If j = 1 Then
                xHasil2 = xHasil1 + " Rupiah"
            Else
                xHasil2 = xHasil2 + " " + xHasil1 + " Sen"
            End If
        Else
            If j = 1 Then
                xHasil2 = xHasil1 + " US Dolar"
            Else
                If Sen Then
                    xHasil2 = xHasil2 + " " + xHasil1 + " Sen"
                Else
                    xHasil2 = xHasil2 + " " + xHasil1
                End If
            End If
        End If
    Next
    Angka2Huruf = xHasil2
End Function
Public Function FormatTgl(ddate As Date) As String
    FormatTgl = Format(ddate, "mm/dd/yyyy")
End Function

Public Function TglYYYYMMDD(kata As String) As String
    TglYYYYMMDD = Right(kata, 4) + "-" + Mid(kata, 4, 2) + "-" + Left(kata, 2)
End Function

Public Function TglDDMMYYYY(kata As Date) As String
    TglDDMMYYYY = Format(Day(kata), "00") + "-" + Format(Month(kata), "00") + "-" + Trim(str(Year(kata)))
End Function

Public Sub CenterForm(xForm As Form, xMDI As Boolean)
    If xMDI Then
        xForm.Top = (Screen.Height - xForm.Height) / 2 - (xForm.Height / 24)
        xForm.Left = (Screen.Width - xForm.Width) / 2
    Else
        xForm.Top = (Screen.Height - xForm.Height) / 2 + (xForm.Height / 16)
        xForm.Left = (Screen.Width - xForm.Width) / 2
    End If
End Sub

Public Function CekTgl(kata As String) As Boolean
    Dim Tgl As Date
    On Error GoTo errornih
        CekTgl = True
        If Val(Mid(kata, 4, 2)) > 12 Then
            CekTgl = False
        End If
        If Val(Right(kata, 4)) < 1900 Then
            CekTgl = False
        End If
        Tgl = kata
        On Error GoTo 0
    Exit Function
    
errornih:
    CekTgl = False
    On Error GoTo 0
End Function

Public Function CekJam(kata As String) As Boolean
    Dim i As Integer
    
    CekJam = True
    
    If Trim(kata) <> ">>:>>" Then
        If Val(Left(Trim(kata), 2)) >= 24 Or Val(Left(Trim(kata), 2)) < 0 Then
            CekJam = False
        End If
        If Val(Right(Trim(kata), 2)) >= 60 Or Val(Left(Trim(kata), 2)) < 0 Then
            CekJam = False
        End If
        If Len(Trim(kata)) <> 5 Then
            CekJam = False
        End If
        For i = 1 To Len(kata)
            If Mid(kata, i, 1) = ">" Then
                CekJam = False
            End If
        Next
    End If
End Function

Public Function KurangM1M2(Nilai1 As String, Nilai2 As String) As String
    Dim Menit1 As Double
    Dim Menit2 As Double
    Dim Jam As String
    Dim menit As String
    
    Menit1 = (Val(Left(Trim(Nilai1), 2)) * 60) + Val(Right(Trim(Nilai1), 2))
    Menit2 = (Val(Left(Trim(Nilai2), 2)) * 60) + Val(Right(Trim(Nilai2), 2))
    
    Jam = Format((Menit2 - Menit1) \ 60, "00")
    menit = Format((Menit2 - Menit1) Mod 60, "00")
    
    If Jam >= 24 Or Jam < 0 Then
        KurangM1M2 = "00:00"
    Else
        KurangM1M2 = Trim(Jam) + ":" + Trim(menit)
    End If
End Function

Public Sub ClearsSql()
    Dim i As Byte
    For i = 0 To 90
        gsSql1(i) = ""
        gsSql2(i) = ""
    Next
End Sub

Public Function Replicate(ByVal Huruf As String, Kali As Integer) As String
    Dim j As Integer
    Replicate = ""
    For j = 1 To Kali
        Replicate = Replicate + Huruf
    Next j
End Function

Public Function AddSpace(ByVal xVariabel As String, xPJ As Integer, LCR As Byte, CharSpace As String) As String
    'LCR=left,center,right
    Dim Kiri As Integer, Kanan As Integer, Tambahan As Integer
    If xPJ > Len(xVariabel) Then
        Select Case LCR
        Case 1
            AddSpace = xVariabel & Replicate(CharSpace, xPJ - Len(xVariabel))
        Case 2
            Tambahan = xPJ - Len(xVariabel)
            Kiri = Round(Tambahan / 2)
            Kanan = Tambahan - Kiri
            AddSpace = Replicate(CharSpace, Kiri) + xVariabel + Replicate(CharSpace, Kanan)
        Case 3
            AddSpace = Replicate(CharSpace, xPJ - Len(xVariabel)) & xVariabel
        End Select
    Else
        AddSpace = Left(xVariabel, xPJ)
    End If
End Function

Public Function GetNamaBulan(ByVal blnKe As Byte) As String
    Select Case blnKe
        Case 1
            GetNamaBulan = "JANUARI"
        Case 2
            GetNamaBulan = "FEBRUARI"
        Case 3
            GetNamaBulan = "MARET"
        Case 4
            GetNamaBulan = "APRIL"
        Case 5
            GetNamaBulan = "MEI"
        Case 6
            GetNamaBulan = "JUNI"
        Case 7
            GetNamaBulan = "JULI"
        Case 8
            GetNamaBulan = "AGUSTUS"
        Case 9
            GetNamaBulan = "SEPTEMBER"
        Case 10
            GetNamaBulan = "OKTOBER"
        Case 11
            GetNamaBulan = "NOVEMBER"
        Case 12
            GetNamaBulan = "DESEMBER"
    End Select
End Function

Public Function GetTglAkhir(ByVal xBulan As Integer, ByVal xTahun As Integer, tipe As Byte) As String
    'tipe 1 maka bulan diganti mis 1=januari
    Dim xTanggal As String
    xBulan = xBulan + 1
    If xBulan = 13 Then
        xBulan = 1
        xTahun = xTahun + 1
    End If
    xTanggal = "01-" + Format(xBulan, "00") + "-" + str(xTahun)
    xTanggal = TglDDMMYYYY(CDate(TglYYYYMMDD(xTanggal)) - 1)
    If tipe = 1 Then
        xTanggal = Left(xTanggal, 2) + " " + GetNamaBulan(Val(Mid(xTanggal, 4, 2))) + " " + Right(xTanggal, 4)
    End If
    GetTglAkhir = xTanggal
End Function

Public Function SetFont(Jenis As Byte)
    Print #1, Chr(27) + "@";
    Print #1, Chr(27) + "x0";
    Select Case Jenis
    'conden
    Case 0
        Print #1, Chr(27) + Chr(15);
    '15
    Case 1
        Print #1, Chr(27) + "g";
    '12
    Case 2
        Print #1, Chr(27) + Chr(77);
    '10
    Case 3
        Print #1, Chr(27) + Chr(80);
    'besar
    Case 4
        Print #1, Chr(27) + Chr(87) + Chr(1);
    End Select
End Function

Public Function DriveSerialNumber(ByVal Drive As String) As Long
    'usage: SN = DriveSerialNumber("C:\")
    Dim lAns As Long
    Dim lRet As Long
    Dim sVolumeName As String, sDriveType As String
    Dim sDrive As String

    'Deal with one and two character input values
    sDrive = Drive
    If Len(sDrive) = 1 Then
        sDrive = sDrive & ":\"
    ElseIf Len(sDrive) = 2 And Right(sDrive, 1) = ":" Then
        sDrive = sDrive & "\"
    End If
  
    sVolumeName = String$(255, Chr$(0))
    sDriveType = String$(255, Chr$(0))
    
    lRet = GetVolumeInformation(sDrive, sVolumeName, _
55, lAns, 0, 0, sDriveType, 255)

    DriveSerialNumber = lAns
End Function

Public Sub HapusTabel(vNmTbl As String)
    sSQL = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & vNmTbl & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" _
         & " drop table [dbo].[" & vNmTbl & "]"
    oConn.Execute sSQL
End Sub

Public Function SysAcces() As String
    sSQL = "Select * from T_PROFIL"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        SysAcces = oRs1("FS_SYS_AAS")
    End If
    oRs1.Close
End Function

Public Function LetakAcces() As String
    If oRs1.State Then oRs1.Close
    sSQL = "Select * from T_PROFIL"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        LetakAcces = oRs1("FS_DB_AAS")
    End If
    oRs1.Close
End Function

Public Function NisKomp() As String
    If oRs1.State Then oRs1.Close
    sSQL = "Select * from T_PROFIL"
    oRs1.Open sSQL, oConn, adOpenStatic
    NisKomp = "SERVER"
    If Not oRs1.BOF Then
        If Not IsNull(oRs1("FS_NIS_Komp")) Then
               NisKomp = Trim(oRs1("FS_NIS_Komp"))
        End If
    End If
    oRs1.Close
End Function

Public Function NisUSER() As String
    If oRs1.State Then oRs1.Close
    sSQL = "Select * from T_PROFIL"
    oRs1.Open sSQL, oConn, adOpenStatic
    NisUSER = "SA"
    If Not oRs1.BOF Then
        If Not IsNull(oRs1("FS_NIS_USER")) Then
               NisUSER = Trim(oRs1("FS_NIS_USER"))
        End If
    End If
    oRs1.Close
End Function

Public Function NisPass() As String
    If oRs1.State Then oRs1.Close
    sSQL = "Select * from T_PROFIL"
    oRs1.Open sSQL, oConn, adOpenStatic
    NisPass = ""
    If Not oRs1.BOF Then
        If Not IsNull(oRs1("FS_NIS_Pass")) Then
               NisPass = Trim(oRs1("FS_NIS_Pass"))
        End If
    End If
    oRs1.Close
End Function

Public Function NisDb() As String

    If oRs1.State Then oRs1.Close
    sSQL = "Select * from T_PROFIL"
    oRs1.Open sSQL, oConn, adOpenStatic
    NisDb = ""
    If Not oRs1.BOF Then
        If Not IsNull(oRs1("FS_NIS_Db")) Then
               NisDb = Trim(oRs1("FS_NIS_Db"))
        End If
    End If
    oRs1.Close
End Function

Function McDelIndex(ModulId As String) As String
   Dim acak As Double
   Randomize Timer
   acak = Int(Rnd * 99999999)
   McDelIndex = ModulId & acak
End Function

Sub Main()
Dim A As String
Dim rsDt As New ADODB.Recordset

    On Error GoTo Setup
    
    gServerDb = Encry.BenahiAcakkata(GetSetting(appname:="BENGKEL", Section:="Server", Key:="Top", Default:="(local)"))
    gUser = Encry.BenahiAcakkata(GetSetting(appname:="BENGKEL", Section:="User", Key:="Top", Default:="sa"))
    gPwd = Encry.BenahiAcakkata(GetSetting(appname:="BENGKEL", Section:="Pwd", Key:="Top", Default:=""))
    gDb = "BENGKEL"
    
    If Encry.BenahiAcakkata(GetSetting(appname:="BENGKEL", Section:="Regist", Key:="Top", Default:="")) <> _
       Encry.AcakNo(Encry.AcakNo(DriveSerialNumber("C:\"))) Then
        MsgBox "Belum Registrasi !"
        'MEMORY_15/4/2009'frmSetup.Show
        Exit Sub
    End If
    
    'gVersi = "Version 1.3.8"
    'gVersi = "Version 2.0.1"
    gVersi = "Version 3.0.1"
    
    'gCopyright = "Copyright (c) 2001 Andalan EDP"
    gCopyright = "Copyright (c) 2009 New Ratna Motors IT"
    'MEMORY_15/4/2009' frmSplash.lblCompany = "PT. New Ratna Motors"
    
    gKode = "SV"
    gKodeRetur = "R"

    'LogonStr = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
    
    If Get_Connection(gDb, gUser, gPwd, gServerDb) Then
        Set MyArray = New XArrayDB
        
        If rsDt.State Then rsDt.Close
        Set rsDt = oConn.Execute("SELECT getDate() as Tgl")
        If Not rsDt.EOF Then
            gTanggal = TglDDMMYYYY(Format(rsDt("TGL"), "DD-MM-YYYY"))
            gTglServer = rsDt("TGL")
            gTglDefault = "3000-01-01"
            'MEMORY_15/4/2009'frmSplash.Show
        End If
    End If


'.:: ======================================== .:: BEGIN ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: M R S ::. ========================================= ::.
   Dim IsText
   Dim Ao As Integer
   Open PathMrs32 For Input As #1
   Do Until EOF(1)
   Line Input #1, IsText
      Ao = Ao + 1
      Select Case Ao
         Case 1: Log.LogServ = IsText
         Case 2: Log.LogUser = IsText
         Case 3: Log.LogPass = IsText
         Case 4: Log.LogNas = IsText
         Case 5: Log.LogAlamat = IsText
         Case 6: Log.LogBengkel = IsText
         Case 7: Log.LogSPK = IsText
      End Select
      If Ao = 7 Then Exit Do
   Loop
   Close #1
   
   Log.LogServ = Crypt(Log.LogServ, "antztein", False)
      Log.LogServ = Crypt(Log.LogServ, "tandayu", False)
   Log.LogUser = Crypt(Log.LogUser, "antztein", False)
      Log.LogUser = Crypt(Log.LogUser, "tandayu", False)
   Log.LogPass = Crypt(Log.LogPass, "antztein", False)
      Log.LogPass = Crypt(Log.LogPass, "tandayu", False)
   Log.LogNas = Crypt(Log.LogNas, "antztein", False)
      Log.LogNas = Crypt(Log.LogNas, "tandayu", False)
   Log.LogAlamat = Crypt(Log.LogAlamat, "antztein", False)
      Log.LogAlamat = Crypt(Log.LogAlamat, "tandayu", False)
   Log.LogBengkel = Crypt(Log.LogBengkel, "antztein", False)
      Log.LogBengkel = Crypt(Log.LogBengkel, "tandayu", False)
   Log.LogSPK = Crypt(Log.LogSPK, "antztein", False)
      Log.LogSPK = Crypt(Log.LogSPK, "tandayu", False)
   SQLCon = "Provider = SQLOLEDB.1; User ID=" & Log.LogUser & "; Password=" & Log.LogPass & "; Initial Catalog=" & Log.LogBengkel & "; Data Source =" & Log.LogServ
   SQLConSPK = "Provider = SQLOLEDB.1; User ID=" & Log.LogUser & "; Password=" & Log.LogPass & "; Initial Catalog=" & Log.LogSPK & "; Data Source =" & Log.LogServ
                                
   Call SetTimerNow
   Dim f As New FormMainMRS
   f.Show
'.:: ======================================== .:: M R S ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: E N D ::. ========================================= ::.
    
   Exit Sub
Setup:
    If Err.Number = -2147467259 Then
        MsgBox "Anda belum Log in, Log off dahulu !", vbCritical + vbOKOnly
        MsgBox Err.Description
    Else
        MsgBox "Server tidak terhubung/Alamat salah !"
        MsgBox Err.Description
    End If
    'MEMORY_15/4/2009'frmSetup.Show
End Sub


'.:: ======================================== .:: BEGIN ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: M R S ::. ========================================= ::.
      Private Sub SetTimerNow()
         Dim mConn As New ADODB.Connection
         Dim rs As New ADODB.Recordset
         mConn.Open (SQLCon)
         rs.CursorLocation = adUseClient
         rs.Open "Select GetDate()", mConn, adOpenStatic, adLockReadOnly, adCmdText
         IsNow = rs.Fields(0).value
      End Sub

      Public Function Crypt(strOriginal As String, strPassword As String, encrypting As Boolean) As String
                              
      Dim valPass As Long, valString As Long
      Dim X As Integer
          For X = 1 To Len(strPassword)
              valPass = valPass + Asc(Mid$(strPassword, X, 1)) + 13
          Next X
          For X = 1 To Len(strOriginal)
              If encrypting = True Then
                  valString = Asc(Mid$(strOriginal, X, 1)) + valPass + (X ^ 2)
                  Do While valString > 255
                      valString = valString - 255
                  Loop
              Else
                  valString = Asc(Mid$(strOriginal, X, 1)) - valPass - (X ^ 2)
                  Do While valString <= 0
                      valString = valString + 255
                  Loop
              End If
              Crypt = Crypt & Chr(valString)
          Next X
      End Function
      
      Public Sub EnterPress(ByVal Key As Integer)
            If Key = 13 Then SendKeys "{tab}"
      End Sub
'
'      Function FormatTgl(ddate As Date) As String
'         FormatTgl = Format(ddate, "mm/dd/yyyy")
'      End Function
      
      Function NullText(value As Variant) As Variant
         NullText = IIf(IsNull(value), "", value)
      End Function
      
      Function NullTgl(Tgl As Variant) As Variant
         If IsNull(Tgl) Or Tgl = "" Then
            NullTgl = "Null"
         Else
            NullTgl = "'" & FormatTgl(CDate(Tgl)) & "'"
         End If
      End Function

      Public Function HargaJualPart(ITEM_PART As String) As Double
         If Get_ConnectionNis(NisDb, NisUSER, NisPass, NisKomp) Then
             Dim rsHarga As New ADODB.Recordset
             Set rsHarga = oConnNis.Execute("SELECT hargajuala FROM inmst WHERE kodeItem = '" & ITEM_PART & "'")
             If Not rsHarga.EOF Then
                 HargaJualPart = CDbl(rsHarga("HARGAJUALA"))
             End If
             rsHarga.Close
         End If
      End Function

'.:: ======================================== .:: M R S ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: E N D ::. ========================================= ::.

Public Function Get_Connection(DBname As String, _
                               UsrName As String, _
                               Pass As String, _
                               SerName As String) As Boolean
1      On Error GoTo Vb_Error
'2      App.Title = "BENGKEL-Nasmoco"
3      Get_Connection = True
4      If oConn.State Then oConn.Close
5      oConn.CursorLocation = adUseClient
6      oConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       KonDbs = oConn.ConnectionString
7      adoCon = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       oConn.ConnectionTimeout = 3600
8      oConn.Open
9      Exit Function

Vb_Error:
10    Get_Connection = False
11    MsgBox ErrMessage(Erl, Err.Number, "Function : Get_Connection"), vbCritical, "Open Connection"
End Function

Public Function Get_ConnectionNis(DBname As String, _
                                  UsrName As String, _
                                  Pass As String, _
                                  SerName As String) As Boolean
1      On Error GoTo Vb_Error
'2      App.Title = "BENGKEL-Nasmoco"
3      Get_ConnectionNis = True
4      If oConnNis.State Then oConnNis.Close
5      oConnNis.CursorLocation = adUseClient
6      oConnNis.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       KonDbNis = oConnNis.ConnectionString
7      adoConNis = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       oConnNis.ConnectionTimeout = 3600
8      oConnNis.Open
9      Exit Function

Vb_Error:
10    Get_ConnectionNis = False
11    MsgBox ErrMessage(Erl, Err.Number, "Function : Get_ConnectionNis"), vbCritical, "Open Connection"
End Function

Public Function Get_ConnectionAcces(DBname As String, _
                                    UsrName As String, _
                                    Pass As String, _
                                    SerName As String) As Boolean
1      On Error GoTo Vb_Error
2      'App.Title = "BENGKEL-Nasmoco"
3      Get_ConnectionAcces = True
4      If oConnAcces.State Then oConnAcces.Close
5      oConnAcces.CursorLocation = adUseClient
6      oConnAcces.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       KonDbAcces = oConnAcces.ConnectionString
7      adoConAcces = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       oConnAcces.ConnectionTimeout = 3600
8      oConnAcces.Open
9      Exit Function

Vb_Error:
10    Get_ConnectionAcces = False
11    MsgBox ErrMessage(Erl, Err.Number, "Function : Get_ConnectionAcces "), vbCritical, "Open Connection"
End Function

Public Function Get_ConnectionAcces2(DBname As String, _
                                     UsrName As String, _
                                     Pass As String, _
                                     SerName As String) As Boolean
1      On Error GoTo Vb_Error
'2      App.Title = "BENGKEL-Nasmoco"
3      Get_ConnectionAcces2 = True
4      If oConnAcces2.State Then oConnAcces2.Close
5      oConnAcces2.CursorLocation = adUseClient
6      oConnAcces2.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       KonDbAcces2 = oConnAcces2.ConnectionString
7      adoConAcces2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       oConnAcces2.ConnectionTimeout = 3600
8      oConnAcces2.Open
9      Exit Function

Vb_Error:
10    Get_ConnectionAcces2 = False
11    MsgBox ErrMessage(Erl, Err.Number, "Function : Get_ConnectionAcces2 "), vbCritical, "Open Connection"
End Function

Public Function ErrMessage(ByVal Errline As Long, _
                           ByVal ErrNumber As Long, _
                           FunctionName As String) As String
    ErrMessage = "Error line Number = " & Errline & vbCrLf & _
                 "Error Number      = " & ErrNumber & vbCrLf & _
                 "Error Description = " & Error$(ErrNumber) & vbCrLf & _
                 "Location Error    = " & FunctionName
End Function

Public Function isNotBlank(txt As String) As Boolean
    If Trim(txt) = "" Then
        isNotBlank = False
    Else
        isNotBlank = True
    End If
End Function

Public Sub Enter(ByVal Key As Integer)
    If Key = 13 Then SendKeys "{tab}"
End Sub

Function Nul(Kode As Variant, Optional Data)
      If IsMissing(Data) Then Data = ""
      Nul = IIf(IsNull(Kode) Or Kode = "", Data, Kode)
End Function

Function NulA(Kode As Variant, Optional Data)
      If IsMissing(Data) Then Data = "0"
      NulA = IIf(IsNull(Kode) Or Kode = "", Data, Kode)
End Function

Function UpCase(Key As Integer) As Integer
  UpCase = Asc(UCase(Chr(Key)))
End Function

Function notBlank(txt As String) As Boolean
    If Trim(txt) = "" Then
        notBlank = False
    Else
        notBlank = True
    End If
End Function

Function isiHari(hr As Integer) As String
    Select Case hr
    Case "2"
        isiHari = "Senin"
    Case "3"
        isiHari = "Selasa"
    Case "4"
        isiHari = "Rabu"
    Case "5"
        isiHari = "Kamis"
    Case "6"
        isiHari = "Jumat"
    Case "7"
        isiHari = "Sabtu"
    Case "1"
        isiHari = "Minggu"
    End Select
End Function

Function PanggilCust(sAwal As String, txt As String, sAkhir As String, antri As Boolean, Optional sLoket As String, Optional lain As String)
    Dim h(50) As String
    Dim A() As String
    Dim i As Integer
    Dim hasil As String
    

    If antri Then
        hasil = Terbilang(Val(txt))
        hasil = Replace(hasil, ";;", ";")
        hasil = IIf(Right(hasil, 1) = ";", Left(hasil, Len(hasil) - 1), hasil)
        A() = Split(hasil, ";")
    Else
        txt = Replace(txt, "-", "")
        For i = 0 To Len(Trim(txt))
            h(i) = Mid(Trim(txt), i + 1, 1)
        Next
    End If
    
    sndPlaySound App.Path & "\sound\" & sAwal, 0
    
    If Trim(lain) = "A" Then
        sndPlaySound App.Path & "\sound\ambil.wav", 0
    ElseIf Trim(lain) = "B" Then
        sndPlaySound App.Path & "\sound\booking.wav", 0
    ElseIf Trim(lain) = "N" Then
        sndPlaySound App.Path & "\sound\nonBook.wav", 0
    End If
    
    If antri Then
        For i = 0 To UBound(A)
           sndPlaySound App.Path & "\sound\" & A(i), 0
           DoEvents
        Next
    Else
        ' Panggil no Polisi dengan suara
        For i = 0 To Len(Trim(txt)) - 1
            sndPlaySound App.Path & "\sound\" & h(i), 0
            DoEvents
        Next i
    End If
    
    sndPlaySound App.Path & "\sound\" & sAkhir, 0
        
    If Trim(sLoket) <> "" Then
        sndPlaySound App.Path & "\sound\" & sLoket, 0
    End If
End Function

Function GensSql2(xTabel As String, xUpdate As Boolean) As String
    Dim aa As String
    Dim xHasil As String
    Dim i As Integer
    
    xHasil = ""
    If xUpdate Then
        xHasil = "UPDATE " + xTabel + " SET "
        i = 1
        Do While gsSql1(i) <> ""
            If i <> 1 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC"
                xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
            Case Else
                xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
            End Select
            i = i + 1
        Loop
    Else
        xHasil = "INSERT INTO " + xTabel + " ("
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            xHasil = xHasil + gsSql1(i)
            i = i + 1
        Loop
        xHasil = xHasil + ") VALUES ("
        
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC"
                xHasil = xHasil + gsSql2(i)
            Case Else
                xHasil = xHasil + "'" + gsSql2(i) + "'"
            End Select
            i = i + 1
        Loop
        xHasil = xHasil + ")"
    End If
    GensSql2 = xHasil
End Function

Function AlfaBulan(bl As Integer) As String
    Select Case bl
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            AlfaBulan = bl
        Case 10
            AlfaBulan = "A"
        Case 11
            AlfaBulan = "B"
        Case 12
            AlfaBulan = "C"
    End Select
End Function

Function HilangVocal(str As String) As String
    Dim n As Integer
    Dim h1 As String
    str = Replace(str, " ", "")
    n = Len(str) - 1
    h1 = Left(UCase(str), 1)
    str = Replace(Right(UCase(str), n), "A", "")
    str = Replace(Right(UCase(str), n), "I", "")
    str = Replace(Right(UCase(str), n), "U", "")
    str = Replace(Right(UCase(str), n), "E", "")
    str = Replace(Right(UCase(str), n), "O", "")
    HilangVocal = h1 & str
End Function

Function AmbilInisial(str As String, digt As Integer) As String
    str = HilangVocal(str)
    AmbilInisial = Left(str, digt)
End Function

Function cekPartPack(tpJob As String, jobOr As String, nMDL As String, idx As Integer, kodeGrid As Integer, Optional barispart As Integer, Optional frm As Object, Optional ByRef RSs As Recordset)
Dim rsHrg As New ADODB.Recordset
Dim hrg
Dim nC As Integer
Dim i As Integer
Dim j As Integer

    If oRs2.State Then oRs2.Close
    Set oRs2 = oConn.Execute("SELECT * FROM ho_partPack WHERE job = '" & tpJob & "' AND FS_ITEm = '" & jobOr & "' AND mdl = '" & nMDL & "'")
    If Not oRs2.EOF Then
        Do While Not oRs2.EOF
            If cekGrid(oRs2, barispart, frm, kodeGrid, RSs) And idx = 1 Then
                If Get_ConnectionNis(NisDb, NisUSER, NisPass, NisKomp) Then
                    If rsHrg.State Then rsHrg.Close
                    Set rsHrg = oConnNis.Execute("SELECT hargajuala FROM inmst WHERE kodeItem = '" & Trim(oRs2("ITEM_PART")) & "'")
                    If Not rsHrg.EOF Then
                        hrg = CDbl(rsHrg("HARGAJUALA"))
                    End If
                    If kodeGrid = 1 Then        'DENGAN FLEXGRID
                        With frm
                            .partGrid.TextMatrix(barispart, 1) = Trim(oRs2("ITEM_PART"))
                            .partGrid.TextMatrix(barispart, 2) = Trim(oRs2("NAMA_PART"))
                            .partGrid.TextMatrix(barispart, 3) = Trim(oRs2("QTY"))
                            .Text(40) = CDbl(.Text(40)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
                            .Text(40) = Format(.Text(40), "#0,0")
                            barispart = barispart + 1
                        End With
                    ElseIf kodeGrid = 2 Then    'DENGAN RECORDSET
                        With RSs
                            .AddNew
                            !kodeItem = Trim(oRs2("ITEM_PART"))
                            !namaItem = Trim(oRs2("NAMA_PART"))
                            !QTY = Trim(oRs2("QTY"))
                            .Update
                            RSs.MoveNext
'                            With frm
'                                .Text(57) = CDbl(.Text(57)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
'                                .Text(57) = Format(.Text(57), "#0,0")
'                            End With
                        End With
                    End If
                End If
                oConnNis.Close
                Set oConnNis = Nothing
            ElseIf idx = 2 Then 'Not cekGrid(oRs2, BarisPart, frm, kodeGrid) And
                If kodeGrid = 1 Then
                    For i = 1 To barispart
                        With frm
                            If .partGrid.TextMatrix(i, 1) = Trim(oRs2("ITEM_PART")) Then
                                If Get_ConnectionNis(NisDb, NisUSER, NisPass, NisKomp) Then
                                    If rsHrg.State Then rsHrg.Close
                                    Set rsHrg = oConnNis.Execute("SELECT hargajuala FROM inmst WHERE kodeItem = '" & Trim(oRs2("ITEM_PART")) & "'")
                                    If Not rsHrg.EOF Then
                                        hrg = CDbl(rsHrg("HARGAJUALA"))
                                    End If
                                    rsHrg.Close
                                End If
                                oConnNis.Close
                                .Text(40) = CDbl(.Text(40)) - (CDbl(.partGrid.TextMatrix(i, 3)) * CDbl(hrg))
                                .Text(40) = Format(.Text(40), "#0,0")
                                
                                nC = i
                                For j = nC To barispart
                                    With .partGrid
                                        .TextMatrix(j, 0) = .TextMatrix(j + 1, 0)
                                        .TextMatrix(j, 1) = .TextMatrix(j + 1, 1)
                                        .TextMatrix(j, 2) = .TextMatrix(j + 1, 2)
                                        .TextMatrix(j, 3) = .TextMatrix(j + 1, 3)
                                    End With
                                Next
                                barispart = barispart - 1
                                
                            End If
                        End With
                    Next
                ElseIf kodeGrid = 2 Then
                    With RSs
                        .Find "KODEITEM = '" & Trim(oRs2("ITEM_PART")) & "'", , adSearchForward, 1
                        If Not .EOF Then
                            If Get_ConnectionNis(NisDb, NisUSER, NisPass, NisKomp) Then
                                If rsHrg.State Then rsHrg.Close
                                Set rsHrg = oConnNis.Execute("SELECT hargajuala FROM inmst WHERE kodeItem = '" & Trim(oRs2("ITEM_PART")) & "'")
                                If Not rsHrg.EOF Then
                                    hrg = CDbl(rsHrg("HARGAJUALA"))
                                End If
                                rsHrg.Close
                            End If
                            oConnNis.Close
                            With frm
                                .Text(57) = CDbl(.Text(57)) - (CDbl(RSs("QTY")) * CDbl(hrg))
                                .Text(57) = Format(.Text(57), "#0,0")
                            End With
                            .Delete
                        End If
                    End With
                End If
            End If
            oRs2.MoveNext
        Loop
    End If
    oRs2.Close
End Function

Function cekGrid(rsT As ADODB.Recordset, barispart As Integer, frmP As Object, kdGrid As Integer, Optional RSs As Recordset) As Boolean
Dim i As Integer

    cekGrid = True
    If kdGrid = 1 Then
        For i = 1 To barispart
            With frmP
                If .partGrid.TextMatrix(i, 1) = Trim(rsT("ITEM_PART")) Then
                    cekGrid = False
                End If
            End With
        Next
    ElseIf kdGrid = 2 Then
        With RSs
            If .RecordCount > 1 Then
                Do While Not .EOF
                    If !kodeItem = Trim(rsT("ITEM_PART")) Then
                        cekGrid = False
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
End Function

Function cekMOU(kMdl As String, nRgk As String) As String

    If oRs1.State Then oRs1.Close
    Set oRs1 = oConn.Execute("SELECT * FROM mouhdr A INNER JOIN moudtl B ON A.no_mou = B.no_mou WHERE mdl = '" & kMdl & "' AND no_Rangka = '" & nRgk & "' ORDER BY mdl, no_Rangka")
    If Not oRs1.EOF Then
        cekMOU = oRs1("NO_MOU")
    End If
    oRs1.Close
End Function

Function cekPiutang(nKdPlg As String, kMdl As String, nRgk As String)
Dim rsMOU As New ADODB.Recordset

    If Get_ConnectionNis(NisDb, NisUSER, NisPass, NisKomp) Then
        If oRs1.State Then oRs1.Close
        Set oRs1 = oConnNis.Execute("SELECT kodelgn, sum(jumlahtrn) as totAR FROM artrn WHERE kodeLgn = '" & nKdPlg & "' GROUP BY kodeLgn ORDER BY kodeLgn")
        If Not oRs1.EOF Then
            If CDbl(oRs1("TOTAR")) > 0 Then
                If rsMOU.State Then rsMOU.Close
                Set rsMOU = oConn.Execute("SELECT isnull(limitar,0) as MaxAr, * FROM mouhdr A INNER JOIN moudtl B ON A.no_mou = B.no_mou WHERE mdl = '" & kMdl & "' AND no_Rangka = '" & nRgk & "' ORDER BY mdl, no_Rangka")
                If Not rsMOU.EOF Then
                    If CDbl(oRs1("TOTAR")) >= CDbl(rsMOU("MAXAR")) Then
                        MsgBox "Masih ada AR sejumlah = " & CDbl(oRs1("TOTAR")) & vbCr & "Mohon untuk dilunasi sebelum hari H service.", vbInformation, "Piutang"
                    End If
                End If
            End If
        End If
    End If
End Function

Function hitungLabour(nRt As Double, klas As String) As Double

    If oRs1.State Then oRs1.Close
    Set oRs1 = oConn.Execute("SELECT * FROM carclass WHERE classcode = '" & Trim(klas) & "'")
    If Not oRs1.EOF Then
        hitungLabour = Val(oRs1("PRICE") * nRt)
    End If
End Function

Function hitungPart(nQty As Integer, nHrg As Double) As Double
    hitungPart = CDbl(nQty) * CDbl(nHrg)
End Function

Function hitungWaktu(nRt As Double, jamMulai As String) As String
Dim jM As Integer   'Waktu Mulai
Dim jS As Integer
Dim mM As Integer   'Waktu Selesai
Dim mS As Integer
Dim Jam As Integer  'Rate
Dim menit As Integer

    jM = jS = mM = mS = 0
    Jam = Round(nRt - 0.5)
    menit = (nRt * 60) Mod 60
    
    'AMBIL WAKTU MULAI
    jM = Left(jamMulai, InStr(jamMulai, ":") - 1)
    mM = Right(jamMulai, InStr(jamMulai, ":") - 1)
    
    jS = jM + Jam + IIf((mM + menit) >= 60, 1, 0)
    mS = mM + menit + IIf((mM + menit) >= 60, (-60), 0)
    
    hitungWaktu = Format(jS, "00") & ":" & Format(mS, "00")
End Function

Function hitMenit(jamMulai As String, jamSelesai As String) As Integer
Dim jM As Integer
Dim mM As Integer
Dim jS As Integer
Dim mS As Integer

    jM = mM = jS = mS = 0
    jM = Left(jamMulai, InStr(jamMulai, ":") - 1) * 60
    mM = Right(jamMulai, InStr(jamMulai, ":") - 1)
    
    jS = Left(jamSelesai, InStr(jamSelesai, ":") - 1) * 60
    mS = Right(jamSelesai, InStr(jamSelesai, ":") - 1)
    
    If jM < jS Then
        hitMenit = (jS + mS) - (jM + mM)
    Else
        hitMenit = 0
    End If
    
End Function

Function cekMRS(nRgk As String, nkm As Integer, fRC As String, tipeJb As Integer, Optional rsJb As ADODB.Recordset, Optional TheGrid As Object) As String
Dim rsHlp As New ADODB.Recordset
Dim rsHlp1 As New ADODB.Recordset
Dim gSBE As Boolean

    If fRC = "" Then
        If tipeJb = 1 Then
            rsJb.MoveFirst
            Do While Not rsJb.EOF
                If rsHlp.State Then rsHlp.Close
                Set rsHlp = oConn.Execute("SELECT SBE FROM jobOrder WHERE fs_item = '" & rsJb("JOB") & "'")
                If Not rsHlp.EOF Then
                    gSBE = IIf(rsHlp("SBE") = "Y", True, False)
                    Exit Do
                Else
                    gSBE = False
                End If
                rsHlp.Close
                rsJb.MoveNext
            Loop
        ElseIf tipeJb = 2 Then
            Dim i As Integer
            For i = 1 To 7
                If rsHlp.State Then rsHlp.Close
                Set rsHlp = oConn.Execute("SELECT SBE FROM jobOrder WHERE fs_item = '" & TheGrid.TextMatrix(i, 1) & "'")
                If Not rsHlp.EOF Then
                    gSBE = IIf(rsHlp("SBE") = "Y", True, False)
                    Exit For
                Else
                    gSBE = False
                End If
                rsHlp.Close
            Next
        End If
        
        If gSBE Then
            If rsHlp1.State Then rsHlp1.Close
            Set rsHlp1 = oConn.Execute("SELECT * FROM M_ESTIMASI WHERE noRangka = '" & nRgk & "' " _
                                     & "AND JT_TGL BETWEEN '" & gTglServer - 15 & "' AND '" & gTglServer + 15 & "' " _
                                     & "AND KMnext BETWEEN '" & nkm - 1500 & "' AND '" & nkm + 1500 & "' ORDER BY JT_Tgl DESC")
            If Not rsHlp1.EOF Then
                cekMRS = "GRP"
            Else
                cekMRS = ""
            End If
            rsHlp1.Close
        End If
    Else
        cekMRS = "SBE"
    End If
End Function

Function cekNoPol(nPol As String) As String
Dim lf As String
    
    lf = Mid(nPol, 2, 1)
    If Not IsNumeric(lf) Then
        cekNoPol = nPol
    Else
        cekNoPol = " " & nPol
    End If
End Function

