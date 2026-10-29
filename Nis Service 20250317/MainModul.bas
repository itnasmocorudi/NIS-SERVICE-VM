Attribute VB_Name = "MainModul"
Option Explicit

'==================================== VERSI NIS GRP ==============================================
'==================================== VERSI NIS GRP ==============================================

''Public Const gVersi As String = "Version 5.4.54"
Public Const gVersi As String = "Version 5.5.30" 'CABANG
''Public Const gVersi As String = "Version 9.4.51" 'CABANG
''Public Const gVersi As String = "Version 5.4.42" 'HO
Public Const oriFileNameNISService As String = "yarisbluelotus.exe"

Public Const newEnkrip As Boolean = False
Public Const dmsNonDss As Boolean = True
Public Const dmsDSS As Boolean = True

''20180207 CONST REPLIKASI untuk database yang dipilih
Public Const dbReplikasiHO      As String = "ReplBengkel99"
Public Const dbReplikasiHOSPK   As String = "RplNas"

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Long
Private Declare Function SendARP Lib "iphlpapi.dll" (ByVal DestIp As Long, ByVal ScrIP As Long, pMacAddr As Long, PhyAddrLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILE_TIME, lpLastAccessTime As FILE_TIME, lpLastWriteTime As FILE_TIME) As Long
Private Declare Function GetFileVersionInfo Lib "Verson.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OF_STRUCT, ByVal wStyle As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILE_TIME, lpLocalFileTime As FILE_TIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILE_TIME, lpSystemTime As SYSTEM_TIME) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function SHGetSpecialFolderPathW Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal lpszPath As Long, ByVal nFolder As Long, ByVal fCreate As Long) As Boolean
Private Const MD As Long = &H5
Private Const OF_READ = &H0
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OFS_MAXPATHNAME = 128

' ===== From Win32 Ver.h =================
' ----- VS_VERSION.dwFileFlags -----
Private Const VS_FFI_SIGNATURE = &HFEEF04BD
Private Const VS_FFI_STRUCVERSION = &H10000
Private Const VS_FFI_FILEFLAGSMASK = &H3F&

' ----- VS_VERSION.dwFileFlags -----
Private Const VS_FF_DEBUG = &H1
Private Const VS_FF_PRERELEASE = &H2
Private Const VS_FF_PATCHED = &H4
Private Const VS_FF_PRIVATEBUILD = &H8
Private Const VS_FF_INFOINFERRED = &H10
Private Const VS_FF_SPECIALBUILD = &H20

' ----- VS_VERSION.dwFileOS -----
Private Const VOS_UNKNOWN = &H0
Private Const VOS_DOS = &H10000
Private Const VOS_OS216 = &H20000
Private Const VOS_OS232 = &H30000
Private Const VOS_NT = &H40000
Private Const VOS__BASE = &H0
Private Const VOS__WINDOWS16 = &H1
Private Const VOS__PM16 = &H2
Private Const VOS__PM32 = &H3
Private Const VOS__WINDOWS32 = &H4

Private Const VOS_DOS_WINDOWS16 = &H10001
Private Const VOS_DOS_WINDOWS32 = &H10004
Private Const VOS_OS216_PM16 = &H20002
Private Const VOS_OS232_PM32 = &H30003
Private Const VOS_NT_WINDOWS32 = &H40004

' ----- VS_VERSION.dwFileType -----
Private Const VFT_UNKNOWN = &H0
Private Const VFT_APP = &H1
Private Const VFT_DLL = &H2
Private Const VFT_DRV = &H3
Private Const VFT_FONT = &H4
Private Const VFT_VXD = &H5
Private Const VFT_STATIC_LIB = &H7

' ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV -----
Private Const VFT2_UNKNOWN = &H0
Private Const VFT2_DRV_PRINTER = &H1
Private Const VFT2_DRV_KEYBOARD = &H2
Private Const VFT2_DRV_LANGUAGE = &H3
Private Const VFT2_DRV_DISPLAY = &H4
Private Const VFT2_DRV_MOUSE = &H5
Private Const VFT2_DRV_NETWORK = &H6
Private Const VFT2_DRV_SYSTEM = &H7
Private Const VFT2_DRV_INSTALLABLE = &H8
Private Const VFT2_DRV_SOUND = &H9
Private Const VFT2_DRV_COMM = &HA
Public OriName As String
Public FileVersion As String

Private Type VS_FIXEDFILEINFO
dwSignature As Long
dwStrucVersionl As Integer ' e.g. = &h0000 = 0
dwStrucVersionh As Integer ' e.g. = &h0042 = .42
dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
dwFileFlagsMask As Long ' = &h3F For version "0.42"
dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
dwFileType As Long ' e.g. VFT_DRIVER
dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
dwFileDateMS As Long ' e.g. 0
dwFileDateLS As Long ' e.g. 0
End Type


Public Type FILE_ATTRIBUTES
bArchive As Boolean
bCompressed As Boolean
bDirectory As Boolean
bHidden As Boolean
bNormal As Boolean
bReadOnly As Boolean
bSystem As Boolean
bTemporary As Boolean
End Type

Public Type FILE_INFORMATION
cFilename As String
cDirectory As String
cFullFilePath As String
cFileType As String
nVerMajor As Long
nVerMinor As Long
nVerRevision As Long
nVerNotUsedVB As Long
nFileSize As Long
nFileAttributes As Long
nFileType As Long
faFileAttributes As FILE_ATTRIBUTES
dtCreationDate As Date
dtLastModifyTime As Date
dtLastAccessTime As Date
sCompanyName As String
sFileDescription As String
sFileVersion As String
sInternalName As String
sLegalCopyright As String
sOriginalFileName As String
sProductName As String
sProductVersion As String
End Type

Private Type SYSTEM_TIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type

Private Type FILE_TIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Private Type OF_STRUCT
cBytes As Byte
fFixedDisk As Byte
nErrCode As Integer
Reserved1 As Integer
Reserved2 As Integer
szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Const MF_BYPOSITION = &H400&

Dim hMenu As Long

'==================================== VERSI NIS GRP ==============================================
'==================================== VERSI NIS GRP ==============================================
     
    
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

Private Const MAX_PATH = 260
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" _
    Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim strBuffer As String
Dim lngReturn As Long
Dim strWindowsSystemDirectory As String
Dim PathWindows As String

Public nData(2) As String

Public isEnkripsi As Boolean

Public gManPart As Boolean
Public gTglManAw As Date
Public gTglManAk As Date

Public allowIn As Boolean
Public gSudahInv As Boolean
Public gGudang As String
Public gGenerasi As String
Public gFiskal As String
Public gKodeRetur As String
Public gBlnBedaPart As String

Public gTglPart As String
Public gUkuranOrder As Boolean
Public gUkuranInvoice As Boolean
Public gUkuranRetur As Boolean
Public gPreviewCetak As Boolean
Public gIERC As Boolean
Public gFPart As Boolean
Public gCASH As Boolean
Public gJurnalPart As Boolean

Public gDiscMax As Integer
Public gUpSublet As Integer
Public gBank As String
Public gAcc As String
Public gRefreshInfo As Integer
Public gNoShow As Integer
Public gStatKoneksi As Boolean
Public gQueue As Boolean
Public statusKoneksi As String

Public gInvNo As String
Public gInvTgl As String
Public gInvCust As String
Public gInvAlmt As String
Public gInvKota As String
Public gInvNPWP As String
Public gInvPKP As String
Public gInvNoPolisi As String
Public gInvModel As String
Public gInvNoMesin As String
Public gInvNoChas As String
Public gInvNoOrder As String
Public gInvSA As String
Public gInvStatus As String
Public gInvMech As String
Public gInvPrint As String
Public gInvBilling As String
Public gInvUraian1 As String
Public gInvUraian2 As String
Public gInvUraian3 As String
Public gInvUraian4 As String
Public gInvUraian5 As String
Public gInvUraian6 As String
Public gInvUraian7 As String
Public gInvRate1 As String
Public gInvRate2 As String
Public gInvRate3 As String
Public gInvRate4 As String
Public gInvRate5 As String
Public gInvRate6 As String
Public gInvRate7 As String
Public gInvRp1 As String
Public gInvRp2 As String
Public gInvRp3 As String
Public gInvRp4 As String
Public gInvRp5 As String
Public gInvRp6 As String
Public gInvRp7 As String
Public gGInvRp As String
Public gInvTRp1 As String
Public gInvTRp2 As String
Public gInvTRp3 As String
Public gInvTRp4 As String
Public gInvTRp5 As String
Public gInvTRp6 As String
Public gInvTRp7 As String
Public gGInvTRp As String
Public gInvTerbilang As String
Public gInvSS As String
Public gInvCatatan As String
Public gInvSM As String

Public gNoOrder As String
Public gNoInvoice As String
Public gKetSA As String
Public gCust As String
Public gAlmt As String
Public gKota As String
Public gKotaCbg As String
Public gKotaFaktur As String
Public gTelpon As String
Public gCatatan As String
Public gNoPolisi As String
Public gModel As String
Public gWarna As String
Public gTglMasuk As String
Public gJamMasuk As String
Public gDealer As String
Public gDD As String
Public gMDLOrder As String
Public gFRM As String
Public gEG As String
Public gKeluhan(5) As String
Public gOrder(10) As String
Public gSublet(10) As String
Public gPartKd(7) As String
Public gPart(7) As String
Public gPartQty(7) As String
Public gCatatanB(4) As String
Public gKM As String
Public gPKhusus As String
Public gSAdvisor As String
Public gPSelesai As String
Public gTahun As String
Public gKodePlg As String
Public tPesan As String
Public gKdCbg As String
Public gNmCbg As String
Public gTglKukuh As String
Public gNmHO As String
Public gAlmtHO As String
Public gKotaHO As String
Public gTelpHO As String
Public gNpwpHO As String


Public gKdPlgTAM As String
Public gNmTAM As String
Public gAlmtTAM As String
Public gKotaTAM As String
Public gTelpTAM As String
Public gNpwpTAM As String

Public gCopyright As String
Public gSecmanager As String
Public gADH As String
Public gNpwp As String
Public gPkp As String
Public gTelp As String
Public gNoPolis As String
Public gGroupModel As String
Public gServerDb As String
Public gDb As String
Public gDbVersi As String
Public gHrSama As Boolean
Public dRangka As String
Public gMaxHariBooking As Integer


Public LogonStr As String
Public LogonStrHO As String
Public LogonStrProfil As String

Public LogonSrNis As String
Public SQLCONBengkelHO As String
Public SQLCONBengkelProfil As String

Public oConn As New ADODB.Connection
Public aConn As New ADODB.Connection
Public oConnAcces As New ADODB.Connection
Public oConnAcces2 As New ADODB.Connection
Public oConnNis As New ADODB.Connection
Public oConnHO As New ADODB.Connection
Public oConnSPK As New ADODB.Connection
Public oConnSPKHO As New ADODB.Connection
Public oConnProfil As New ADODB.Connection
Public oConnLog As New ADODB.Connection
Public oConnNisSP As New ADODB.Connection
Public oConnNisSP2 As New ADODB.Connection
Public oConnTemp As New ADODB.Connection
Public oConnQueue As New ADODB.Connection

Public KonDbs As String
Public KonDbNis As String
Public KonDbNisSP As String
Public KonDbNisSP2 As String
Public KonDbAcces As String
Public KonDbAcces2 As String
Public KonDbHO As String
Public KonDbSPK As String
Public KonDbSPKHO As String
Public adoCon As String
Public adoConAcces As String
Public adoConAcces2 As String
Public adoConNis As String
Public adoConNisSP As String
Public adoConNisSP2 As String
Public adoConHO As String
Public adoConSPK As String
Public adoConSPKHO As String
Public NisDb2 As String

Public gsSql1(200) As String
Public gsSql2(200) As String
Public oRs As New ADODB.Recordset
Public oRs1 As New ADODB.Recordset
Public oRs2 As New ADODB.Recordset
Public oRs3 As New ADODB.Recordset
Public oRsNis As New ADODB.Recordset
Public sSQL As String
Public sSQLi As String
Public sSQLi2 As String
Public sSQLi3 As String
Public sSQLHO As String
Public sSQL1 As String
Public gPrinter As String
Public gPrintCetak As String
Public gMainKeyCari As String
Public gBantuan As String
Public gBantuan2 As String
Public gBantuan3 As String
Public gBantuan4 As String
Public gStall As String
Public gSendTab As Boolean
Public gSearchOption As String
Public gStok As Integer
Public gNamaPart As String
Public gSatPart As String
Public gHargaJual As Double


Public gHasilSearch1    As String
Public gHasilSearch2    As String
Public gHasilSearch3    As String


Public gTglDefault As String
Public gMDL As String
Public gRateCOM As Double
Public gRatePAS As Double
Public gRateIRC As Double
Public gLimitFUDealer As Integer
Public gTglInv As String
Public gFBackup As String
Public gFont As Integer
Public gHeight As Integer

Public gTanggal As String
'Public isnow As Date
Public gSA As String
Public gDiv As String
Public gNama_SA As String
Public gKd_Sama As String

Private gDb2 As String
Private gDbVersiLama As String
Private gDbTanggal As String

Public OnTopSV As Integer
Public OnNewFP As Integer
Public OnDisKBG As Double
Public OnNomor As Double

Public gPerush As String
Public gAlamat As String
Public gLoket As String
Public gNoAntri As Integer
Public gTpAntri As String
Public gCabang As String
Public gTglTemp As String
Public gTglTemp1 As String
Public gTglTemp2 As String
Public gKetTemp As String
Public gReportConn As String
Public gUser As String
Public gPwd As String
Public Encry As New acak.Class
Public gReturn As Integer
Public gKode As String
Public gKodeFaktur As String
Public gConnecAAS As Boolean
Public gTag As String
Public allowMove As Boolean
Public gTipeUser As String
Public tJamSelesai As String
Public adaJobStopped As Boolean
Public fgCabang As Boolean
Public gHo As Boolean
Public gStat As String
Public gUstat As Integer
Public gKdPlgHo As String
Public gAmbilPlg As String
Public flagOk As Boolean
Public FileName As String
Public Gif As String
Public hTipePenghitunganDealer As String
Public ComptName As String
Public hLimitMou As Integer
Dim C As New Convert
Dim objBilling As New BClsBilling
Dim support As New clsFormDealer
Dim allowSaveLog As Boolean
Public gTglUpdate As Date

Public gUserKoneksi As String
Public gPasswordKonesi As String

Public objA As New ClsForm
Public fHeight As Double

'variabel form aktif
Public bookAktif As Boolean
Public orderAktif As Boolean
Public stallAktif As Boolean
Public validAktif As Boolean
Public subletAktif As Boolean
Public subRetAktif As Boolean
Public nC10 As Integer
Public updateHistory As Boolean

'variabel antrian
Public antrian As String
Public sisaBook As Integer
Public sisaNonBook As Integer

'Tabel Stall
Public MyArray As New XArrayDB
Public MyArray2 As New XArrayDB
Public MyArray3 As New XArrayDB
Public rsTmpPart As New ADODB.Recordset
Public gVerifikator As String

' tampilan layar
Public rsLayar As New ADODB.Recordset

' TEMP SPK SURAT
Public sNamaSpk As String
Public sAlamatSpk As String
Public sKotaSpk As String
Public sKodeposSpk As String
Public sTelpSpk As String
Public sGsmSpk As String
Public sCdmaSpk As String
Dim obj As New clsRepo

Public OnKPI10 As Boolean
'.:: ======================================== .:: BEGIN ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: M R S ::. ========================================= ::.
Private Const PathMrs32 As String = "C:\WINDOWS\system32\Mrs32.Ant"
Public SQLCon        As String
Public gNIK          As String

Type LoginServer
     LogServ         As String
     LogUser         As String
     LogPass         As String
     LogNas          As String
     LogAlamat       As String
     LogBengkel      As String
     LogSPK          As String
     LogQueue        As String
     LogFaktur       As String
     LogACC          As String
     LogSMS          As String
    
     LogHODb         As String
     LogHOUser       As String
     LogHOPass       As String
     LogHOKomp       As String
     
     LogSPKDb        As String
     LogSPKUser      As String
     LogSPKPass      As String
     LogSPKKomp      As String
     
     LogFakturDb     As String
     LogFakturUser   As String
     LogFakturPass   As String
     logFakturKomp   As String
               
     LogHODbSPK         As String
     LogHOUserSPK       As String
     LogHOPassSPK       As String
     LogHOKompSPK       As String
     LogKbgName         As String
     
     LogHOConn       As Integer
     
     LogNamaNas      As String
     LockAR          As Boolean
     isFill          As Integer
     AppRate         As String

     LogAccServ         As String
     LogAccUser         As String
     LogAccPass         As String
     
     startAppRate   As Date
     
     'TAMBAHAN FLAG PENYERAGAMAN AKUN DAN AKUN LAMA
     'TANGGAL 04-09-2014
     'RUDI KRISTANTO
     
     isAkun         As Integer
     expPwd         As Integer
     
     estInv         As Integer
     isEmStall      As Integer
     
     'terkait dms non dss
     
     kmwarranty     As Double
     warranty       As Integer
     dmsNonDss      As Integer
     persentaseDP   As Integer
     plafonDP       As Double
     KodeDealer     As String
     barcode        As Integer
     smsBilling     As Integer
     
     isDSS          As Integer
     isWashing      As Integer
     formatSmsBilling     As String
     helpJpcb       As Integer
     newJpcb        As Integer
     
     validasiKabeng As Integer
     FIBeforeSublet As Integer
     prioritas      As String
     StartAffPpn    As Date
     lockWashing     As Integer
     FIR            As Integer
     
     ''2016-07-15 || penambahan log
     followUp        As Integer
     prosentaseFollow   As Integer
     OTD             As Integer
     
     'adit disableexport 20161220 -begin
     PdfOnly             As Integer
     'adit disableexport 20161220 -end
     
     ' update 31 jan 2017, penambahan field utk penanda proses update yg dilakukan hanya 1x dalam 1 hari
     tglUpdtEvd   As Date
     
     '' 2017-03-30 DETAIL WAC
     detWAC         As Integer
     
     '' 2017-04-18 LOCK WAC
     lockWAC        As Integer
     
     ''2017-04-28 ON EXPORT DISABLE
     onExport       As Integer
     
     ''2017-05-03 DISC KACAB
     discKacab      As Integer
     startDiscKacab As Date
     
     ''2017-06-06 NILAI MAX DISC KACAB
     discJasa       As Integer
     discMaxJasa    As Integer
     discPart       As Integer
     discPartMax    As Integer
     
     notifInvNCM    As Integer
     notifInvNonNCM As Integer
     
     repoHarianMech As Integer
     loyalityCustomer As Integer
     ketNasApp        As String
     
     RPLHistori     As Integer
     LogSIDServ         As String
     LogSIDDb           As String
     LogSIDUser         As String
     LogSIDPass         As String
     
     logServBP          As String
     logDbBP            As String
     
     mandatoryKTP       As Integer
     logRetention       As Integer
     
     '' 20180518DEV01'INSPECTION
     CR7                As Integer
     
     FTPCR7User         As String
     FTPCR7Pass         As String
     FTPCR7Host         As String
     
     '' 20180611DEV02'TAMBAHMATERAI
     fakturMaterai      As Integer
     
     '' 20180611DEV01'CETAKANBARU
     newCetak           As Integer
     
     replikasiWacHO     As Integer
     
     lockSCC            As Integer
     
     NoKtpKetFP         As Integer
     NoKtpNonPerorangan As Integer
     
     '' 20180927 SCORING SINGLE ID
     singleIDScoring    As Integer
     
     '' 20181018DEV1 | DEV | DEVELOPMENT  | PENAMBAHAN FLAG UNTUK STANDARISASI SUBLET
     standarisasiSublet As Integer
     
     '' 20190206REV1 | REV | REVISI  | PENAMBAHAN FLAG UNTUK PENGECEKAN NO KTP DAN NO HP
     lockKTPNoHPScoring As Integer
     
     '' 20190625DEV1
     lockFreeSVCHO      As Integer
     tglFreeSVCHO       As Date
     
     ''20190701REV1
     newLockSSC         As Integer
     
     '' 20190816DEV1
     hansaRenkei        As Integer
     
     ''20191115DEV1
     rsA                As Integer
     tglRSA             As Date
     
     ''20200110DEV1
     addMaterai         As Integer
     
     '' 20200113DEV1
     newOhbeya          As Integer
     
     '' 20200210DEV
     multipleMarketing  As Integer
     tglmultipleMarketing   As Date
     
     '' 20200303DEV
     vinRetention       As Integer
     
     '' 20200408DEV
     sentralisasiStars  As Integer
     kodeCabang         As String
     
     '' 20200416DEV
     freepart           As Integer
     
     '' 20200512DEV
     disableExportMra   As Integer
     
     ''20200709DEV
     monitoringScoring  As Integer
     
     ''20200716DEV
     discMaxSublet      As Integer
     
     '' 20200915DEV
     hansaRenkeiBP      As Integer
     
     '' 20201124DEV
     vin2vinRev1        As Integer
     
     '' 20201218DEV
     materai2021        As Integer
     
     ''20201222DEV
     NHS                As Integer
     
     ''20210520DEV
     simpanRepl         As Integer
     
     ''20211116DEV
     flagOffBookingLocal    As Integer
     
     '' 20211125DEV
     FlagSalesHO        As Integer
     
     '' 20220210DEV
     hideService        As Integer
     
     ''20230428DEV
     tcm                As Integer
     tcmMaxHari         As Integer
     
     ''20240119DEV
     DIDX               As Integer
     kdDealerDidx       As String
     DiDxP1             As Integer
     dpplain             As Integer
End Type

'' ============================================
'' 20180907
'' ============================================

Type Profile
    gNamaProfil         As String
    AlamatProfil1       As String
    KotaProfil          As String
    TeleponProfil       As String
    FaxProfil           As String
    npwp                As String
    SKPengukuhan        As String
    TglPengukuhan       As Date
    nilaippn            As Double ' 1/11
    nilaitotal          As Double ' 1+10 = 11
    nilaidpp            As Double ' 10/11
End Type
'' ============================================

'' 20180927DEV  ' SCORING SINGLE ID
Type ProfileScoringSID
    SearchingKTP        As Boolean
    SearchingHP         As Boolean
    SearchingNama       As Boolean
    ScoringNamaKhusus   As Boolean
    nScoringNamaKhusus  As Single
    LockNama1Khusus     As Single
    ScoringNamaSimultan As Boolean
    LockNama1           As Single
    LockNama2           As Single
    nScoringNama2       As Single
    ScoringNoHP         As Boolean
    ScoringTglLahir     As Boolean
    ScoringKelurahan    As Boolean
    ScoringKecamatan    As Boolean
End Type
'' ============================================

Public Log              As LoginServer
Public mConnHO          As ADODB.Connection
Public SQLConBengkel99  As String
Public ConHO            As Boolean
Public profilACC        As Profile
Public ProfilScoring    As ProfileScoringSID

'SMS... Public Const AccessCon  As String = "Provider = Microsoft.JET.OLEDB.4.0;Data Source =  C:\Inetpub\wwwroot\sms\dbSMS.mdb"

Public isnow As Date
Public isNowBilling As Date
Public gTanggalStall As Date


'.:: ======================================== .:: M R S ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: END ::. ========================================= ::.

Public SQLConProfilBK As String
Public SQLConProfilSP As String
Public SQLCONBengkel  As String
Public SQLConAAS      As String
Public SQLConACC      As String
Public SQLConACC1     As String
Public SQLConACC2     As String
Public SQLConFP       As String
Public SQLConSPK      As String
Public SQLConSPKHO    As String
Public SQLConQueue    As String
Public SQLConHO       As String
Public SQLConSMS       As String
Public SQLConSID      As String
Public SQLConBP       As String

Private DBConBengkel  As ADODB.Connection
Private DBConACC      As ADODB.Connection
Private DBConFP       As ADODB.Connection
Private DBConSMS      As ADODB.Connection
Private DBConSPK      As ADODB.Connection
Private DBConQueue    As ADODB.Connection
Private DBConSingleID As ADODB.Connection

Public erTranHO       As Boolean
Public erTranAll      As Boolean
Public errHO          As Boolean

Public objHO          As New ClsTransferHO
Public objO           As New oClassNonTrans
Public objHT          As New oClassHOTrans
Public onTR           As Boolean
Public goHo           As Boolean
Public ndFP           As Double
Private rs            As ADODB.Recordset
Public rsMK           As New ADODB.Recordset
Public UpdateMK(100)  As String
Public nMK            As Integer
Public LoopMk         As Boolean
Public objSID         As New clsSID

Public Const fColor = &HC0E0FF          'Merah
Public Const tColor = &H80000005        'Putih
Public Const tColorx = &HC0FFFF         'Kuning
Public Const fColorz = &H202020         'Hitam
Public Const tColorz = &H0&             'Hitam Pekat

Public ShowForm     As String
Public nHis         As Boolean
Public onTimer      As Boolean

Public objSt As New clsStall
Public nDiscNRM     As Double
Public nDiscNON     As Double
Public myChassis    As String
Public Const nJualNRM As String = "NCM" ''"PENJUALAN NASMOCO GROUP"
Public isJualNRM As Boolean
Public isOnCarClass As Boolean
Public tglJualNRM As Date

'' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
Public varErlError  As Long
Public varErrNumber As Long
Public gAlamatDepo  As String

Public Function Koma2Titik(nilai As Variant) As Double
    On Error GoTo Keluar
    Dim temp As String
    Dim i As Integer
    
    temp = str(nilai)
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

Public Sub BuatTableValidasiTWC(namaTable As String)
    Call support.dropTable(namaTable)
    Call support.createTableValidasi(namaTable)
    
End Sub

Public Sub BuatTabelRepo(vNamaTabel As String)
    Call support.dropTable(vNamaTabel)
    Call support.createTableRepo(vNamaTabel)
End Sub

Public Sub BuatTabelRepoUnitMasuk(vNamaTabel As String)
    Call support.dropTable(vNamaTabel)
    Call support.createTableRepoUnitMasuk(vNamaTabel)
    
End Sub

Public Sub BuatTabelInv(vNamaTabel As String)
    Call support.dropTable(vNamaTabel)
    Call support.createTableInv(vNamaTabel)
End Sub

Public Function SimpanKoma2Titik(nilai As Variant) As String
    On Error GoTo Keluar
    Dim temp As String
    Dim i As Integer
    
    If nilai <> "" Then
        temp = str(nilai)
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

Public Function UbahKoma2Titik(nilai As String) As String
    On Error GoTo Keluar
    Dim temp As String
    Dim i As Integer
    
    If nilai <> "" Then
        temp = nilai
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
        Sendkeys ("0")
    Case vbKeyF2
        Sendkeys ("00")
    Case vbKeyF3
        Sendkeys ("000")
    Case vbKeyF4
        Sendkeys ("0000")
    Case 38
        Sendkeys ("+{TAB}")
    Case 40, 13
        Sendkeys ("{TAB}")
    End Select
End Sub

Public Sub CekTombolCombo(Akey As Integer)
    Select Case Akey
    Case vbKeyF1
        Sendkeys ("0")
    Case vbKeyF2
        Sendkeys ("00")
    Case vbKeyF3
        Sendkeys ("000")
    Case vbKeyF4
        Sendkeys ("0000")
    Case 13
        Sendkeys ("{TAB}")
    End Select
End Sub

Public Function CekNumerik(Akey As Integer) As Integer
    Select Case Akey
        ' 44,  -> UNTUK TANDA KOMA (,)
        Case 48 To 57, 8
            CekNumerik = Akey
        Case Else
            CekNumerik = 0
            Sendkeys ("")
    End Select
End Function

Public Function CekAlfabet(Akey As Integer) As Integer
    Select Case Akey
        ' 44,  -> UNTUK TANDA KOMA (,)
        Case 65 To 90
            CekAlfabet = Akey
        Case Else
            CekAlfabet = 0
            Sendkeys ("")
    End Select
End Function

Public Function Angka2Huruf(vNilai As Double, Jenis As Byte) As String
1         Dim xNilai As String: Dim xPjgNilai As Byte: Dim xProses As String
2         Dim xHasil As String: Dim xHasil1 As String: Dim i As Integer: Dim PerTiga As Byte
3         Dim Koma As Byte: Dim KomaPosisi As Byte: Dim xxNilai As String: Dim xHasil2 As String
          Dim Sen As Boolean
          Dim DiKoma As Integer, Ambil As Integer
          Dim j As Integer, k As Integer
          Dim Spasi As Integer
          
4        On Error GoTo Angka2Huruf_Error

5         xxNilai = Trim(str(vNilai))
6         Koma = 1
7         KomaPosisi = 0
8         For i = 1 To Len(xxNilai)
9             If Mid(xxNilai, i, 1) = "." Then
10                Koma = 2
11                KomaPosisi = i
12            End If
13        Next
          
14        DiKoma = 0
15        Sen = False
16        If KomaPosisi <> 0 Then
17            If Len(xxNilai) - KomaPosisi <= 2 Then
18                If Len(xxNilai) - KomaPosisi <> 0 Then
19                    DiKoma = 2
20                    Sen = True
21                End If
22            Else
23                DiKoma = Len(xxNilai) - KomaPosisi
24            End If
25        End If
          
          Dim Formating As String
26        Formating = "#0"
27        If DiKoma <> 0 Then
28            Formating = Formating + "."
29            For i = 1 To DiKoma
30                Formating = Formating + "0"
31            Next
32        End If
33        xxNilai = Trim(Format(vNilai, Formating))
          
          
34        For j = 1 To Koma
35            xHasil = ""
36            xHasil1 = ""
37            If j = 1 Then
38                If KomaPosisi = 0 Then
39                    xNilai = xxNilai
40                Else
41                    xNilai = Mid(xxNilai, 1, KomaPosisi - 1)
42                End If
43            Else
44                xNilai = Mid(xxNilai, KomaPosisi + 1, Len(xxNilai))
45            End If
              
46            PerTiga = 0
47            If vNilai <> "0" Then
48                Do While xNilai <> ""
49                   xProses = Right(xNilai, 3)
50                   xNilai = Left(xNilai, Len(xNilai) - Len(xProses))
51                   If xProses <> "000" Then
52                        Select Case PerTiga
                          Case 1
53                             xHasil = " ribu" + xHasil
54                        Case 2
55                             xHasil = " juta" + xHasil
56                        Case 3
57                             xHasil = " miliar" + xHasil
58                        End Select
59                   End If
                     
60                   For i = Len(xProses) To 1 Step -1
61                       If Mid(xProses, i, 1) <> "0" Then
62                            Select Case Len(xProses) - i
                              Case 2
63                                 xHasil = " ratus" + xHasil
64                            Case 1
65                                 If Mid(xProses, i, 1) <> "1" Then
66                                    xHasil = " puluh" + xHasil
67                                 Else
68                                    If Mid(xProses, Len(xProses), 1) = "0" Then
69                                        xHasil = " sepuluh" + xHasil
70                                    Else
71                                        Spasi = 20
72                                        For k = 2 To 10
73                                            If Mid(xHasil, k, 1) = " " Then
74                                                Spasi = k
75                                                Exit For
76                                            End If
77                                        Next
78                                        If Spasi = 20 Then
79                                            xHasil = xHasil + " belas"
80                                        Else
81                                            xHasil = Left(xHasil, (k - 1)) + " belas" + Right(xHasil, Len(xHasil) - (k - 1))
82                                        End If
83                                        If Left(xHasil, 5) = " satu" Then
84                                            xHasil = " se" + Right(xHasil, Len(xHasil) - 6)
85                                        End If
86                                    End If
87                                 End If
88                            End Select
89                       End If
90                       Select Case Mid(xProses, i, 1)
                         Case "1"
91                            If Len(xProses) - i = 0 Then
92                                If PerTiga = 1 And Len(xProses) = 1 Then
93                                    xHasil = " se" + xHasil
94                                Else
95                                    xHasil = " satu" + xHasil
96                                End If
97                            Else
98                                If Len(xProses) - i <> 1 Then
99                                    xHasil = " se" + xHasil
100                               End If
101                           End If
102                      Case "2"
103                           xHasil = " dua" + xHasil
104                      Case "3"
105                           xHasil = " tiga" + xHasil
106                      Case "4"
107                           xHasil = " empat" + xHasil
108                      Case "5"
109                           xHasil = " lima" + xHasil
110                      Case "6"
111                           xHasil = " enam" + xHasil
112                      Case "7"
113                           xHasil = " tujuh" + xHasil
114                      Case "8"
115                           xHasil = " delapan" + xHasil
116                      Case "9"
117                           xHasil = " sembilan" + xHasil
118                      End Select
119                  Next
120                  PerTiga = PerTiga + 1
121               Loop
122           Else
123               xHasil = " Nol"
124           End If
125           Ambil = 1
126           For i = 1 To Len(xHasil) - 3
127               If Mid(xHasil, i, 3) = "se " Then
128                   xHasil1 = xHasil1 + Mid(xHasil, Ambil, i + 2 - Ambil)
129                   Ambil = i + 3
130               End If
131           Next
132           xHasil1 = xHasil1 + Mid(xHasil, Ambil, Len(xHasil) - (Ambil - 1))
133           xHasil1 = UCase(Mid(xHasil1, 2, 1)) + Right(xHasil1, Len(xHasil1) - 2)
134           If Jenis = 1 Then
135               If j = 1 Then
136                   xHasil2 = xHasil1 + " Rupiah"
137               Else
138                   xHasil2 = xHasil2 + " " + xHasil1 + " Sen"
139               End If
140           Else
141               If j = 1 Then
142                   xHasil2 = xHasil1 + " US Dolar"
143               Else
144                   If Sen Then
145                       xHasil2 = xHasil2 + " " + xHasil1 + " Sen"
146                   Else
147                       xHasil2 = xHasil2 + " " + xHasil1
148                   End If
149               End If
150           End If
151       Next
152       Angka2Huruf = xHasil2

153      On Error GoTo 0
154      Exit Function

Angka2Huruf_Error:

155       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Angka2Huruf of Module MainModul")
End Function

Public Function FormatTgl(dDate As Date) As String
    FormatTgl = Format(dDate, "mm/dd/yyyy")
End Function

Public Function TglYYYYMMDD(kata As String) As String
    TglYYYYMMDD = Right(kata, 4) + "-" + Mid(kata, 4, 2) + "-" + Left(kata, 2)
End Function

Public Function TglDDMMYYYY(kata As Date) As String
    TglDDMMYYYY = Format(Day(kata), "00") + "-" + Format(Month(kata), "00") + "-" + Trim(str(year(kata)))
End Function

Public Function TglDDMMMMYYYY(kata As Date, Optional seri As Integer) As String
    If seri = 1 Then
        TglDDMMMMYYYY = Format(Day(kata), "00") + " " + GetNamaBulan(Month(kata), 1) + " " + Trim(str(year(kata)))
    Else
        TglDDMMMMYYYY = Format(Day(kata), "00") + " - " + GetNamaBulan(Month(kata), 1) + " - " + Trim(str(year(kata)))
    End If
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
    Else
        CekJam = False
    End If
End Function

Public Function KurangM1M2(Nilai1 As String, Nilai2 As String) As String
    Dim Menit1 As Double
    Dim Menit2 As Double
    Dim jam As String
    Dim menit As String
    
    Menit1 = (Val(Left(Trim(Nilai1), 2)) * 60) + Val(Right(Trim(Nilai1), 2))
    Menit2 = (Val(Left(Trim(Nilai2), 2)) * 60) + Val(Right(Trim(Nilai2), 2))
    
    jam = Format((Menit2 - Menit1) \ 60, "00")
    menit = Format((Menit2 - Menit1) Mod 60, "00")
    
    If jam >= 24 Or jam < 0 Then
        KurangM1M2 = "00:00"
    Else
        KurangM1M2 = Trim(jam) + ":" + Trim(menit)
    End If
End Function

Public Sub ClearsSql()
1     On Error GoTo vb_error
          Dim i As Integer
2         For i = 0 To 200
3             gsSql1(i) = ""
4             gsSql2(i) = ""
5         Next
6         Exit Sub
vb_error:
7         erTranAll = True
8         ErrMsg.judul = ""
9         ErrMsg.message = ErrMessage(Erl, err.Number, "Procedure : Clear SQL")
10        ErrMsg.show
          
      '    MsgBox ErrMessage(Erl, err.Number, "Procedure : Clear SQL"), vbInformation
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

Public Function GetNamaBulan(ByVal blnKe As Byte, Optional tipe As Integer = 0) As String
    Select Case tipe
        Case 0
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
        Case 1
            Select Case blnKe
                Case 1
                    GetNamaBulan = "Januari"
                Case 2
                    GetNamaBulan = "Februari"
                Case 3
                    GetNamaBulan = "Maret"
                Case 4
                    GetNamaBulan = "April"
                Case 5
                    GetNamaBulan = "Mei"
                Case 6
                    GetNamaBulan = "Juni"
                Case 7
                    GetNamaBulan = "Juli"
                Case 8
                    GetNamaBulan = "Agustus"
                Case 9
                    GetNamaBulan = "September"
                Case 10
                    GetNamaBulan = "Oktober"
                Case 11
                    GetNamaBulan = "November"
                Case 12
                    GetNamaBulan = "Desember"
            End Select
    End Select
End Function

Public Function GetAngkaBulan(Nmbln As String) As Integer
    Select Case Nmbln
        Case "JANUARI"
            GetAngkaBulan = 1
        Case "FEBRUARI"
            GetAngkaBulan = 2
        Case "MARET"
            GetAngkaBulan = 3
        Case "APRIL"
            GetAngkaBulan = 4
        Case "MEI"
            GetAngkaBulan = 5
        Case "JUNI"
            GetAngkaBulan = 6
        Case "JULI"
            GetAngkaBulan = 7
        Case "AGUSTUS"
            GetAngkaBulan = 8
        Case "SEPTEMBER"
            GetAngkaBulan = 9
        Case "OKTOBER"
            GetAngkaBulan = 10
        Case "NOVEMBER"
            GetAngkaBulan = 11
        Case "DESEMBER"
            GetAngkaBulan = 12
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
    Call support.dropTable(vNmTbl)
End Sub

Public Function SysAcces() As String
    SysAcces = objO.FetchData_String("SELECT FS_SYS_AAS FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function LetakAcces() As String
    LetakAcces = objO.FetchData_String("SELECT FS_DB_AAS FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function HOKomp() As String
    HOKomp = objO.FetchData_String("SELECT FS_HO_KOMP FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function HOUser() As String
    If newEnkrip Then
        HOUser = Dekripsi(objO.FetchData_String("SELECT FS_HO_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(3))
    Else
        HOUser = objO.FetchData_String("SELECT FS_HO_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    HOUser = Log.LogHOUser 'Dekripsi(objO.FetchData_String("SELECT FS_HO_USER FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(3))
End Function

Public Function HOPass() As String

    If newEnkrip Then
        HOPass = Dekripsi(objO.FetchData_String("SELECT FS_HO_PASS FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(5))
    Else
        HOPass = objO.FetchData_String("SELECT FS_HO_PASS FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    HOPass = Log.LogHOPass 'Dekripsi(objO.FetchData_String("SELECT FS_HO_PASS FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(5))
End Function

Public Function HODb() As String
    HODb = objO.FetchData_String("SELECT FS_HO_DB FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function NisKomp() As String
    NisKomp = objO.FetchData_String("SELECT FS_NIS_Komp FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function NisUSER() As String
    If newEnkrip Then
        NisUSER = Dekripsi(objO.FetchData_String("SELECT FS_NIS_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(3))
    Else
        NisUSER = objO.FetchData_String("SELECT FS_NIS_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    NisUSER = Log.LogAccUser 'Dekripsi(objO.FetchData_String("SELECT FS_NIS_USER FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(3))
    If NisUSER = "" Then NisUSER = "SA"
End Function

Public Function NisPass() As String
    If newEnkrip Then
        NisPass = Dekripsi(objO.FetchData_String("SELECT FS_NIS_Pass FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(5))
    Else
        NisPass = objO.FetchData_String("SELECT FS_NIS_Pass FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    NisPass = Log.LogAccPass 'Dekripsi(objO.FetchData_String("SELECT FS_NIS_Pass FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(5))
End Function

Public Function NisDb() As String
    NisDb = objO.FetchData_String("SELECT FS_NIS_Db FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

'///// SPAREPART NIS /////  --> KONEKSI UNTUK AMBIL DATA INV SPAREPART
Public Function NisSPKomp() As String
    NisSPKomp = objO.FetchData_String("SELECT FS_NISSP_Komp FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function NisSPUSER() As String
    If newEnkrip Then
        NisSPUSER = Dekripsi(objO.FetchData_String("SELECT FS_NISSP_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(3))
    Else
        NisSPUSER = "SA" 'objO.FetchData_String("SELECT FS_NISSP_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    NisSPUSER = Log.LogAccUser 'Dekripsi(objO.FetchData_String("SELECT FS_NISSP_USER FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(3))
End Function

Public Function NisSPPass() As String
    If newEnkrip Then
        NisSPPass = Dekripsi(objO.FetchData_String("SELECT FS_NISSP_Pass FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(5))
    Else
        NisSPPass = objO.FetchData_String("SELECT FS_NISSP_Pass FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    NisSPPass = Log.LogAccPass 'Dekripsi(objO.FetchData_String("SELECT FS_NISSP_Pass FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(5))
End Function

Public Function NisSPDb() As String
    NisSPDb = objO.FetchData_String("SELECT FS_NISSP_Db FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function NisSPDb2() As String
    NisSPDb2 = objO.FetchData_String("SELECT FS_NISSP_Db2 FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function SpkKomp() As String
    SpkKomp = objO.FetchData_String("SELECT FS_SPK_Komp FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function SpkUSER() As String
    If newEnkrip Then
        SpkUSER = Dekripsi(objO.FetchData_String("SELECT FS_SPK_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(3))
    Else
        SpkUSER = objO.FetchData_String("SELECT FS_SPK_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    SpkUSER = Log.LogSPKUser 'Dekripsi(objO.FetchData_String("SELECT FS_SPK_USER FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(3))
    If SpkUSER = "" Then SpkUSER = "SA"
End Function

Public Function SpkPass() As String
    If newEnkrip Then
        SpkPass = Dekripsi(objO.FetchData_String("SELECT FS_SPK_Pass FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(5))
    Else
        SpkPass = objO.FetchData_String("SELECT FS_SPK_Pass FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    SpkPass = Log.LogSPKPass 'Dekripsi(objO.FetchData_String("SELECT FS_SPK_Pass FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(5))
End Function

Public Function SpkDb() As String
    SpkDb = objO.FetchData_String("SELECT FS_SPK_Db FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function SpkHOKomp() As String
    SpkHOKomp = objO.FetchData_String("SELECT FS_SPKHO_Komp FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Public Function SpkHOUSER() As String
    If newEnkrip Then
        SpkHOUSER = Dekripsi(objO.FetchData_String("SELECT FS_SPKHO_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(3))
    Else
        SpkHOUSER = objO.FetchData_String("SELECT FS_SPKHO_USER FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    SpkHOUSER = Log.LogHOUserSPK 'Dekripsi(objO.FetchData_String("SELECT FS_SPKHO_USER FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(3))
    If SpkHOUSER = "" Then SpkHOUSER = "SA"
End Function

Public Function SpkHOPass() As String
    If newEnkrip Then
        SpkHOPass = Dekripsi(objO.FetchData_String("SELECT FS_SPKHO_Pass FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)"), KataKunci(5))
    Else
        SpkHOPass = objO.FetchData_String("SELECT FS_SPKHO_Pass FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
    End If
    SpkHOPass = Log.LogHOPassSPK 'Dekripsi(objO.FetchData_String("SELECT FS_SPKHO_Pass FROM " & Log.LogBengkel & "..encT_PROFIL WITH (NOLOCK)"), KataKunci(5))
End Function

Public Function SpkHODb() As String
    SpkHODb = objO.FetchData_String("SELECT FS_SPKHO_Db FROM " & Log.LogBengkel & "..T_PROFIL WITH (NOLOCK)")
End Function

Function McDelIndex(ModulId As String) As String
   Dim acak As Double
   Randomize Timer
   acak = Int(Rnd * 99999999)
   McDelIndex = ModulId & acak
End Function

Public Sub getLogin()
    On Error GoTo vbError
    Dim Nama As String
    Dim a As Integer
    Open FileName For Input As #1
    Do Until EOF(1)
    Line Input #1, Nama
        a = a + 1
        Select Case a
            Case 1: gUser = Trim(C.decryp_pass(21, Trim(Nama)))
            Case 2: gPwd = Trim(C.decryp_pass(21, Trim(Nama)))
        End Select
        If a = 2 Then Exit Do
    Loop
    Close #1
vbError:
    
End Sub

Public Function getShell(nameFile As String) As String
    On Error GoTo vbError
    Dim strComplete
    Dim Nama As String
    Dim a As Integer
    
    strComplete = ""
    Open nameFile For Input As #1
    Do Until EOF(1)
    Line Input #1, Nama
        strComplete = strComplete & Trim(Nama)
    Loop
    Close #1
    getShell = strComplete
    Exit Function
    
vbError:
    Exit Function
End Function

Sub Main()
      Dim a As String

1         On Error GoTo Setup
          
2         strBuffer = Space$(MAX_PATH)
3         lngReturn = GetSystemDirectory(strBuffer, MAX_PATH)
4         strWindowsSystemDirectory = Left$(strBuffer, Len(strBuffer) - 1)
5         PathWindows = Trim(strWindowsSystemDirectory)
6         FileName = Left(PathWindows, Len(PathWindows) - 1) & "\" & "SysWin.svc"
7         Gif = Left(PathWindows, Len(PathWindows) - 1) & "\" & "SVC.gif"
          Dim Counter     As Long
          Dim X           As String
8         Counter = 200
9         X = String(Counter, 0)

10        GetComputerName X, Counter
11        ComptName = Left(X, Counter)
12        Call getLogin
          
      '''    gServerDb = GetSetting(appname:="SERVICE", Section:="Server", Key:="Top", Default:="(local)")
      '''    gUser = IIf(gUser = "", GetSetting(appname:="SERVICE", Section:="User", Key:="Top", Default:="sa"), gUser)
      '''    gPwd = IIf(gPwd = "", GetSetting(appname:="SERVICE", Section:="Pwd", Key:="Top", Default:=""), gPwd)
      '''    gUstat = IIf(gServerDb = "10.10.11.191", 1, 0) 'GetSetting(appname:="SERVICE", Section:="Stat", Key:="Top", Default:="0")
13       gUstat = 1
          
          'KONEKSi BARU ======================================================================
14        gServerDb = GetSetting("NIS-LOCAL", "Startup", "Server", "")
          'If gServerDb = "10.10.10.181" Then
          '    gServerDb = GetSetting("SVC-HO", "Startup", "Server", "")
          'End If

      '    gUser = "sa"
      '    gPwd = "nagUser = "koneksi"
      '    gPwd = "G4j4hM4da"smoco"

          'On Error GoTo Setup
          Dim rs As New ADODB.Recordset
15        If newEnkrip Then
16            gUserKoneksi = userKon '"koneksi"
17            gPasswordKonesi = pwdKon ' "G4j4hM4da"
18        Else
19            gUserKoneksi = "sa" '"koneksi"
20            gPasswordKonesi = "nasmoco~123" ' "Nasmoco~123"  '"G4j4hM4da"
21        End If
22       ' rs.Open "Select * from Koneksi2 WITH (NOLOCK)", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & gUserKoneksi & ";password=" & gPasswordKonesi & ";Initial Catalog=koneksi;server=" & gServerDb), adOpenStatic, adLockBatchOptimistic, adCmdText

      '    Dim oCr As New Crypt
      '    gUser = oCr.Crypt(rs.Fields(0).value, "M4d462", False)
      '    gPwd = oCr.Crypt(rs.Fields(1).value, "G4j4H", False)
          
23        If newEnkrip Then
24            gUser = Dekripsi(rs.Fields(0).value, KataKunci(3))
25            gPwd = Dekripsi(rs.Fields(1).value, KataKunci(5))
26        Else
27            gUser = "sa" 'DekripsiGamaLama(rs.Fields(0).value, KataKunciGamaLama(3))
28            gPwd = "nasmoco~123" '"Nasmoco~123"  'DekripsiGamaLama(rs.Fields(1).value, KataKunciGamaLama(5))
29        End If
          
30        Set rs = Nothing
      '  Call Check_Security
          'KONEKSi BARU======================================================================


          'gUstat = IIf(gServerDb = "10.10.11.191", 1, 0) 'GetSetting(appname:="SERVICE", Section:="Stat", Key:="Top", Default:="0")
          ''20181022
31        gUstat = IIf(gServerDb = "10.10.10.181", 1, 0) 'GetSetting(appname:="SERVICE", Section:="Stat", Key:="Top", Default:="0")
          '[gUstat = 1
          gAlamatDepo = "[10.10.10.52\ho]."
          gAlamatDepo = ""
32        NisDb2 = ""
          
33        allowIn = False
34        bookAktif = False
35        orderAktif = False
36        stallAktif = False
37        validAktif = False
38        subletAktif = False
39        subRetAktif = False
40        flagOk = False
41        gQueue = False
42        If str(Date) <> Format(Date, "dd/mm/yyyy") Or Format(1000.24, "#,##0.#0") <> "1.000,24" Then
43           MsgBox "Silahkan Setting komputer anda ke Setting Indonesian :" & vbNewLine & _
                    "Format Date (dd/mm/yyyy)" & vbNewLine & _
                    "Decimal Symbol (,)" & vbNewLine & _
                    "Digit Grouping (.)", vbInformation, "Logout"
44           Shell "control.exe intl.cpl", vbNormalFocus  'memanggil settingan regional
45           Exit Sub
46        End If
          
47        Call cekEXE

48        gKode = "SV"
49        gKodeRetur = "R"
          
      '.:: ======================================== .:: BEGIN ::. ========================================= ::.
      '.:: ======================================== 15 / 04 /2009 ========================================= ::.
      '.:: ======================================== .:: M R S ::. ========================================= ::.

50        Log.LogServ = gServerDb
51        Log.LogUser = gUser
52        Log.LogPass = gPwd
53        Log.LogNas = gPerush
54        Log.LogAlamat = gAlamat
          
          
          'Log.LogUser = "SA"
          'Log.LogPass = "nasmoco"
          
      '.:: ======================================== .:: M R S ::. ========================================= ::.
      '.:: ======================================== 15 / 04 /2009 ========================================= ::.
      '.:: ======================================== .:: E N D ::. ========================================= ::.
          
55            If Get_ConnectionProfil_Bengkel = False Then
56                GoTo Setup
57                Exit Sub
58            End If
              Dim xDate As Date
59            xDate = gTglServer
60            gTanggal = TglDDMMYYYY(Format(xDate, "DD-MM-YYYY"))
61            gTglDefault = "1900-01-01"
              
62            gDb = ""
63            If gUstat = 1 Then
64                gDb = "BENGKEL99"
65            Else
66                gUstat = 0
67                gDb = gDb2
68            End If
                  
              '20180323 'REV '
69            If Log.LogHOConn = 1 And Log.RPLHistori = 0 Then 'If Log.LogHOConn = 1 Then
70                ConHO = True
71                If Get_ConnectionProfil_Bengkel99 = False Then
72                    ConHO = False
73                    Log.LogHOConn = 0
74                    MsgBox "Ada Kesalahan Setting ke HO," & Chr(13) & "Cek VPN dan konektivitas ke HO" & Chr(13) & "(Lakukan Routing IP)", vbInformation + vbOKOnly, "NIS Service"
75                End If
76            ElseIf Log.RPLHistori = 0 Then
77                ConHO = False
78            Else
79                ConHO = True
80            End If
          
81            If DateDiff("d", xDate, gDbTanggal) > 0 Then
82                If Trim(gDbVersiLama) <> Trim(gDbVersi) Then
83                    If Trim(gDbVersiLama) = Trim(gVersi) Then
84                        MsgBox "Download NIS Service Terbaru : " & Trim(gDbVersi) & vbCr & vbCr & "Versi Lama : " & Trim(gDbVersiLama) & " akan di-LOCK pada Tanggal " & Format(gDbTanggal, "DD MMMM YYYY") & vbCr & vbCr & vbCr & "KURANG " & DateDiff("d", isnow, gDbTanggal) & " Hari.", vbCritical, "Update Versi Lama Anda (Versi Anda : " & gDbVersiLama & ")"
85                        allowIn = True
86                    ElseIf Trim(gDbVersiLama) = Trim(gVersi) Then
87                        allowIn = True
88                    ElseIf Trim(gDbVersi) = Trim(gVersi) Then
89                        allowIn = True
90                    End If
91                Else
92                    allowIn = True
93                End If
94            ElseIf DateDiff("d", xDate, gDbTanggal) <= 0 Then
95                If Trim(gDbVersi) = Trim(gVersi) Then
96                    allowIn = True
97                End If
98            End If
99            SQLCon = SQLCONBengkel
              
      '        Dim rsProfilBengkel As New ADODB.Recordset
      '        Set rsProfilBengkel = getDataProfilBengkel("SELECT * FROM t_profil WITH (NOLOCK)")
      '        If Not rsProfilBengkel.EOF Then
      '            'mengambil setting spk
      '
      '            Log.LogSPKDb = rsProfilBengkel!FS_SPK_DB
      '            Log.LogSPKKomp = rsProfilBengkel!FS_SPK_KOMP
      '            Log.LogSPKPass = Dekripsi(rsProfilBengkel!FS_SPK_PASS, KataKunci(5))
      '            Log.LogSPKUser = Dekripsi(rsProfilBengkel!FS_SPK_USER, KataKunci(3))
      '
      '            'mengambil setting faktur
      '            Log.logFakturKomp = rsProfilBengkel!FS_NIS_KOMP
      '            Log.LogFakturPass = Dekripsi(rsProfilBengkel!FS_NIS_PASS, KataKunci(5))
      '            Log.LogFakturUser = Dekripsi(rsProfilBengkel!FS_NIS_USER, KataKunci(3))
      '
      '        End If
              
100           If gUstat <> 1 Then
101               If Get_ConnectionProfil_SingleID = False Then
102                   GoTo Setup
103                   Exit Sub
104               End If
105               If Get_ConnectionProfil_SPK = False Then
106                   GoTo Setup
107                   Exit Sub
108               End If
109               Log.LogQueue = "Queue" & Trim(Right(gDb, 2))
110               If Get_ConnectionProfil_Queue = False Then
111                   GoTo Setup
112                   Exit Sub
113               End If
114               Log.LogFaktur = "Faktur" & Trim(Right(gDb, 2))
115               If Get_ConnectionProfil_FP = False Then
116                   GoTo Setup
117                   Exit Sub
118               End If
119               Log.LogSMS = "SMS" & Trim(Right(gDb, 2))
120               If Get_ConnectionProfil_SMS = False Then
121                   GoTo Setup
122                   Exit Sub
123               End If
124               Log.LogACC = NisDb
125               If Get_ConnectionProfil_ACC = False Then
126                   GoTo Setup
127                   Exit Sub
128               End If
129           End If
130           goHo = False
              
131           Call SetTimerNow

132           If allowIn Then
133               If gDb <> "" Then
134                   If Get_Connection(gDb, gUser, gPwd, gServerDb) Then
135                       Set MyArray = New XArrayDB
136                       Call dropTempTable
137                       If gUstat = 0 Then
                              ' update 27 Jan 2017
                              ' terkait dgn mekanisme CSL ada status LS dengan kondisi DTLFU seperti belum tersambung maka ditambahkan
                              ' filter penghapusan selain respon DTLFU LS
138                           Call objO.DeleteData("DELETE DTLFU WHERE SAMBUNG = 'N' AND COUNTER < 1 and respon not in ('LS')")
                              'Call updateGLtrn
139                           Call updateHarian(TglNul(xDate))
140                       ElseIf gUstat = 1 Then
141                           gStatKoneksi = True
142                           If Not Get_ConnectionHO_New Then
143                               MsgBox "Gagal Koneksi ke HO"
144                               End
145                           End If
146                       End If
147                       allowMove = True
                         ''frmReportLS.Show
148                       frmSplash.show
149                   End If
150               Else
151                   allowIn = False
152               End If
153           Else
154               MsgBox "Update versi terbaru " & Trim(gDbVersi) & " !! Lakukan copy NIS !!"
155               End
156           End If
          
157       Exit Sub
Setup:
158       If err.Number = -2147467259 Then
159           MsgBox "Anda belum Log in, Log off dahulu !", vbCritical + vbOKOnly
160           MsgBox err.Description
161       Else
162           MsgBox "Server atau Database tidak terhubung !"
              'MsgBox err.Description
              
163       End If
          'frmSetup.Show
164       allowIn = False
165       Exit Sub
End Sub

Public Function Get_ConnectionProfil_Bengkel99() As Boolean
    On Error GoTo Vb_Error99
    'GoTo Vb_Error99
    Set mConnHO = New ADODB.Connection
    mConnHO.CursorLocation = adUseClient
    SQLConBengkel99 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogHOUser & ";password=" & Log.LogHOPass & ";Initial Catalog=" & Log.LogHODb & " ;server=" & Log.LogHOKomp & ";Connect Timeout=10"
    mConnHO.ConnectionString = SQLConBengkel99
    mConnHO.ConnectionTimeout = 10
    mConnHO.Open
    Get_ConnectionProfil_Bengkel99 = True
    If mConnHO.State = adStateOpen Then mConnHO.Close
    ConHO = True
    Exit Function
Vb_Error99:
    If mConnHO.State = adStateOpen Then mConnHO.Close
    Get_ConnectionProfil_Bengkel99 = False
    ConHO = False
End Function

Public Function getDataProfilBengkel(SQL As String) As ADODB.Recordset
    On Error GoTo errHandler
    'Dim mconn As New ADODB.Recordset
    'If mconn.State = adStateClosed Then mconn.Open (SQLCon)
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open SQL, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set getDataProfilBengkel = rs
    'If mconn.State <> adStateClosed Then mconn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function Get_ConnectionProfil_Bengkel() As Boolean
On Error GoTo vb_error
    Set DBConBengkel = New ADODB.Connection
    DBConBengkel.CursorLocation = adUseClient
    DBConBengkel.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog=ProfilBengkel ;server=" & Log.LogServ
    DBConBengkel.Open
    Get_ConnectionProfil_Bengkel = True
    If DBConBengkel.State = adStateOpen Then DBConBengkel.Close
    SQLConProfilBK = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog=ProfilBengkel ;server=" & Log.LogServ
    Call GetProfileBengkel
    Call GetBengkelHO
    Exit Function
vb_error:
    Get_ConnectionProfil_Bengkel = False
    MsgBox "Gagal Koneksi Ke database ProfilBengkel, Proses Dibatalkan !", vbExclamation
    'Shell App.Path & "\Setting NIS.exe", vbNormalFocus
    End
End Function

Public Sub GetProfileBengkel()
On Error GoTo vb_error
    Dim RsBengkel As New ADODB.Recordset
    RsBengkel.CursorLocation = adUseClient
    If gUstat = 0 Then
        RsBengkel.Open "SELECT * FROM profil WITH (NOLOCK)", SQLConProfilBK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Else
        'RsBengkel.Open "SELECT * FROM profil1 WITH (NOLOCK)", SQLConProfilBK, adOpenStatic, adLockBatchOptimistic, adCmdText
        RsBengkel.Open "SELECT * FROM profil WITH (NOLOCK)", SQLConProfilBK, adOpenStatic, adLockBatchOptimistic, adCmdText
    End If
    Set RsBengkel.ActiveConnection = Nothing
    
    If RsBengkel.RecordCount > 0 Then
       Log.LogBengkel = Trim(Nul(RsBengkel!DBASECABANG))    '" & DB_bengkel & "
         ''xx''DB_Bengkel = "Bengkel09"
            gDb2 = Trim(RsBengkel("DBASECABANG"))
            gDbVersi = Trim(RsBengkel("VERSION"))
            gDbVersiLama = Trim(RsBengkel("VERSIONLAMA"))
            gDbTanggal = Trim(RsBengkel("TGLAKTIF"))
            allowSaveLog = IIf(Trim(RsBengkel("LOGMENU")) = 0, False, True)
            'If gDb2 <> "BENGKEL99" Then
            OnTopSV = CInt(RsBengkel("TopAR"))
            OnNewFP = CInt(RsBengkel("NewFP"))
            'End if
            
            If gUstat = 0 Then
                Log.isFill = CInt(RsBengkel("Fill"))
                Log.AppRate = Trim(RsBengkel("AppRate"))
                Log.startAppRate = TglNul(RsBengkel("startAppRate"))
                                            
                'TAMBAHAN FLAG PENYERAGAMAN AKUN / TIDAK
                Log.isAkun = CInt(RsBengkel("Akun"))
                Log.expPwd = CInt(RsBengkel("expPwd"))
                Log.isEmStall = CInt(RsBengkel("ISEMSTALL"))
                               
                
                'TAMBAHAN FLAG DMS NON DSS
                If dmsNonDss = True Then
                    Log.dmsNonDss = CInt(RsBengkel("DMSNONDSS"))
                    Log.barcode = CInt(RsBengkel("BARCODE"))
                End If
                Log.isDSS = CInt(RsBengkel("DMSDSS"))
                Log.smsBilling = CInt(RsBengkel("SMSBILLING"))
                Log.isWashing = CInt(RsBengkel("isWashing"))
                Log.helpJpcb = IIf(CInt(RsBengkel("refJPCB")) > 0, 1, 0)
                Log.newJpcb = CInt(RsBengkel("newJpcb"))
                Log.validasiKabeng = CInt(RsBengkel!validasiKabeng)
                Log.FIBeforeSublet = CInt(RsBengkel!FIBeforeSublet)
                Log.StartAffPpn = TglNul(RsBengkel!StartAffPpn)
                Log.lockWashing = CInt(RsBengkel!lockWashing)
                Log.FIR = CInt(RsBengkel!FIR)
                Log.followUp = CInt(RsBengkel!followUp)
                Log.OTD = CInt(RsBengkel!OTD)
                Log.detWAC = CInt(RsBengkel!detWAC)
                Log.lockWAC = CInt(RsBengkel!lockWAC)
                Log.onExport = CInt(RsBengkel!onExport)
                Log.discKacab = CInt(RsBengkel!discKacab)
                Log.startDiscKacab = TglNul(RsBengkel!tglaktifdisckacab)
                
                ' update 31 jan 2017, penambahan field utk penanda proses update yg dilakukan hanya 1x dalam 1 hari
                'Log.tglUpdtEvd = TglNul(RsBengkel("tglUpdtEvd"))
                
                '' 20190701REV1
                Log.newLockSSC = CInt(RsBengkel!newLockSSC)
            End If
          'adit disableexport 20161220 -begin
          Log.PdfOnly = CInt(RsBengkel!onExport) ' CInt(RsBengkel!PdfOnly)
          'adit disableexport 20161220 -end
          Log.notifInvNCM = CInt(RsBengkel!notifInvNCM)
          Log.notifInvNonNCM = CInt(RsBengkel!notifInvNonNCM)
          Log.repoHarianMech = CInt(RsBengkel!repoHarianMech)
          
          Log.loyalityCustomer = CInt(RsBengkel!loyalityCustomer)
          Log.ketNasApp = CStr(RsBengkel!ketnasmocoapp)
          
          Log.RPLHistori = CInt(RsBengkel!RPLHistori)
          Log.mandatoryKTP = CInt(RsBengkel!mandatoryKTP)
          Log.logRetention = CInt(RsBengkel!Retention)
          Log.CR7 = CInt(RsBengkel!CR7)
          If gUstat = 0 Then
            Log.fakturMaterai = CInt(RsBengkel!fakturMaterai)
            Log.newCetak = CInt(RsBengkel!newcetaksimplify)
            Log.replikasiWacHO = CInt(RsBengkel!replikasiWacHO)
            Log.lockSCC = CInt(RsBengkel!lockSCC)
            
            '' 20180724DEV01 ============================================
            Log.NoKtpKetFP = CInt(RsBengkel!NoKtpKetFP)
            Log.NoKtpNonPerorangan = CInt(RsBengkel!NoKtpNonPerorangan)
            '' ==========================================================
            '' 20180927DEV  ' SCORING SINGLE ID
            Log.singleIDScoring = CInt(RsBengkel!singleIDScoring)
            
            
            '' 20190206REV1 | REV | REVISI  | PENAMBAHAN FLAG UNTUK PENGECEKAN NO KTP DAN NO HP
            Log.lockKTPNoHPScoring = "0" 'CInt(RsBengkel!lockKTPNoHPScoring)
            
            '' 20190816DEV1
            Log.hansaRenkei = CInt(RsBengkel!hansaRenkei)
            '' 20190625DEV1
            Log.lockFreeSVCHO = CInt(RsBengkel!freesvcho)
            Log.tglFreeSVCHO = CDate(Format(RsBengkel!tglFreeSVCHO, "dd/mm/yyyy"))
            Log.tglRSA = CDate(Format(RsBengkel!tglRSA, "dd/mm/yyyy"))
            Log.addMaterai = CInt(RsBengkel!addMaterai)
            Log.newOhbeya = CInt(RsBengkel!newOhbeya)
            
            '' 20200408DEV
            Log.sentralisasiStars = CInt(RsBengkel!sentralisasiStars)
            
            '' 20200512DEV
            Log.disableExportMra = CInt(RsBengkel!disableExportMra)
            ''20200709DEV
            Log.monitoringScoring = CInt(RsBengkel!monitoringScoring)
            
            Log.hansaRenkeiBP = CInt(RsBengkel!hansaRenkeiBP)
            
            
            '' 20201218DEV
            Log.materai2021 = CInt(RsBengkel!materai2021)
                        
            '' 20201222DEV
            Log.NHS = CInt(RsBengkel!NHS)
            
            '' 20210520DEV
            Log.simpanRepl = CInt(RsBengkel!simpanRepl)
            Log.flagOffBookingLocal = CInt(RsBengkel!flagOffBookingLocal)
            
            '' 20211125DEV
            Log.FlagSalesHO = CInt(RsBengkel!FlagSalesHO)
            
            '' 20220210DEV
            Log.hideService = CInt(RsBengkel!hideService)
            
            ''20230428DEV
            Log.tcm = CInt(RsBengkel!tcm)
            Log.tcmMaxHari = CInt(RsBengkel!tcmMaxHari)
            ''20240119DEV
            Log.DIDX = CInt(RsBengkel!DIDX)
            'Log.kdDealerDidx = RsBengkel!kdDealerDidx
            Log.DiDxP1 = "0" 'RsBengkel!DiDxP1
            Log.dpplain = "1" ' RsBengkel!dpplain
          End If
          
          '' 20181018DEV1 | DEV | DEVELOPMENT  | PENAMBAHAN FLAG UNTUK STANDARISASI SUBLET
          Log.standarisasiSublet = CInt(RsBengkel!standarisasiSublet)
          
          '' 20191115DEV1 | penambahan flag
          Log.rsA = CInt(RsBengkel!rsA)
          
          Log.multipleMarketing = CInt(RsBengkel!multipleMarketing)
          Log.tglmultipleMarketing = CDate(Format(RsBengkel!tglmultipleMarketing, "dd/mm/yyyy"))
          
          '' 20200303DEV
          Log.vinRetention = CInt(RsBengkel!vinByVinRetention)
          
          '' 20201124DEV
          Log.vin2vinRev1 = CInt(RsBengkel!vin2vinRev1)
          
          '' 20200416DEV
          Log.freepart = CInt(RsBengkel!freepart)
          
    End If
    RsBengkel.Close
    If DBConBengkel.State = adStateOpen Then DBConBengkel.Close
    Set RsBengkel = Nothing
    SQLCONBengkel = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog=" & Log.LogBengkel & " ;server=" & Log.LogServ
    adoCon = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog=" & Log.LogBengkel & " ;server=" & Log.LogServ
    SQLConAAS = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog=BENGKEL ;server=" & Log.LogServ
    Set DBConBengkel = Nothing
 Exit Sub
vb_error:
  MsgBox ErrMessage(Erl, err.Number, "GetProfileBengkel")
  MsgBox "Error Setting !!!"
  End
End Sub

Public Sub GetBengkelHO()
         
1         On Error GoTo vb_error
          Dim RsBengkel As New ADODB.Recordset
2         RsBengkel.CursorLocation = adUseClient
3         If gUstat = 1 Or gServerDb = "10.10.10.181" Then 'GetSetting("SVC-HO", "Startup", "Server", "") Then
4             RsBengkel.Open "select P.*, '' pesan1 from T_PROFIL p WITH (NOLOCK)", SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
5         Else
6             RsBengkel.Open "select P.*, isnull(T.Pesan1,'') pesan1 from T_PROFIL p WITH (NOLOCK) left join M_Template T WITH (NOLOCK) ON P.SMSBILLING = T.TypeSMS", SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
7         End If
          
8         If RsBengkel.RecordCount > 0 Then
9             If newEnkrip Then
10                Log.LogHOPass = Dekripsi(Trim(RsBengkel("FS_HO_PASS")), KataKunci(5))
11                Log.LogHOUser = Dekripsi(Trim(RsBengkel("FS_HO_USER")), KataKunci(3))
                  
12                Log.LogHOPassSPK = Dekripsi(Trim(RsBengkel("FS_SPKHO_PASS")), KataKunci(5))
13                Log.LogHOUserSPK = Dekripsi(Trim(RsBengkel("FS_SPKHO_USER")), KataKunci(3))
                  
14                Log.LogAccUser = Dekripsi(Trim(RsBengkel("FS_NIS_USER")), KataKunci(3))
15                Log.LogAccPass = Dekripsi(Trim(RsBengkel("FS_NIS_PASS")), KataKunci(5))
                  
16                Log.LogSPKPass = Dekripsi(RsBengkel("FS_SPK_PASS"), KataKunci(5))
17                Log.LogSPKUser = Dekripsi(RsBengkel("FS_SPK_USER"), KataKunci(3))
                  
18                Log.LogFakturPass = Dekripsi(RsBengkel("FS_NIS_PASS"), KataKunci(5))
19                Log.LogFakturUser = Dekripsi(RsBengkel("FS_NIS_USER"), KataKunci(3))
                  
                  
20                Log.LogSIDPass = Dekripsi(RsBengkel("FS_SID_PASS"), KataKunci(5))
21                Log.LogSIDUser = Dekripsi(RsBengkel("FS_SID_USER"), KataKunci(3))
22            Else
23                Log.LogHOPass = gPwd '"N@sm0c0W0r!d"  ' Trim(RsBengkel("FS_HO_PASS"))
24                Log.LogHOUser = gUser '"Nis" ' Trim(RsBengkel("FS_HO_USER"))
                  
25                Log.LogHOPassSPK = gPwd '"N@sm0c0W0r!d" 'Trim(RsBengkel("FS_SPKHO_PASS"))
26                Log.LogHOUserSPK = gUser '"Nis" ' Trim(RsBengkel("FS_SPKHO_USER"))
                  
27                Log.LogAccPass = gPwd 'Trim(RsBengkel("FS_NIS_PASS"))
28                Log.LogAccUser = gUser ' Trim(RsBengkel("FS_NIS_USER"))
                  
29                Log.LogSPKPass = gPwd ' RsBengkel("FS_SPK_PASS")
30                Log.LogSPKUser = gUser ' RsBengkel("FS_SPK_USER")
                  
31                Log.LogFakturPass = gPwd ' RsBengkel("FS_NIS_PASS")
32                Log.LogFakturUser = gUser 'RsBengkel("FS_NIS_USER")
                  
                  
33                Log.LogSIDPass = gPwd 'RsBengkel("FS_SID_PASS")
34                Log.LogSIDUser = gUser 'RsBengkel("FS_SID_USER")
35            End If

Dim RsNewEnc2018 As New ADODB.Recordset
    If newEnkrip Then
  RsNewEnc2018.CursorLocation = adUseClient
  RsNewEnc2018.Open "select * from encT_PROFIL WITH (NOLOCK)", SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
  If RsBengkel.RecordCount > 0 And newEnkrip Then
      Log.LogHOUser = Dekripsi(RsNewEnc2018("FS_HO_USER"), KataKunci(3))
      Log.LogHOPass = Dekripsi(RsNewEnc2018("FS_HO_PASS"), KataKunci(5))
      Log.LogHOUserSPK = Dekripsi(RsNewEnc2018("FS_SPKHO_USER"), KataKunci(3))
      Log.LogHOPassSPK = Dekripsi(RsNewEnc2018("FS_SPKHO_PASS"), KataKunci(5))
      Log.LogAccUser = Dekripsi(RsNewEnc2018("FS_NIS_USER"), KataKunci(3))
      Log.LogAccPass = Dekripsi(RsNewEnc2018("FS_NIS_PASS"), KataKunci(5))
      Log.LogSPKUser = Dekripsi(RsNewEnc2018("FS_SPK_USER"), KataKunci(3))
      Log.LogSPKPass = Dekripsi(RsNewEnc2018("FS_SPK_PASS"), KataKunci(5))
      Log.LogFakturUser = Dekripsi(RsNewEnc2018("FS_NIS_USER"), KataKunci(3))
      Log.LogFakturPass = Dekripsi(RsNewEnc2018("FS_NIS_PASS"), KataKunci(5))
      Log.LogSIDUser = Dekripsi(RsNewEnc2018("FS_SID_USER"), KataKunci(3))
      Log.LogSIDPass = Dekripsi(RsNewEnc2018("FS_SID_PASS"), KataKunci(5))
  End If
  End If

36            Log.LogHODb = Trim(RsBengkel("FS_HO_DB"))
37            Log.LogHOKomp = Trim(RsBengkel("FS_HO_KOMP")) '"10.10.10.181"
38            Log.LogHOConn = IIf(Trim(RsBengkel("KONEKSIHO")) = "FAILED", 0, 1)
39            Log.LogNamaNas = Trim(RsBengkel("FS_NM_Dealer"))
40            Log.LockAR = IIf(Trim(RsBengkel("LOCKAR")) = 1, True, False)
              
41            Log.LogHODbSPK = Trim(RsBengkel("FS_SPKHO_DB"))
42            Log.LogHOKompSPK = Trim(RsBengkel("FS_SPKHO_KOMP"))
43            Log.LogKbgName = Trim(RsBengkel("FS_KA_BENGKEL"))
              
44            Log.LogAccServ = Trim(RsBengkel("FS_NIS_KOMP"))
45            Log.LogSPKDb = RsBengkel("FS_SPK_DB")
46            Log.LogSPKKomp = RsBengkel("FS_SPK_KOMP")
47            Log.LogSIDServ = RsBengkel("FS_SID_KOMP")
48            Log.LogSIDDb = RsBengkel("FS_SID_DB")
              
              'mengambil setting faktur
49            Log.logFakturKomp = RsBengkel("FS_NIS_KOMP")
50            If gUstat = 0 Then
51                Log.estInv = CInt(RsBengkel("ESTINV"))
52                Log.KodeDealer = Trim(RsBengkel("FS_KD_DEALER"))
                  'TAMBAHAN FLAG DMS NON DSS
53                If dmsNonDss = True Then
54                    Log.kmwarranty = CDbl(RsBengkel("KMWARRANTY"))
55                    Log.warranty = CInt(RsBengkel("WARRANTY"))
56                    Log.persentaseDP = CInt(RsBengkel("PERSENTASEDP"))
57                    Log.plafonDP = CDbl(RsBengkel("PLAFONDP"))
58                    Log.KodeDealer = Trim(RsBengkel("FS_KD_DEALER"))
59                End If
                  
60                Log.formatSmsBilling = Trim(RsBengkel("PESAN1"))
                  
                  '' 20200408DEV
61                Log.kodeCabang = NulSID(RsBengkel!kodeCabang)
62            End If

63            If Log.FIBeforeSublet = 1 Then
64                Log.prioritas = Trim(RsBengkel!prioritas)
65            End If
              
66            If Log.followUp = 1 Then
67                Log.prosentaseFollow = CInt(RsBengkel!persenfollow)
68            End If
              
69            If Log.discKacab = 1 Then
70                Log.discJasa = CInt(RsBengkel!discmax)
71                Log.discMaxJasa = CInt(RsBengkel!DISCMAXKBG)
72                Log.discPart = CInt(RsBengkel!discPartMax)
73                Log.discPartMax = CInt(RsBengkel!discpartkacabmax)
                  
74            End If
              
              ''========================================================
              '' 20180524DEV01'INSPECTION
              ''========================================================
75            If Log.CR7 = 1 Then
76                Log.FTPCR7Host = Trim(RsBengkel!FTPCR7Host)
77                Log.FTPCR7User = Trim(RsBengkel!FTPCR7User)
78                Log.FTPCR7Pass = Trim(RsBengkel!ftpcr7password)
79            End If
              
80            If Log.standarisasiSublet = 1 And gUstat = 0 Then
81                Log.discMaxSublet = CInt(RsBengkel!discMaxSublet)
82            End If
              
              
83            SQLConHO = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogHOUser & ";password=" & Log.LogHOPass & ";Initial Catalog=" & Log.LogHODb & ";server=" & Log.LogHOKomp
84            SQLConSPKHO = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogHOUserSPK & ";password=" & Log.LogHOPassSPK & ";Initial Catalog=" & Log.LogHODbSPK & ";server=" & Log.LogHOKompSPK
              ''If gServerDb <> "10.10.11.191" Then OnDisKBG = Val(RsBengkel("DiscMaxKBG"))
              ''20181022
85            If gServerDb <> "10.10.10.181" Then OnDisKBG = Val(RsBengkel("DiscMaxKBG"))
86        End If
87        Set RsBengkel.ActiveConnection = Nothing
88        Set RsBengkel = Nothing
89     Exit Sub
vb_error:
90      MsgBox ErrMessage(Erl, err.Number, "GetProfileBengkel")
End Sub

Public Function Get_ConnectionProfil_SPK() As Boolean
 On Error GoTo vb_error
 Set DBConSPK = New ADODB.Connection
 DBConSPK.CursorLocation = adUseClient
 DBConSPK.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogSPKUser & ";password=" & Log.LogSPKPass & ";Initial Catalog=ProfilSPK ;server=" & Log.LogSPKKomp
 DBConSPK.Open
 Get_ConnectionProfil_SPK = True
 If DBConSPK.State = adStateOpen Then DBConSPK.Close
 SQLConProfilSP = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogSPKUser & ";password=" & Log.LogSPKPass & ";Initial Catalog=ProfilSPK ;server=" & Log.LogSPKKomp
 Call GetProfileSPK
 Exit Function
 
vb_error:
    Get_ConnectionProfil_SPK = False
    MsgBox "Gagal Koneksi Ke database ProfilSPK, Proses Dibatalkan !", vbExclamation
    Shell App.Path & "\Setting NIS.exe", vbNormalFocus
    End
End Function

Public Sub GetProfileSPK()
On Error GoTo vb_error
    Dim rsSPK As New ADODB.Recordset
    rsSPK.CursorLocation = adUseClient
    rsSPK.Open "SELECT KodeCabang FROM profil WITH (NOLOCK)", SQLConProfilSP, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rsSPK.ActiveConnection = Nothing
    If rsSPK.RecordCount > 0 Then
        Log.LogSPK = "SPK" + Trim(rsSPK!kodeCabang)
    End If
    rsSPK.Close
    If DBConSPK.State = adStateOpen Then DBConSPK.Close
    Set rsSPK = Nothing
    SQLConSPK = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogSPKUser & ";password=" & Log.LogSPKPass & ";Initial Catalog=" & Log.LogSPK & " ;server=" & Log.LogSPKKomp
    Set DBConSPK = Nothing
 Exit Sub
vb_error:
  MsgBox ErrMessage(Erl, err.Number, "main.SetProfileAcc")
End Sub

Public Function Get_ConnectionProfil_Queue() As Boolean
 On Error GoTo vb_error
 Set DBConQueue = New ADODB.Connection
 DBConQueue.CursorLocation = adUseClient
 DBConQueue.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog= " & Log.LogQueue & " ;server=" & Log.LogServ
 DBConQueue.Open
 Get_ConnectionProfil_Queue = True
 If DBConQueue.State = adStateOpen Then DBConQueue.Close
 SQLConQueue = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog= " & Log.LogQueue & " ;server=" & Log.LogServ
 Exit Function
vb_error:
    Get_ConnectionProfil_Queue = False
    MsgBox "Gagal Koneksi Ke database Queue, Proses Dibatalkan !", vbExclamation
    Shell App.Path & "\Setting NIS.exe", vbNormalFocus
    End
End Function

Public Function Get_ConnectionProfil_FP() As Boolean
 On Error GoTo vb_error
    Set DBConFP = New ADODB.Connection
    DBConFP.CursorLocation = adUseClient
    DBConFP.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogFakturUser & ";password=" & Log.LogFakturPass & ";Initial Catalog= " & Log.LogFaktur & " ;server=" & Log.logFakturKomp
    DBConFP.Open
    Get_ConnectionProfil_FP = True
    If DBConFP.State = adStateOpen Then DBConFP.Close
    SQLConFP = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogFakturUser & ";password=" & Log.LogFakturPass & ";Initial Catalog= " & Log.LogFaktur & " ;server=" & Log.logFakturKomp
 Exit Function
vb_error:
    Get_ConnectionProfil_FP = False
    MsgBox "Gagal Koneksi Ke database Faktur, Proses Dibatalkan !", vbExclamation
    Shell App.Path & "\Setting NIS.exe", vbNormalFocus
    End
End Function

Public Function Get_ConnectionProfil_SingleID() As Boolean
 On Error GoTo vb_error
 Set DBConSingleID = New ADODB.Connection
 DBConSingleID.CursorLocation = adUseClient
 DBConSingleID.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogSIDUser & ";password=" & Log.LogSIDPass & ";Initial Catalog= " & Log.LogSIDDb & " ;server=" & Log.LogSIDServ
 DBConSingleID.Open
 Get_ConnectionProfil_SingleID = True
 If DBConSingleID.State = adStateOpen Then DBConSingleID.Close
 SQLConSID = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogSIDUser & ";password=" & Log.LogSIDPass & ";Initial Catalog= " & Log.LogSIDDb & " ;server=" & Log.LogSIDServ
 Call GetProfileScoring
 Exit Function
vb_error:
    Get_ConnectionProfil_SingleID = False
    MsgBox "Gagal Koneksi Ke database Single ID, Proses Dibatalkan !", vbExclamation
    Shell App.Path & "\Setting NIS.exe", vbNormalFocus
    End
End Function

Public Function Get_ConnectionProfil_SMS() As Boolean
 On Error GoTo vb_error
 Set DBConSMS = New ADODB.Connection
 DBConSMS.CursorLocation = adUseClient
 DBConSMS.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog= " & Log.LogSMS & " ;server=" & Log.LogServ
 DBConSMS.Open
 Get_ConnectionProfil_SMS = True
 If DBConSMS.State = adStateOpen Then DBConSMS.Close
 SQLConSMS = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog= " & Log.LogSMS & " ;server=" & Log.LogServ
 Exit Function
vb_error:
    Get_ConnectionProfil_SMS = False
    MsgBox "Gagal Koneksi Ke database SMS, Proses Dibatalkan !", vbExclamation
    Shell App.Path & "\Setting NIS.exe", vbNormalFocus
    End
End Function

Public Function Get_ConnectionProfil_ACC() As Boolean
On Error GoTo vb_error
    Set DBConACC = New ADODB.Connection
    DBConACC.CursorLocation = adUseClient
    DBConACC.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogAccUser & ";password=" & Log.LogAccPass & ";Initial Catalog=" & Log.LogACC & " ;server=" & Log.LogAccServ
    DBConACC.Open
    SQLConACC = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogAccUser & ";password=" & Log.LogAccPass & ";Initial Catalog=" & Log.LogACC & " ;server=" & Log.LogAccServ
    Call GetProfileACC
    Get_ConnectionProfil_ACC = True
    If DBConACC.State = adStateOpen Then DBConACC.Close
    Exit Function
    
vb_error:
    Get_ConnectionProfil_ACC = False
    MsgBox "Gagal Koneksi Ke database ProfilAcc, Proses Dibatalkan !", vbExclamation
    Shell App.Path & "\Setting NIS.exe", vbNormalFocus
    End
End Function

'' ====================================================================
'' 20180907 PENAMBAHAN PENGAMBILAN PROFIL ACC
'' ====================================================================

Public Sub GetProfileACC()
On Error GoTo vb_error
    Dim RsBengkel As New ADODB.Recordset
    RsBengkel.CursorLocation = adUseClient
    RsBengkel.Open "select NamaProfil,AlamatProfil1,KotaProfil,TeleponProfil,FaxProfil,NPWP,SKPengukuhan,TglPengukuhan from ProfilACC..Profil WITH (NOLOCK)", SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
     
    Set RsBengkel.ActiveConnection = Nothing
    
    If RsBengkel.RecordCount > 0 Then
        With profilACC
            .gNamaProfil = NulSID(RsBengkel!namaprofil)
            .AlamatProfil1 = NulSID(RsBengkel!AlamatProfil1)
            .KotaProfil = NulSID(RsBengkel!KotaProfil)
            .TeleponProfil = NulSID(RsBengkel!TeleponProfil)
            .FaxProfil = NulSID(RsBengkel!FaxProfil)
            .npwp = NulSID(RsBengkel!npwp)
            .SKPengukuhan = NulSID(RsBengkel!SKPengukuhan)
            .TglPengukuhan = Format(NulSID(RsBengkel!TglPengukuhan), "dd/mm/yyyy")
        End With
    End If
    RsBengkel.Close
    Set RsBengkel = Nothing
    RsBengkel.CursorLocation = adUseClient
    RsBengkel.Open "select * from ProfilACC..settingppn WITH (NOLOCK) where (getdate() between tglawal and tglakhir) and aktif = 1", SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
     
    Set RsBengkel.ActiveConnection = Nothing
    
    If RsBengkel.RecordCount > 0 Then
        With profilACC
            .nilaippn = NulSID(RsBengkel!nilaippn)
            .nilaitotal = NulSID(RsBengkel!nilaippn) + 100
            .nilaidpp = .nilaitotal - .nilaippn
            
        End With
    Else
        MsgBox "Setting PPn tidak ditemukan" & vbCrLf & "Silahkan hubungi IT NRM", vbInformation
        End
    End If
    RsBengkel.Close
    Set RsBengkel = Nothing
 Exit Sub
vb_error:
  MsgBox ErrMessage(Erl, err.Number, "GetProfileBengkel")
  MsgBox "Error Setting !!!"
  End
End Sub
'' ====================================================================

'.:: ======================================== .:: BEGIN ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: M R S ::. ========================================= ::.

Public Function Get_Connection(DBname As String, UsrName As String, Pass As String, SerName As String) As Boolean
     On Error GoTo vb_error
     
     Get_Connection = True
     If oConn.State Then oConn.Close
     oConn.CursorLocation = adUseClient
     oConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
     LogonStr = oConn.ConnectionString
     adoCon = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
     oConn.ConnectionTimeout = 15
     oConn.Open
     Exit Function
    
vb_error:
    Get_Connection = False
    MsgBox ErrMessage(Erl, err.Number, "Function : Get_Connection"), vbCritical, "Open Connection"
End Function

      Private Sub SetTimerNow()
         Dim mConn As New ADODB.Connection
         Dim rs As New ADODB.Recordset
         mConn.Open (SQLCon)
         rs.CursorLocation = adUseClient
         rs.Open "Select GetDate()", mConn, adOpenStatic, adLockBatchOptimistic, adCmdText
         isnow = rs.Fields(0).value
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
            If Key = 13 Then Sendkeys "{tab}"
      End Sub
      
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

      Public Function HargaJualPart(item_part As String) As Double
        Dim rsHarga As New ADODB.Recordset
        If rsHarga.State <> adStateClosed Then rsHarga.Close
        Set rsHarga = support.partPrice(Trim(item_part))
        If Not rsHarga.EOF Then
            HargaJualPart = CDbl(rsHarga("HARGAJUALA"))
        End If
        If rsHarga.State <> adStateClosed Then rsHarga.Close
        Set rsHarga = Nothing
      End Function

'.:: ======================================== .:: M R S ::. ========================================= ::.
'.:: ======================================== 15 / 04 /2009 ========================================= ::.
'.:: ======================================== .:: E N D ::. ========================================= ::.



Public Function Get_ConnectionLog(DBname As String, _
                                    UsrName As String, _
                                    Pass As String, _
                                    SerName As String) As Boolean
1      On Error GoTo vb_error
3      Get_ConnectionLog = True
4      If oConnLog.State Then oConnLog.Close
5      oConnLog.CursorLocation = adUseClient
6      oConnLog.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       oConnLog.ConnectionTimeout = 15
8      oConnLog.Open
9      oConnLog.Close
10      Exit Function

vb_error:
11    Get_ConnectionLog = False
End Function

Public Function Get_ConnectionProfil(DBname As String, _
                                     UsrName As String, _
                                     Pass As String, _
                                     SerName As String) As Boolean
                                   
1       On Error GoTo vb_error
2       Get_ConnectionProfil = True
3       If oConnProfil.State Then oConnProfil.Close
4       oConnProfil.CursorLocation = adUseClient
5       oConnProfil.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
        SQLCONBengkelProfil = oConnProfil.ConnectionString
6       oConnProfil.ConnectionTimeout = 15
7       oConnProfil.Open
        oConnProfil.Close
8       Exit Function

vb_error:
9    Get_ConnectionProfil = False
10    MsgBox ErrMessage(Erl, err.Number, "Function : Get_ConnectionProfil"), vbCritical, "Open Connection"
End Function

Public Function Get_ConnectionHO(DBname As String, _
                                 UsrName As String, _
                                 Pass As String, _
                                 SerName As String) As Boolean
1      On Error GoTo vb_error
3          Get_ConnectionHO = True
4          If oConnHO.State Then oConnHO.Close
5          oConnHO.CursorLocation = adUseClient
6          oConnHO.ConnectionString = SQLConHO
           SQLCONBengkelHO = SQLConHO
7          adoConHO = SQLConHO
           
            gStatKoneksi = True
            If gStatKoneksi Then
                 oConnHO.ConnectionTimeout = 15
8                oConnHO.Open
                 oConnHO.Close
            Else
                 oConnHO.ConnectionTimeout = 2
                 GoTo vb_error
            End If
9      Exit Function

vb_error:
10    Get_ConnectionHO = False
      gStatKoneksi = False
      
      Exit Function
End Function

Public Function Get_ConnectionHO_New() As Boolean
1      On Error GoTo vb_error
3          Get_ConnectionHO_New = True
4          If oConnHO.State Then oConnHO.Close
5          oConnHO.CursorLocation = adUseClient
6          oConnHO.ConnectionString = SQLConHO
            If gStatKoneksi Then
                 oConnHO.ConnectionTimeout = 15
8                oConnHO.Open
                 oConnHO.Close
            Else
                 oConnHO.ConnectionTimeout = 2
                 GoTo vb_error
            End If
9      Exit Function
vb_error:
10    Get_ConnectionHO_New = False
      gStatKoneksi = False
      Exit Function
End Function

Public Function Get_ConnectionNisSP(DBname As String, _
                                    UsrName As String, _
                                    Pass As String, _
                                    SerName As String) As Boolean
1      On Error GoTo vb_error
'2      App.Title = "BENGKEL-Nasmoco"
3      Get_ConnectionNisSP = True
4      If oConnNisSP.State Then oConnNisSP.Close
5      oConnNisSP.CursorLocation = adUseClient
6      oConnNisSP.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       KonDbNisSP = oConnNis.ConnectionString
7      adoConNisSP = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       SQLConACC1 = adoConNisSP
       oConnNisSP.ConnectionTimeout = 15
8      oConnNisSP.Open
       oConnNisSP.Close
       
        Dim nDB As New ADODB.Connection
        nDB.Open ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName)
       If nDB.State = adStateOpen Then nDB.Close
       
9      Exit Function

vb_error:
10    Get_ConnectionNisSP = False
'11    MsgBox ErrMessage(Erl, err.Number, "Function : Get_ConnectionNis"), vbCritical, "Open Connection"
End Function

Public Function Get_ConnectionNisSP2(DBname As String, _
                                    UsrName As String, _
                                    Pass As String, _
                                    SerName As String) As Boolean
1      On Error GoTo vb_error
'2      App.Title = "BENGKEL-Nasmoco"
3      Get_ConnectionNisSP2 = True
4      If oConnNisSP2.State Then oConnNisSP2.Close
5      oConnNisSP2.CursorLocation = adUseClient
6      oConnNisSP2.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       KonDbNisSP2 = oConnNis.ConnectionString
7      adoConNisSP2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       SQLConACC2 = adoConNisSP2
       oConnNisSP2.ConnectionTimeout = 15
8      oConnNisSP2.Open
       oConnNisSP2.Close
9      Exit Function

vb_error:
10    Get_ConnectionNisSP2 = False
'11    MsgBox ErrMessage(Erl, err.Number, "Function : Get_ConnectionNis"), vbCritical, "Open Connection"
End Function


Public Function Get_ConnectionQueue(DBname As String, _
                               UsrName As String, _
                               Pass As String, _
                               SerName As String) As Boolean
1      On Error GoTo vb_error
3      Get_ConnectionQueue = True
4      If oConnQueue.State Then oConnQueue.Close
5      oConnQueue.CursorLocation = adUseClient
6      oConnQueue.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       oConnQueue.ConnectionTimeout = 15
8      oConnQueue.Open

9      Exit Function

vb_error:
10    Get_ConnectionQueue = False
'11    MsgBox ErrMessage(Erl, err.Number, "Function : Get_ConnectionQueue"), vbCritical, "Open Connection"
End Function


Public Function Get_ConnectionAcces(DBname As String, _
                                    UsrName As String, _
                                    Pass As String, _
                                    SerName As String) As Boolean
1      On Error GoTo vb_error
2      'App.Title = "BENGKEL-Nasmoco"
3      Get_ConnectionAcces = True
4      If oConnAcces.State Then oConnAcces.Close
5      oConnAcces.CursorLocation = adUseClient
6      oConnAcces.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       KonDbAcces = oConnAcces.ConnectionString
7      adoConAcces = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       oConnAcces.ConnectionTimeout = 15
8      oConnAcces.Open
'       oConnAcces.Close
9      Exit Function

vb_error:

10    Get_ConnectionAcces = False
11    'MsgBox ErrMessage(Erl, err.Number, "Function : Get_ConnectionAcces "), vbCritical, "Open Connection"
End Function

Public Function Get_ConnectionAcces2(DBname As String, _
                                     UsrName As String, _
                                     Pass As String, _
                                     SerName As String) As Boolean
1      On Error GoTo vb_error
'2      App.Title = "BENGKEL-Nasmoco"
3      Get_ConnectionAcces2 = True
4      If oConnAcces2.State Then oConnAcces2.Close
5      oConnAcces2.CursorLocation = adUseClient
6      oConnAcces2.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       KonDbAcces2 = oConnAcces2.ConnectionString
7      adoConAcces2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       oConnAcces2.ConnectionTimeout = 15
8      oConnAcces2.Open
       oConnAcces2.Close
9      Exit Function

vb_error:
10    Get_ConnectionAcces2 = False
'11    MsgBox ErrMessage(Erl, err.Number, "Function : Get_ConnectionAcces2 "), vbCritical, "Open Connection"
End Function

Public Function ErrMessage(ByVal Errline As Long, _
                           ByVal ErrNumber As Long, _
                           FunctionName As String) As String
    ErrMessage = "Error Line Number = " & Errline & vbCrLf & _
                 "Error Number      = " & ErrNumber & vbCrLf & _
                 "Error Description = " & Error$(ErrNumber) & vbCrLf & _
                 "Location Error    = " & FunctionName
End Function

Public Function isNotBlank(txt As String) As Boolean
1        On Error GoTo isNotBlank_Error
2         If Trim(txt) = "" Then
3             isNotBlank = False
4         Else
5             isNotBlank = True
6         End If
7        On Error GoTo 0
8        Exit Function
isNotBlank_Error:
9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isNotBlank of Module MainModul")
End Function

Public Sub Enter(ByVal Key As Integer)
    If Key = 13 Then Sendkeys "{tab}"
End Sub

Function NulSID(kode As Variant, Optional Data)
1        On Error GoTo Nul_Error
2           kode = Trim(kode)
3           If IsMissing(Data) Then Data = ""
4           NulSID = UCase(Replace(IIf(IsNull(kode) Or kode = "", Data, kode), vbNewLine, "")) 'Replace(IIf(IsNull(kode) Or kode = "", Data, kode), vbNewLine, "")
5        On Error GoTo 0
6        Exit Function
Nul_Error:
7         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Nul of Module MainModul")
End Function

Function Nul(kode As Variant, Optional Data)
1        On Error GoTo Nul_Error
2           kode = Trim(kode)
3           If IsMissing(Data) Then Data = ""
4           Nul = Replace(IIf(IsNull(kode) Or kode = "", Data, kode), vbNewLine, "") 'Replace(IIf(IsNull(kode) Or kode = "", Data, kode), vbNewLine, "")
5        On Error GoTo 0
6        Exit Function
Nul_Error:
7         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Nul of Module MainModul")
End Function


Function NulTgl(kode As Variant)
    NulTgl = IIf(IsNull(kode) Or kode = "" Or kode = "0", "01-01-1900", kode)
End Function

Function NulA(kode As Variant, Optional Data)
      If IsMissing(Data) Then Data = "0"
      NulA = IIf(IsNull(kode) Or kode = "", Data, kode)
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
1        On Error GoTo isiHari_Error

2         Select Case hr
          Case "2"
3             isiHari = "Senin"
4         Case "3"
5             isiHari = "Selasa"
6         Case "4"
7             isiHari = "Rabu"
8         Case "5"
9             isiHari = "Kamis"
10        Case "6"
11            isiHari = "Jumat"
12        Case "7"
13            isiHari = "Sabtu"
14        Case "1"
15            isiHari = "Minggu"
16        End Select

17       On Error GoTo 0
18       Exit Function

isiHari_Error:

19        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiHari of Module MainModul")
End Function

Function PanggilCust(sAwal As String, txt As String, sAkhir As String, antri As Boolean, Optional sLoket As String, Optional lain As String)
    Dim h(50) As String
    Dim a() As String
    Dim i As Integer
    Dim hasil As String
    

    If antri Then
        hasil = Terbilang(Val(txt))
        hasil = Replace(hasil, ";;", ";")
        hasil = IIf(Right(hasil, 1) = ";", Left(hasil, Len(hasil) - 1), hasil)
        a() = Split(hasil, ";")
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
        For i = 0 To UBound(a)
           sndPlaySound App.Path & "\sound\" & a(i), 0
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
            
            Select Case UCase(gsSql2(i))
                Case "GETDATE()"
                    xHasil = Left(xHasil, Len(xHasil) - 1)
            End Select
                
            aa = UCase(gsSql1(i))
            Select Case aa
                Case "PPN", "DISC", "DISCT", "PART_G", "PART_NG", "OIL_G", "OIL_NG", "MTRL", "MET", "SBLT", "LABOR", "BODY", "TOTAL", "FN_EXT_DISKON"
                    xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
                Case "TIME", "INVTIME"
                    xHasil = Left(xHasil, Len(xHasil) - 1)
                Case Else
                    If UCase(gsSql2(i)) <> "GETDATE()" Then
                        xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
                    End If
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
            
            Select Case UCase(gsSql2(i))
                Case "GETDATE()"
                    gsSql2(i) = "Convert(VarChar(12), getdate(), 101)"
                    xHasil = xHasil + gsSql2(i)
            End Select
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC", "DISCT", "PART_G", "PART_NG", "OIL_G", "OIL_NG", "MTRL", "MET", "SBLT", "LABOR", "BODY", "TOTAL", "FN_EXT_DISKON"
                xHasil = xHasil + gsSql2(i)
            Case "TIME", "INVTIME"
                gsSql2(i) = "Convert(VarChar(5), getdate(), 108)"
                xHasil = xHasil + gsSql2(i)
            Case Else
                If UCase(gsSql2(i)) <> "CONVERT(VARCHAR(12), GETDATE(), 101)" Then
                    xHasil = xHasil + "'" + gsSql2(i) + "'"
                End If
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

Function cekPartPack(tpJob As String, jobOr As String, nMDL As String, idx As Integer, kodeGrid As Integer, Optional barispart As Integer, Optional frm As Object, Optional ByRef rsS As Recordset)
      Dim rsHrg As ADODB.Recordset
      Dim hrg
      Dim satPart As String
      Dim oliPart As Boolean
      Dim stokCukup As Boolean
      Dim nC As Integer
      Dim i As Integer
      Dim j As Integer
      Dim inputFrc As String
      Dim temmp As String
      Dim jmlStok As Integer
      Dim boleh As Boolean
      Dim iNo As Double
      Dim nNomor As Double
1     nNomor = 0

2         Set oRs2 = support.partPackage(tpJob, jobOr, nMDL) 'oConn.Execute("SELECT * FROM ho_partPack WHERE job = '" & Trim(tpJob) & "' AND FS_ITEm = '" & Trim(jobOr) & "' AND mdl = '" & nMDL & "' and FS_LOG_DEL = ''")
3         If Not oRs2.EOF Then
4             Do While Not oRs2.EOF
5                 boleh = True
6                 If cekGrid(oRs2, barispart, frm, kodeGrid, rsS) And idx = 1 Then
ulangi:
                      
7                         If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
8                             inputFrc = "X"
9                         Else
10                            inputFrc = InputBox("Isi Kode Franchise part " & vbCr & vbCr & "Kode : " & Trim(oRs2("ITEM_PART")) & vbCr & "Nama : " & Trim(oRs2("NAMA_PART")), "FRANCHISE")
11                            If inputFrc = "" Then
12                                temmp = MsgBox("Membatalkan Part ?", vbYesNo, "VALIDASI PART")
13                                If temmp = vbNo Then
14                                    GoTo ulangi
15                                Else
16                                    Exit Function
17                                End If
18                            End If
19                        End If
                          
20                        If inputFrc <> "" Then
21                            inputFrc = UCase(inputFrc)
                              
                              Dim rsHlp As ADODB.Recordset
22                            If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
23                                Set rsHrg = objO.FetchDataACC("SELECT hargajuala, kodesatuan FROM inmst WITH (NOLOCK) WHERE kodeItem like '%" & Trim(oRs2("ITEM_PART")) & "'")
24                                If Not rsHrg.EOF Then
25                                    hrg = CDbl(rsHrg("HARGAJUALA"))
26                                    satPart = Trim(rsHrg("KODESATUAN"))
27                                Else
28                                    Set rsHlp = objO.FetchDataACC("SELECT * FROM inmst WITH (NOLOCK) WHERE kodeItem LIKE '%" & Trim(oRs2("ITEM_PART")) & "%'")
29                                    If rsHlp.EOF Then
                                          'MsgBox "Part " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ")" & " = " & CDbl(oRs2("QTY")) & vbCr & "Tidak ada di stok Sparepart. ", vbInformation, "KELENGKAPAN PART"
30                                        boleh = False
31                                    Else
                                          'MsgBox "Part " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ")" & " = " & CDbl(oRs2("QTY")) & vbCr & "Tidak ada di stok Sparepart. ", vbInformation, "KELENGKAPAN PART"
                                          ''jika kodepart beda
32                                        If (Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy")) = False Then
33                                            hrg = 0
34                                            satPart = ""
35                                            GoTo ulangi
36                                        End If
37                                    End If
38                                    rsHlp.Close
39                                End If
40                                Call setClose(rsHrg)
41                            Else
42                                Set rsHrg = objO.FetchDataACC("SELECT hargajuala, kodesatuan FROM inmst WITH (NOLOCK) WHERE kodeItem = '" & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & "'")
43                                If Not rsHrg.EOF Then
44                                    hrg = CDbl(rsHrg("HARGAJUALA"))
45                                    satPart = Trim(rsHrg("KODESATUAN"))
46                                Else
47                                    Set rsHlp = objO.FetchDataACC("SELECT * FROM inmst WITH (NOLOCK) WHERE kodeItem LIKE '%" & Trim(oRs2("ITEM_PART")) & "%'")
48                                    If rsHlp.EOF Then
49                                        MsgBox "Part " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ")" & " = " & CDbl(oRs2("QTY")) & vbCr & "Tidak ada di stok Sparepart. ", vbInformation, "KELENGKAPAN PART"
50                                        boleh = False
51                                    Else
52                                        MsgBox "Part " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ")" & " = " & CDbl(oRs2("QTY")) & vbCr & "Tidak ada di stok Sparepart. ", vbInformation, "KELENGKAPAN PART"
53                                        hrg = 0
54                                        satPart = ""
55                                        GoTo ulangi
56                                    End If
57                                    rsHlp.Close
58                                End If
59                                Call setClose(rsHrg)
60                            End If

61                            If boleh Then
62                            If kodeGrid = 1 Then        'DENGAN FLEXGRID
63                                With frm
                                      
64                                    If nNomor = 0 Then
65                                        If .partGrid.TextMatrix(barispart - 1, 8) <> "" Then
66                                            nNomor = .partGrid.TextMatrix(barispart - 1, 8)
67                                        Else
68                                            nNomor = 1
69                                        End If
70                                    End If
71                                    If Not cekDouble(.partGrid, Trim(inputFrc & "-" & Trim(oRs2("ITEM_PART"))), barispart, 1) Then
inPutNewPart:
72                                        If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
73                                            .partGrid.TextMatrix(barispart, 1) = Trim(oRs2("item_part")) ''""
74                                            .partGrid.TextMatrix(barispart, 2) = IIf(.Name = "frmEstimasi", Trim(oRs2("NAMA_PART")), "")
75                                            .partGrid.TextMatrix(barispart, 3) = Trim(satPart)
76                                            .partGrid.TextMatrix(barispart, 4) = Trim(oRs2("QTY"))
77                                            .partGrid.TextMatrix(barispart, 5) = CDbl(hrg)
                                              
78                                            Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
79                                            Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("item_part"))) & "' and B.Jenis = 'OLI'")
                                                  
80                                            If Not rsHrg.EOF Then
81                                                .partGrid.TextMatrix(barispart, 6) = True
82                                                oliPart = True
83                                            Else
84                                                .partGrid.TextMatrix(barispart, 6) = False
85                                                oliPart = False
86                                            End If
                                              
87                                            .partGrid.TextMatrix(barispart, 7) = IIf(.Name = "frmEstimasi", "", Trim(oRs2("NAMA_PART")))
88                                            .partGrid.TextMatrix(barispart, 8) = barispart + nNomor + 1
89                                            barispart = barispart + 1
                                          
90                                            Call setClose(rsHrg)
91                                            If .Name = "frmBooking" Then
92                                                If oliPart Then
93                                                    .Text(55) = CDbl(.Text(55)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
94                                                    .Text(55) = Format(.Text(55), "#0,0")
95                                                Else
96                                                    .Text(40) = CDbl(.Text(40)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
97                                                    .Text(40) = Format(.Text(40), "#0,0")
98                                                End If
99                                            ElseIf .Name = "frmEstimasi" Then
100                                               If oliPart Then
101                                                   .Text(28) = CDbl(.Text(28)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
102                                                   .Text(28) = Format(.Text(28), "#0,0")
103                                               Else
104                                                   .Text(24) = CDbl(.Text(24)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
105                                                   .Text(24) = Format(.Text(24), "#0,0")
106                                               End If
107                                           End If
                                          
108                                       Else
109                                           .partGrid.TextMatrix(barispart, 1) = inputFrc & "-" & Trim(oRs2("ITEM_PART"))
110                                           .partGrid.TextMatrix(barispart, 2) = Trim(oRs2("NAMA_PART"))
111                                           .partGrid.TextMatrix(barispart, 3) = Trim(satPart)
112                                           .partGrid.TextMatrix(barispart, 4) = Trim(oRs2("QTY"))
113                                           .partGrid.TextMatrix(barispart, 5) = CDbl(hrg)
                                              
114                                           Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
115                                           If Not rsHrg.EOF Then
116                                               .partGrid.TextMatrix(barispart, 6) = True
117                                               oliPart = True
118                                           Else
119                                               .partGrid.TextMatrix(barispart, 6) = False
120                                               oliPart = False
121                                           End If
                                              
122                                           .partGrid.TextMatrix(barispart, 7) = ""
123                                           .partGrid.TextMatrix(barispart, 8) = nNomor + barispart + 1
124                                           barispart = barispart + 1
                                          
125                                           Call setClose(rsHrg)
126                                           If .Name = "frmBooking" Then
127                                               If oliPart Then
128                                                   .Text(55) = CDbl(.Text(55)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
129                                                   .Text(55) = Format(.Text(55), "#0,0")
130                                               Else
131                                                   .Text(40) = CDbl(.Text(40)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
132                                                   .Text(40) = Format(.Text(40), "#0,0")
133                                               End If
134                                           ElseIf .Name = "frmEstimasi" Then
135                                               If oliPart Then
136                                                   .Text(28) = CDbl(.Text(28)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
137                                                   .Text(28) = Format(.Text(28), "#0,0")
138                                               Else
139                                                   .Text(24) = CDbl(.Text(24)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
140                                                   .Text(24) = Format(.Text(24), "#0,0")
141                                               End If
142                                           End If
143                                       End If
144                                   Else
145                                       If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
146                                           GoTo inPutNewPart
147                                       Else
148                                           If MsgBox("Part " & Trim(inputFrc & "-" & Trim(oRs2("ITEM_PART"))) & " Sudah Terdaftar, Tetap Inputkan Part ?", vbQuestion + vbYesNo, "Part Paket") = vbYes Then
149                                               GoTo inPutNewPart
150                                           End If
151                                       End If
152                                   End If
153                               End With
154                           ElseIf kodeGrid = 2 Then    'DENGAN RECORDSET
155                               If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
156                                   Set rsHrg = objO.FetchDataACC("SELECT SUM(ISNULL(B.QTY,0)) as QTYSTOCK FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem like '%" & Trim(oRs2("ITEM_PART")) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
157                                   If Not rsHrg.EOF Then
158                                       stokCukup = IIf(CDbl(rsHrg("QTYSTOCK")) >= CDbl(oRs2("QTY")), True, False)
159                                   End If
                                      
160                                   If stokCukup Then
161                                       With rsS
                                              '.Find "KODEITEM = '" & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & "'", , adSearchForward, 1
                                              'If .EOF Then
162                                               .AddNew
163                                               !KodeItem = Trim(oRs2("ITEM_PART")) ''""
164                                               !NAMAITEM = Trim(oRs2("NAMA_PART")) ''""
165                                               !SATITEM = Trim(satPart)
166                                               !qty = Trim(oRs2("QTY"))
                                                  ''Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
167                                               Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("ITEM_PART"))) & "' and B.Jenis = 'OLI'")
                                                  
168                                               If Not rsHrg.EOF Then
169                                                   !oli = True
170                                               Else
171                                                   !oli = False
172                                               End If
                                                  
173                                               !namabySA = Trim(oRs2("NAMA_PART"))
174                                               !nID = rsS.RecordCount + 1
                      
175                                               Call setClose(rsHrg)
176                                               .Update
177                                               If frm.Name = "frmOrder" Then
178                                                   With frm
179                                                       If rsS("OLI") Then
180                                                           .Text(59) = CDbl(.Text(59)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
181                                                           .Text(59) = Format(.Text(59), "#0,0")
182                                                       Else
183                                                           .Text(57) = CDbl(.Text(57)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
184                                                           .Text(57) = Format(.Text(57), "#0,0")
185                                                       End If
186                                                   End With
187                                               ElseIf frm.Name = "frmOAddJob" Then
188                                                   With frm
189                                                       If rsS("OLI") Then
190                                                           .Text(23) = CDbl(.Text(23)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
191                                                           .Text(23) = Format(.Text(23), "#0,0")
192                                                       Else
193                                                           .Text(22) = CDbl(.Text(22)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
194                                                           .Text(22) = Format(.Text(22), "#0,0")
195                                                       End If
196                                                   End With
197                                               End If
                                              'End If
198                                           rsS.MoveNext
199                                       End With
200                                   Else
201                                       With rsS
                                              '.Find "KODEITEM = '" & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & "'", , adSearchForward, 1
                                              'If .EOF Then
202                                               .AddNew
203                                               !KodeItem = inputFrc & "-" & Trim(oRs2("ITEM_PART"))
204                                               !NAMAITEM = Trim(oRs2("NAMA_PART"))
205                                               !SATITEM = Trim(satPart)
206                                               !qty = Trim(oRs2("QTY"))
207                                               Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
208                                               Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("ITEM_PART"))) & "' and B.Jenis = 'OLI'")
209                                               If Not rsHrg.EOF Then
210                                                   !oli = True
211                                               Else
212                                                   !oli = False
213                                               End If
                                                  
214                                               !namabySA = ""
215                                               !nID = rsS.RecordCount + 1
                      
216                                               Call setClose(rsHrg)
217                                               .Update
218                                               If frm.Name = "frmOrder" Then
                                                      
219                                               ElseIf frm.Name = "frmOAddJob" Then
220                                                   With frm
221                                                       If rsS("OLI") Then
222                                                           .Text(23) = CDbl(.Text(23)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
223                                                           .Text(23) = Format(.Text(23), "#0,0")
224                                                       Else
225                                                           .Text(22) = CDbl(.Text(22)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
226                                                           .Text(22) = Format(.Text(22), "#0,0")
227                                                       End If
228                                                   End With
229                                               End If
                                              'End If
230                                           rsS.MoveNext
231                                       End With
                                          
232                                   End If
233                               Else
234                                   Set rsHrg = objO.FetchDataACC("SELECT SUM(ISNULL(B.QTY,0)) as QTYSTOCK FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Trim(inputFrc & "-" & Trim(oRs2("ITEM_PART"))) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
235                                   If Not rsHrg.EOF Then
236                                       stokCukup = IIf(CDbl(rsHrg("QTYSTOCK")) >= CDbl(oRs2("QTY")), True, False)
237                                   End If
                                      
238                                   If stokCukup Then
239                                       With rsS
240                                           .Find "KODEITEM = '" & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & "'", , adSearchForward, 1
241                                           If .EOF Then
242                                               .AddNew
243                                               !KodeItem = inputFrc & "-" & Trim(oRs2("ITEM_PART"))
244                                               !NAMAITEM = Trim(oRs2("NAMA_PART"))
245                                               !SATITEM = Trim(satPart)
246                                               !qty = Trim(oRs2("QTY"))
247                                               Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
248                                               If Not rsHrg.EOF Then
249                                                   !oli = True
250                                               Else
251                                                   !oli = False
252                                               End If
                                                  
253                                               !namabySA = ""
254                                               !nID = rsS.RecordCount + 1
                      
255                                               Call setClose(rsHrg)
256                                               .Update
257                                               If frm.Name = "frmOrder" Then
258                                               ElseIf frm.Name = "frmOAddJob" Then
259                                                   With frm
260                                                       If rsS("OLI") Then
261                                                           .Text(23) = CDbl(.Text(23)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
262                                                           .Text(23) = Format(.Text(23), "#0,0")
263                                                       Else
264                                                           .Text(22) = CDbl(.Text(22)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
265                                                           .Text(22) = Format(.Text(22), "#0,0")
266                                                       End If
267                                                   End With
268                                               End If
269                                           End If
270                                           rsS.MoveNext
271                                       End With
272                                   Else
273                                       MsgBox "Jumlah Stok " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ") tidak memenuhi permintaan."
274                                   End If
275                               End If
276                           End If
277                           End If
278                       End If
                      
279               ElseIf idx = 2 Then 'Not cekGrid(oRs2, BarisPart, frm, kodeGrid) And
280                   If kodeGrid = 1 Then
281                       For i = 1 To barispart - 1
282                           With frm
283                               If Trim(.partGrid.TextMatrix(i, 1)) <> "" Then
284                                   If Trim(.partGrid.TextMatrix(i, 1)) = Trim(oRs2("ITEM_PART")) Then
                                      ''If Mid(.partGrid.TextMatrix(i, 1), 3, Len(.partGrid.TextMatrix(i, 1)) - 2) = Trim(oRs2("ITEM_PART")) Then
285                                           Set rsHrg = objO.FetchDataACC("SELECT hargajuala FROM inmst WITH (NOLOCK) WHERE kodeItem like '%" & Trim(.partGrid.TextMatrix(i, 1)) & "'")
286                                           If Not rsHrg.EOF Then
287                                               hrg = CDbl(rsHrg("HARGAJUALA"))
288                                           End If
289                                           Call setClose(rsHrg)
                                              
                                              ''Set rsHrg = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(.partGrid.TextMatrix(i, 1), 1) & "' AND TIPE = 'O'")
290                                           Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("ITEM_PART"))) & "' and B.Jenis = 'OLI'")
291                                           If Not rsHrg.EOF Then
292                                               oliPart = True
293                                           Else
294                                               oliPart = False
295                                           End If
296                                       If .Name = "frmBooking" Then
297                                           If oliPart Then
298                                               If CDbl(.Text(55)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg)) >= 0 Then
299                                               .Text(55) = CDbl(.Text(55)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
300                                               Else
301                                               .Text(55) = 0
302                                               End If
303                                               .Text(55) = Format(.Text(55), "#0,0")
304                                           Else
305                                               If CDbl(.Text(40)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg)) >= 0 Then
306                                               .Text(40) = CDbl(.Text(40)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
307                                               Else
308                                               .Text(40) = 0
309                                               End If
310                                               .Text(40) = Format(.Text(40), "#0,0")
311                                           End If
                                              
312                                       ElseIf .Name = "frmEstimasi" Then
313                                           If oliPart Then
314                                               .Text(28) = CDbl(.Text(28)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
315                                               .Text(28) = Format(.Text(28), "#0,0")
316                                           Else
317                                               .Text(24) = CDbl(.Text(24)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
318                                               .Text(24) = Format(.Text(24), "#0,0")
319                                           End If
                                              
320                                       End If
                                          
321                                       nC = i
322                                       For j = nC To barispart
323                                           With .partGrid
324                                               .TextMatrix(j, 0) = .TextMatrix(j + 1, 0)
325                                               .TextMatrix(j, 1) = .TextMatrix(j + 1, 1)
326                                               .TextMatrix(j, 2) = .TextMatrix(j + 1, 2)
327                                               .TextMatrix(j, 3) = .TextMatrix(j + 1, 3)
328                                               .TextMatrix(j, 4) = .TextMatrix(j + 1, 4)
329                                               .TextMatrix(j, 5) = .TextMatrix(j + 1, 5)
330                                               .TextMatrix(j, 6) = .TextMatrix(j + 1, 6)
331                                               .TextMatrix(j, 7) = .TextMatrix(j + 1, 7)
332                                               .TextMatrix(j, 8) = .TextMatrix(j + 1, 8)
333                                           End With
334                                       Next
335                                       barispart = barispart - 1
336                                   End If
337                               ElseIf Trim(.partGrid.TextMatrix(i, IIf(.Name = "frmEstimasi", 2, 7))) <> "" Then
338                                   If (Trim(.partGrid.TextMatrix(i, IIf(.Name = "frmEstimasi", 2, 7))) = Trim(oRs2("NAMA_PART"))) Then
                                              'Set rsHrg = objO.FetchDataACC("SELECT hargajuala FROM inmst WHERE kodeItem = '" & Trim(.partGrid.TextMatrix(i, 1)) & "'")
                                              'If Not rsHrg.EOF Then
339                                               hrg = CDbl(.partGrid.TextMatrix(i, 5))
                                              'End If
                                              'Call setClose(rsHrg)
                                              
                                              'Set rsHrg = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WHERE FS_KD_OLI = '" & Left(.partGrid.TextMatrix(i, 1), 1) & "' AND TIPE = 'O'")
                                              'If Not rsHrg.EOF Then
                                              '    oliPart = True
                                              'Else
340                                               oliPart = False
                                              'End If
341                                       If .Name = "frmBooking" Then
342                                           If oliPart Then
343                                               .Text(55) = CDbl(.Text(55)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
344                                               .Text(55) = Format(.Text(55), "#0,0")
345                                           Else
346                                               .Text(40) = CDbl(.Text(40)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
347                                               .Text(40) = Format(.Text(40), "#0,0")
348                                           End If
                                              
349                                       ElseIf .Name = "frmEstimasi" Then
350                                           If oliPart Then
351                                               .Text(28) = CDbl(.Text(28)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
352                                               .Text(28) = Format(.Text(28), "#0,0")
353                                           Else
354                                               .Text(24) = CDbl(.Text(24)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
355                                               .Text(24) = Format(.Text(24), "#0,0")
356                                           End If
                                              
357                                       End If
                                          
358                                       nC = i
359                                       For j = nC To barispart
360                                           With .partGrid
361                                               .TextMatrix(j, 0) = .TextMatrix(j + 1, 0)
362                                               .TextMatrix(j, 1) = .TextMatrix(j + 1, 1)
363                                               .TextMatrix(j, 2) = .TextMatrix(j + 1, 2)
364                                               .TextMatrix(j, 3) = .TextMatrix(j + 1, 3)
365                                               .TextMatrix(j, 4) = .TextMatrix(j + 1, 4)
366                                               .TextMatrix(j, 5) = .TextMatrix(j + 1, 5)
367                                               .TextMatrix(j, 6) = .TextMatrix(j + 1, 6)
368                                               .TextMatrix(j, 7) = .TextMatrix(j + 1, 7)
369                                               .TextMatrix(j, 8) = .TextMatrix(j + 1, 8)
370                                           End With
371                                       Next
372                                       barispart = barispart - 1
373                                   End If

374                               ElseIf Trim(.partGrid.TextMatrix(i, 2)) <> "" Then
                                              
375                                           hrg = .partGrid.TextMatrix(i, 5)
376                                           oliPart = False
                                              
377                                       If .Name = "frmBooking" Then
378                                           If oliPart Then
379                                               .Text(55) = CDbl(.Text(55)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
380                                               .Text(55) = Format(.Text(55), "#0,0")
381                                           Else
382                                               .Text(40) = CDbl(.Text(40)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
383                                               .Text(40) = Format(.Text(40), "#0,0")
384                                           End If
                                              
385                                       ElseIf .Name = "frmEstimasi" Then
386                                           If oliPart Then
387                                               .Text(28) = CDbl(.Text(28)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
388                                               .Text(28) = Format(.Text(28), "#0,0")
389                                           Else
390                                               .Text(24) = CDbl(.Text(24)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
391                                               .Text(24) = Format(.Text(24), "#0,0")
392                                           End If
                                              
393                                       End If
                                          
394                                       nC = i
395                                       For j = nC To barispart
396                                           With .partGrid
397                                               .TextMatrix(j, 0) = .TextMatrix(j + 1, 0)
398                                               .TextMatrix(j, 1) = .TextMatrix(j + 1, 1)
399                                               .TextMatrix(j, 2) = .TextMatrix(j + 1, 2)
400                                               .TextMatrix(j, 3) = .TextMatrix(j + 1, 3)
401                                               .TextMatrix(j, 4) = .TextMatrix(j + 1, 4)
402                                               .TextMatrix(j, 5) = .TextMatrix(j + 1, 5)
403                                               .TextMatrix(j, 6) = .TextMatrix(j + 1, 6)
404                                               .TextMatrix(j, 7) = .TextMatrix(j + 1, 7)
405                                               .TextMatrix(j, 8) = .TextMatrix(j + 1, 8)
406                                           End With
407                                       Next
408                                       barispart = barispart - 1

409                               End If
410                           End With
411                       Next
412                   ElseIf kodeGrid = 2 Then
413                       With rsS
414                           .Find "KODEITEM LIKE '%" & Trim(oRs2("ITEM_PART")) & "%'", , adSearchForward, 1
415                           If Not .EOF Then
416                                   Set rsHrg = objO.FetchDataACC("SELECT hargajuala FROM inmst WITH (NOLOCK) WHERE kodeItem like '%" & rsS("KODEITEM") & "'")
417                                   If Not rsHrg.EOF Then
418                                       hrg = CDbl(rsHrg("HARGAJUALA"))
419                                   End If
420                                   rsHrg.Close
                                      
                                      ''Set rsHrg = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsS("KODEITEM"), 1) & "' AND TIPE = 'O'")
421                                   Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("ITEM_PART"))) & "' and B.Jenis = 'OLI'")
422                                   If Not rsHrg.EOF Then
423                                       oliPart = True
424                                   Else
425                                       oliPart = False
426                                   End If
                                  
427                               With frm
428                                   If oliPart Then
429                                       .Text(59) = CDbl(.Text(59)) - (CDbl(rsS("QTY")) * CDbl(hrg))
430                                       .Text(59) = Format(.Text(59), "#0,0")
431                                   Else
432                                       .Text(57) = CDbl(.Text(57)) - (CDbl(rsS("QTY")) * CDbl(hrg))
433                                       .Text(57) = Format(.Text(57), "#0,0")
434                                   End If
435                               End With
436                               .Delete
437                               .Update
438                           Else
                              
439                               .Find "NAMABYSA = '" & Trim(oRs2("NAMA_PART")) & "'", , adSearchForward, 1
440                               If Not .EOF Then
                                          'Set rsHrg = objO.FetchDataACC("SELECT hargajuala FROM inmst WHERE NamaBarang = '" & rsS("NAMABYSA") & "'")
                                          'If Not rsHrg.EOF Then
441                                           hrg = 0 'CDbl(.Fields(1).value)
                                          'End If
                                          'rsHrg.Close
                                          
                                          'Set rsHrg = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WHERE FS_KD_OLI = '" & Left(rsS("KODEITEM"), 1) & "' AND TIPE = 'O'")
                                          'If Not rsHrg.EOF Then
                                          '    oliPart = True
                                          'Else
442                                           oliPart = False
                                          'End If
                                      
443                                   With frm
444                                       If oliPart Then
445                                           .Text(59) = CDbl(.Text(59)) - (CDbl(rsS("QTY")) * CDbl(hrg))
446                                           .Text(59) = Format(.Text(59), "#0,0")
447                                       Else
448                                           .Text(57) = CDbl(.Text(57)) - (CDbl(rsS("QTY")) * CDbl(hrg))
449                                           .Text(57) = Format(.Text(57), "#0,0")
450                                       End If
451                                   End With
452                                   .Delete
453                                   .Update
454                               End If
455                           End If
456                       End With
457                   End If
458               End If
459               oRs2.MoveNext
460           Loop
461       End If
462       If oRs2.State <> adStateClosed Then oRs2.Close
End Function


Function cekFopPart(tpJob As String, jobOr As String, nMDL As String, idx As Integer, kodeGrid As Integer, Optional barispart As Integer, Optional frm As Object, Optional ByRef rsS As Recordset)
      Dim rsHrg As ADODB.Recordset
      Dim hrg
      Dim satPart     As String
      Dim namaPart    As String
      Dim oliPart As Boolean
      Dim stokCukup As Boolean
      Dim nC As Integer
      Dim i As Integer
      Dim j As Integer
      Dim inputFrc As String
      Dim temmp As String
      Dim jmlStok As Integer
      Dim boleh As Boolean
      Dim iNo As Double
      Dim nNomor As Double
1     nNomor = 0

2         Set oRs2 = support.partFopPackage(tpJob, jobOr, nMDL) 'oConn.Execute("SELECT * FROM ho_partPack WHERE job = '" & Trim(tpJob) & "' AND FS_ITEm = '" & Trim(jobOr) & "' AND mdl = '" & nMDL & "' and FS_LOG_DEL = ''")
3         If Not oRs2.EOF Then
4             Do While Not oRs2.EOF
5                 boleh = True
6                 If cekGridFOP(oRs2, barispart, frm, kodeGrid, rsS) And idx = 1 Then
ulangi:
7                         namaPart = ""
8                         If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
9                             inputFrc = "X"
10                        Else
11                            inputFrc = InputBox("Isi Kode Franchise part " & vbCr & vbCr & "Kode : " & Trim(oRs2("ITEM_PART")) & vbCr & "Nama : " & Trim(oRs2("NAMA_PART")), "FRANCHISE")
12                            If inputFrc = "" Then
13                                temmp = MsgBox("Membatalkan Part ?", vbYesNo, "VALIDASI PART")
14                                If temmp = vbNo Then
15                                    GoTo ulangi
16                                Else
17                                    Exit Function
18                                End If
19                            End If
20                        End If
                          
21                        If inputFrc <> "" Then
22                            inputFrc = UCase(inputFrc)
                              
                              Dim rsHlp As ADODB.Recordset
23                            If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
24                                Set rsHrg = objO.FetchDataACC("SELECT hargajuala, kodesatuan,namabarang FROM inmst WITH (NOLOCK) WHERE kodeItem like '%" & Trim(oRs2("ITEM_PART")) & "'")
25                                If Not rsHrg.EOF Then
26                                    hrg = CDbl(rsHrg("HARGAJUALA"))
27                                    satPart = Trim(rsHrg("KODESATUAN"))
28                                    namaPart = Trim(rsHrg!NamaBarang)
29                                Else
30                                    Set rsHlp = objO.FetchDataACC("SELECT * FROM inmst WITH (NOLOCK) WHERE kodeItem LIKE '%" & Trim(oRs2("ITEM_PART")) & "%'")
31                                    If rsHlp.EOF Then
                                          'MsgBox "Part " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ")" & " = " & CDbl(oRs2("QTY")) & vbCr & "Tidak ada di stok Sparepart. ", vbInformation, "KELENGKAPAN PART"
32                                        boleh = False
33                                    Else
                                          'MsgBox "Part " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ")" & " = " & CDbl(oRs2("QTY")) & vbCr & "Tidak ada di stok Sparepart. ", vbInformation, "KELENGKAPAN PART"
34                                        hrg = 0
35                                        satPart = ""
36                                        GoTo ulangi
37                                    End If
38                                    rsHlp.Close
39                                End If
40                                Call setClose(rsHrg)
41                            Else
42                                Set rsHrg = objO.FetchDataACC("SELECT hargajuala, kodesatuan FROM inmst WITH (NOLOCK) WHERE kodeItem = '" & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & "'")
43                                If Not rsHrg.EOF Then
44                                    hrg = CDbl(rsHrg("HARGAJUALA"))
45                                    satPart = Trim(rsHrg("KODESATUAN"))
46                                Else
47                                    Set rsHlp = objO.FetchDataACC("SELECT * FROM inmst WITH (NOLOCK) WHERE kodeItem LIKE '%" & Trim(oRs2("ITEM_PART")) & "%'")
48                                    If rsHlp.EOF Then
49                                        MsgBox "Part " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ")" & " = " & CDbl(oRs2("QTY")) & vbCr & "Tidak ada di stok Sparepart. ", vbInformation, "KELENGKAPAN PART"
50                                        boleh = False
51                                    Else
52                                        MsgBox "Part " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ")" & " = " & CDbl(oRs2("QTY")) & vbCr & "Tidak ada di stok Sparepart. ", vbInformation, "KELENGKAPAN PART"
53                                        hrg = 0
54                                        satPart = ""
55                                        GoTo ulangi
56                                    End If
57                                    rsHlp.Close
58                                End If
59                                Call setClose(rsHrg)
60                            End If

61                            If boleh Then
62                            If kodeGrid = 1 Then        'DENGAN FLEXGRID
63                                With frm
                                      
64                                    If nNomor = 0 Then
65                                        If .partGrid.TextMatrix(barispart - 1, 8) <> "" Then
66                                            nNomor = .partGrid.TextMatrix(barispart - 1, 8)
67                                        Else
68                                            nNomor = 1
69                                        End If
70                                    End If
71                                    If Not cekDouble(.partGrid, Trim(inputFrc & "-" & Trim(oRs2("ITEM_PART"))), barispart, 1) Then
inPutNewPart:
72                                        If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
73                                            .partGrid.TextMatrix(barispart, 1) = Trim(oRs2("item_part"))
74                                            .partGrid.TextMatrix(barispart, 2) = namaPart 'IIf(.Name = "frmEstimasi", Trim(oRs2("NAMA_PART")), "")
75                                            .partGrid.TextMatrix(barispart, 3) = Trim(satPart)
76                                            .partGrid.TextMatrix(barispart, 4) = Trim(oRs2("QTY"))
77                                            .partGrid.TextMatrix(barispart, 5) = CDbl(hrg)
                                              
                                              ''Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
78                                            Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("item_part"))) & "' and B.Jenis = 'OLI'")
79                                            If Not rsHrg.EOF Then
80                                                .partGrid.TextMatrix(barispart, 6) = True
81                                                oliPart = True
82                                            Else
83                                                .partGrid.TextMatrix(barispart, 6) = False
84                                                oliPart = False
85                                            End If
                                              
86                                            .partGrid.TextMatrix(barispart, 7) = IIf(.Name = "frmEstimasi", "", Trim(oRs2("NAMA_PART")))
87                                            .partGrid.TextMatrix(barispart, 8) = barispart + nNomor + 1
88                                            barispart = barispart + 1
                                          
89                                            Call setClose(rsHrg)
90                                            If .Name = "frmBooking" Then
91                                                If oliPart Then
92                                                    .Text(55) = CDbl(.Text(55)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
93                                                    .Text(55) = Format(.Text(55), "#0,0")
94                                                Else
95                                                    .Text(40) = CDbl(.Text(40)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
96                                                    .Text(40) = Format(.Text(40), "#0,0")
97                                                End If
98                                            ElseIf .Name = "frmEstimasi" Then
99                                                If oliPart Then
100                                                   .Text(28) = CDbl(.Text(28)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
101                                                   .Text(28) = Format(.Text(28), "#0,0")
102                                               Else
103                                                   .Text(24) = CDbl(.Text(24)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
104                                                   .Text(24) = Format(.Text(24), "#0,0")
105                                               End If
106                                           End If
                                          
107                                       Else
108                                           .partGrid.TextMatrix(barispart, 1) = inputFrc & "-" & Trim(oRs2("ITEM_PART"))
109                                           .partGrid.TextMatrix(barispart, 2) = Trim(oRs2("NAMA_PART"))
110                                           .partGrid.TextMatrix(barispart, 3) = Trim(satPart)
111                                           .partGrid.TextMatrix(barispart, 4) = Trim(oRs2("QTY"))
112                                           .partGrid.TextMatrix(barispart, 5) = CDbl(hrg)
                                              
113                                           Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
114                                           If Not rsHrg.EOF Then
115                                               .partGrid.TextMatrix(barispart, 6) = True
116                                               oliPart = True
117                                           Else
118                                               .partGrid.TextMatrix(barispart, 6) = False
119                                               oliPart = False
120                                           End If
                                              
121                                           .partGrid.TextMatrix(barispart, 7) = ""
122                                           .partGrid.TextMatrix(barispart, 8) = nNomor + barispart + 1
123                                           barispart = barispart + 1
                                          
124                                           Call setClose(rsHrg)
125                                           If .Name = "frmBooking" Then
126                                               If oliPart Then
127                                                   .Text(55) = CDbl(.Text(55)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
128                                                   .Text(55) = Format(.Text(55), "#0,0")
129                                               Else
130                                                   .Text(40) = CDbl(.Text(40)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
131                                                   .Text(40) = Format(.Text(40), "#0,0")
132                                               End If
133                                           ElseIf .Name = "frmEstimasi" Then
134                                               If oliPart Then
135                                                   .Text(28) = CDbl(.Text(28)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
136                                                   .Text(28) = Format(.Text(28), "#0,0")
137                                               Else
138                                                   .Text(24) = CDbl(.Text(24)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
139                                                   .Text(24) = Format(.Text(24), "#0,0")
140                                               End If
141                                           End If
142                                       End If
143                                   Else
144                                       If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
145                                           GoTo inPutNewPart
146                                       Else
147                                           If MsgBox("Part " & Trim(inputFrc & "-" & Trim(oRs2("ITEM_PART"))) & " Sudah Terdaftar, Tetap Inputkan Part ?", vbQuestion + vbYesNo, "Part Paket") = vbYes Then
148                                               GoTo inPutNewPart
149                                           End If
150                                       End If
151                                   End If
152                               End With
153                           ElseIf kodeGrid = 2 Then    'DENGAN RECORDSET
154                               If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
155                                   Set rsHrg = objO.FetchDataACC("SELECT SUM(ISNULL(B.QTY,0)) as QTYSTOCK FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem like '%" & Trim(oRs2("ITEM_PART")) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
156                                   If Not rsHrg.EOF Then
157                                       stokCukup = IIf(CDbl(rsHrg("QTYSTOCK")) >= CDbl(oRs2("QTY")), True, False)
158                                   End If
                                      
159                                   If stokCukup Then
160                                       With rsS
                                              '.Find "KODEITEM = '" & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & "'", , adSearchForward, 1
                                              'If .EOF Then
161                                               .AddNew
162                                               !KodeItem = Trim(oRs2("ITEM_PART")) ''""
163                                               !NAMAITEM = Trim(oRs2("NAMA_PART")) ''""
164                                               !SATITEM = Trim(satPart)
165                                               !qty = Trim(oRs2("QTY"))
                                                  'Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
166                                               Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("ITEM_PART"))) & "' and B.Jenis = 'OLI'")
167                                               If Not rsHrg.EOF Then
168                                                   !oli = True
169                                               Else
170                                                   !oli = False
171                                               End If
                                                  
172                                               !namabySA = Trim(oRs2("NAMA_PART"))
173                                               !nID = rsS.RecordCount + 1
                      
174                                               Call setClose(rsHrg)
175                                               .Update
176                                               If frm.Name = "frmOrder" Then
177                                                   With frm
178                                                       If rsS("OLI") Then
179                                                           .Text(59) = CDbl(.Text(59)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
180                                                           .Text(59) = Format(.Text(59), "#0,0")
181                                                       Else
182                                                           .Text(57) = CDbl(.Text(57)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
183                                                           .Text(57) = Format(.Text(57), "#0,0")
184                                                       End If
185                                                   End With
186                                               ElseIf frm.Name = "frmOAddJob" Then
187                                                   With frm
188                                                       If rsS("OLI") Then
189                                                           .Text(23) = CDbl(.Text(23)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
190                                                           .Text(23) = Format(.Text(23), "#0,0")
191                                                       Else
192                                                           .Text(22) = CDbl(.Text(22)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
193                                                           .Text(22) = Format(.Text(22), "#0,0")
194                                                       End If
195                                                   End With
196                                               End If
                                              'End If
197                                           rsS.MoveNext
198                                       End With
199                                   Else
200                                       With rsS
                                              '.Find "KODEITEM = '" & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & "'", , adSearchForward, 1
                                              'If .EOF Then
201                                               .AddNew
202                                               !KodeItem = inputFrc & "-" & Trim(oRs2("ITEM_PART"))
203                                               !NAMAITEM = Trim(oRs2("NAMA_PART"))
204                                               !SATITEM = Trim(satPart)
205                                               !qty = Trim(oRs2("QTY"))
206                                               Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
207                                               If Not rsHrg.EOF Then
208                                                   !oli = True
209                                               Else
210                                                   !oli = False
211                                               End If
                                                  
212                                               !namabySA = ""
213                                               !nID = rsS.RecordCount + 1
                      
214                                               Call setClose(rsHrg)
215                                               .Update
216                                               If frm.Name = "frmOrder" Then
                                                      
217                                               ElseIf frm.Name = "frmOAddJob" Then
218                                                   With frm
219                                                       If rsS("OLI") Then
220                                                           .Text(23) = CDbl(.Text(23)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
221                                                           .Text(23) = Format(.Text(23), "#0,0")
222                                                       Else
223                                                           .Text(22) = CDbl(.Text(22)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
224                                                           .Text(22) = Format(.Text(22), "#0,0")
225                                                       End If
226                                                   End With
227                                               End If
                                              'End If
228                                           rsS.MoveNext
229                                       End With
                                          
230                                   End If
231                               Else
232                                   Set rsHrg = objO.FetchDataACC("SELECT SUM(ISNULL(B.QTY,0)) as QTYSTOCK FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Trim(inputFrc & "-" & Trim(oRs2("ITEM_PART"))) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
233                                   If Not rsHrg.EOF Then
234                                       stokCukup = IIf(CDbl(rsHrg("QTYSTOCK")) >= CDbl(oRs2("QTY")), True, False)
235                                   End If
                                      
236                                   If stokCukup Then
237                                       With rsS
238                                           .Find "KODEITEM = '" & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & "'", , adSearchForward, 1
239                                           If .EOF Then
240                                               .AddNew
241                                               !KodeItem = inputFrc & "-" & Trim(oRs2("ITEM_PART"))
242                                               !NAMAITEM = Trim(oRs2("NAMA_PART"))
243                                               !SATITEM = Trim(satPart)
244                                               !qty = Trim(oRs2("QTY"))
245                                               Set rsHrg = objO.FetchData("SELECT ISNULL(FS_KD_OLI,'') AS KDOLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & inputFrc & "' AND TIPE = 'O'")
246                                               If Not rsHrg.EOF Then
247                                                   !oli = True
248                                               Else
249                                                   !oli = False
250                                               End If
                                                  
251                                               !namabySA = ""
252                                               !nID = rsS.RecordCount + 1
                      
253                                               Call setClose(rsHrg)
254                                               .Update
255                                               If frm.Name = "frmOrder" Then
256                                               ElseIf frm.Name = "frmOAddJob" Then
257                                                   With frm
258                                                       If rsS("OLI") Then
259                                                           .Text(23) = CDbl(.Text(23)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
260                                                           .Text(23) = Format(.Text(23), "#0,0")
261                                                       Else
262                                                           .Text(22) = CDbl(.Text(22)) + (CDbl(oRs2("QTY")) * CDbl(hrg))
263                                                           .Text(22) = Format(.Text(22), "#0,0")
264                                                       End If
265                                                   End With
266                                               End If
267                                           End If
268                                           rsS.MoveNext
269                                       End With
270                                   Else
271                                       MsgBox "Jumlah Stok " & inputFrc & "-" & Trim(oRs2("ITEM_PART")) & " (" & Trim(oRs2("NAMA_PART")) & ") tidak memenuhi permintaan."
272                                   End If
273                               End If
274                           End If
275                           End If
276                       End If
                      
277               ElseIf idx = 2 Then 'Not cekGrid(oRs2, BarisPart, frm, kodeGrid) And
278                   If kodeGrid = 1 Then
279                       For i = 1 To barispart - 1
280                           With frm
281                               If Trim(.partGrid.TextMatrix(i, 1)) <> "" Then
282                                   If Mid(.partGrid.TextMatrix(i, 1), 3, Len(.partGrid.TextMatrix(i, 1)) - 2) = Trim(oRs2("ITEM_PART")) Then
283                                           Set rsHrg = objO.FetchDataACC("SELECT hargajuala FROM inmst WITH (NOLOCK) WHERE kodeItem like '%" & Trim(.partGrid.TextMatrix(i, 1)) & "'")
284                                           If Not rsHrg.EOF Then
285                                               hrg = CDbl(rsHrg("HARGAJUALA"))
286                                           End If
287                                           Call setClose(rsHrg)
                                              
                                              ''Set rsHrg = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(.partGrid.TextMatrix(i, 1), 1) & "' AND TIPE = 'O'")
288                                           Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("ITEM_PART"))) & "' and B.Jenis = 'OLI'")
289                                           If Not rsHrg.EOF Then
290                                               oliPart = True
291                                           Else
292                                               oliPart = False
293                                           End If
294                                       If .Name = "frmBooking" Then
295                                           If oliPart Then
296                                               .Text(55) = CDbl(.Text(55)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
297                                               .Text(55) = Format(.Text(55), "#0,0")
298                                           Else
299                                               .Text(40) = CDbl(.Text(40)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
300                                               .Text(40) = Format(.Text(40), "#0,0")
301                                           End If
                                              
302                                       ElseIf .Name = "frmEstimasi" Then
303                                           If oliPart Then
304                                               .Text(28) = CDbl(.Text(28)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
305                                               .Text(28) = Format(.Text(28), "#0,0")
306                                           Else
307                                               .Text(24) = CDbl(.Text(24)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
308                                               .Text(24) = Format(.Text(24), "#0,0")
309                                           End If
                                              
310                                       End If
                                          
311                                       nC = i
312                                       For j = nC To barispart
313                                           With .partGrid
314                                               .TextMatrix(j, 0) = .TextMatrix(j + 1, 0)
315                                               .TextMatrix(j, 1) = .TextMatrix(j + 1, 1)
316                                               .TextMatrix(j, 2) = .TextMatrix(j + 1, 2)
317                                               .TextMatrix(j, 3) = .TextMatrix(j + 1, 3)
318                                               .TextMatrix(j, 4) = .TextMatrix(j + 1, 4)
319                                               .TextMatrix(j, 5) = .TextMatrix(j + 1, 5)
320                                               .TextMatrix(j, 6) = .TextMatrix(j + 1, 6)
321                                               .TextMatrix(j, 7) = .TextMatrix(j + 1, 7)
322                                               .TextMatrix(j, 8) = .TextMatrix(j + 1, 8)
323                                           End With
324                                       Next
325                                       barispart = barispart - 1
326                                   End If
327                               ElseIf Trim(.partGrid.TextMatrix(i, IIf(.Name = "frmEstimasi", 2, 7))) <> "" Then
328                                   If (Trim(.partGrid.TextMatrix(i, IIf(.Name = "frmEstimasi", 2, 7))) = Trim(oRs2("NAMA_PART"))) Then
                                              'Set rsHrg = objO.FetchDataACC("SELECT hargajuala FROM inmst WHERE kodeItem = '" & Trim(.partGrid.TextMatrix(i, 1)) & "'")
                                              'If Not rsHrg.EOF Then
329                                               hrg = CDbl(.partGrid.TextMatrix(i, 5))
                                              'End If
                                              'Call setClose(rsHrg)
                                              
                                              'Set rsHrg = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WHERE FS_KD_OLI = '" & Left(.partGrid.TextMatrix(i, 1), 1) & "' AND TIPE = 'O'")
                                              'If Not rsHrg.EOF Then
                                              '    oliPart = True
                                              'Else
330                                               oliPart = False
                                              'End If
331                                       If .Name = "frmBooking" Then
332                                           If oliPart Then
333                                               .Text(55) = CDbl(.Text(55)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
334                                               .Text(55) = Format(.Text(55), "#0,0")
335                                           Else
336                                               .Text(40) = CDbl(.Text(40)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
337                                               .Text(40) = Format(.Text(40), "#0,0")
338                                           End If
                                              
339                                       ElseIf .Name = "frmEstimasi" Then
340                                           If oliPart Then
341                                               .Text(28) = CDbl(.Text(28)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
342                                               .Text(28) = Format(.Text(28), "#0,0")
343                                           Else
344                                               .Text(24) = CDbl(.Text(24)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
345                                               .Text(24) = Format(.Text(24), "#0,0")
346                                           End If
                                              
347                                       End If
                                          
348                                       nC = i
349                                       For j = nC To barispart
350                                           With .partGrid
351                                               .TextMatrix(j, 0) = .TextMatrix(j + 1, 0)
352                                               .TextMatrix(j, 1) = .TextMatrix(j + 1, 1)
353                                               .TextMatrix(j, 2) = .TextMatrix(j + 1, 2)
354                                               .TextMatrix(j, 3) = .TextMatrix(j + 1, 3)
355                                               .TextMatrix(j, 4) = .TextMatrix(j + 1, 4)
356                                               .TextMatrix(j, 5) = .TextMatrix(j + 1, 5)
357                                               .TextMatrix(j, 6) = .TextMatrix(j + 1, 6)
358                                               .TextMatrix(j, 7) = .TextMatrix(j + 1, 7)
359                                               .TextMatrix(j, 8) = .TextMatrix(j + 1, 8)
360                                           End With
361                                       Next
362                                       barispart = barispart - 1
363                                   End If

364                               ElseIf Trim(.partGrid.TextMatrix(i, 2)) <> "" Then
                                              
365                                           hrg = .partGrid.TextMatrix(i, 5)
366                                           oliPart = False
                                              
367                                       If .Name = "frmBooking" Then
368                                           If oliPart Then
369                                               .Text(55) = CDbl(.Text(55)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
370                                               .Text(55) = Format(.Text(55), "#0,0")
371                                           Else
372                                               .Text(40) = CDbl(.Text(40)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
373                                               .Text(40) = Format(.Text(40), "#0,0")
374                                           End If
                                              
375                                       ElseIf .Name = "frmEstimasi" Then
376                                           If oliPart Then
377                                               .Text(28) = CDbl(.Text(28)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
378                                               .Text(28) = Format(.Text(28), "#0,0")
379                                           Else
380                                               .Text(24) = CDbl(.Text(24)) - (CDbl(.partGrid.TextMatrix(i, 4)) * CDbl(hrg))
381                                               .Text(24) = Format(.Text(24), "#0,0")
382                                           End If
                                              
383                                       End If
                                          
384                                       nC = i
385                                       For j = nC To barispart
386                                           With .partGrid
387                                               .TextMatrix(j, 0) = .TextMatrix(j + 1, 0)
388                                               .TextMatrix(j, 1) = .TextMatrix(j + 1, 1)
389                                               .TextMatrix(j, 2) = .TextMatrix(j + 1, 2)
390                                               .TextMatrix(j, 3) = .TextMatrix(j + 1, 3)
391                                               .TextMatrix(j, 4) = .TextMatrix(j + 1, 4)
392                                               .TextMatrix(j, 5) = .TextMatrix(j + 1, 5)
393                                               .TextMatrix(j, 6) = .TextMatrix(j + 1, 6)
394                                               .TextMatrix(j, 7) = .TextMatrix(j + 1, 7)
395                                               .TextMatrix(j, 8) = .TextMatrix(j + 1, 8)
396                                           End With
397                                       Next
398                                       barispart = barispart - 1

399                               End If
400                           End With
401                       Next
402                   ElseIf kodeGrid = 2 Then
403                       With rsS
404                           .Find "KODEITEM LIKE '%" & Trim(oRs2("ITEM_PART")) & "%'", , adSearchForward, 1
405                           If Not .EOF Then
406                                   Set rsHrg = objO.FetchDataACC("SELECT hargajuala FROM inmst WITH (NOLOCK) WHERE kodeItem like '%" & rsS("KODEITEM") & "'")
407                                   If Not rsHrg.EOF Then
408                                       hrg = CDbl(rsHrg("HARGAJUALA"))
409                                   End If
410                                   rsHrg.Close
                                      
                                      ''Set rsHrg = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsS("KODEITEM"), 1) & "' AND TIPE = 'O'")
411                                   Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("ITEM_PART"))) & "' and B.Jenis = 'OLI'")
412                                   If Not rsHrg.EOF Then
413                                       oliPart = True
414                                   Else
415                                       oliPart = False
416                                   End If
                                  
417                               With frm
418                                   If oliPart Then
419                                       .Text(59) = CDbl(.Text(59)) - (CDbl(rsS("QTY")) * CDbl(hrg))
420                                       .Text(59) = Format(.Text(59), "#0,0")
421                                   Else
422                                       .Text(57) = CDbl(.Text(57)) - (CDbl(rsS("QTY")) * CDbl(hrg))
423                                       .Text(57) = Format(.Text(57), "#0,0")
424                                   End If
425                               End With
426                               .Delete
427                               .Update
428                           Else
                              
429                               .Find "NAMABYSA = '" & Trim(oRs2("NAMA_PART")) & "'", , adSearchForward, 1
430                               If Not .EOF Then
                                          'Set rsHrg = objO.FetchDataACC("SELECT hargajuala FROM inmst WHERE NamaBarang = '" & rsS("NAMABYSA") & "'")
                                          'If Not rsHrg.EOF Then
431                                           hrg = 0 'CDbl(.Fields(1).value)
                                          'End If
                                          'rsHrg.Close
                                          
                                          'Set rsHrg = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WHERE FS_KD_OLI = '" & Left(rsS("KODEITEM"), 1) & "' AND TIPE = 'O'")
                                          'If Not rsHrg.EOF Then
                                          '    oliPart = True
                                          'Else
432                                           oliPart = False
                                          'End If
                                      
433                                   With frm
434                                       If oliPart Then
435                                           .Text(59) = CDbl(.Text(59)) - (CDbl(rsS("QTY")) * CDbl(hrg))
436                                           .Text(59) = Format(.Text(59), "#0,0")
437                                       Else
438                                           .Text(57) = CDbl(.Text(57)) - (CDbl(rsS("QTY")) * CDbl(hrg))
439                                           .Text(57) = Format(.Text(57), "#0,0")
440                                       End If
441                                   End With
442                                   .Delete
443                                   .Update
444                               End If
445                           End If
446                       End With
447                   End If
448               End If
449               oRs2.MoveNext
450           Loop
451       End If
452       If oRs2.State <> adStateClosed Then oRs2.Close
End Function

Function cekGrid(rsT As ADODB.Recordset, barispart As Integer, frmP As Object, kdGrid As Integer, Optional rsS As Recordset) As Boolean
Dim i As Integer

    cekGrid = True
    If kdGrid = 1 Then
        For i = 1 To barispart
            With frmP
                If .partGrid.TextMatrix(i, 1) = Trim(rsT("ITEM_PART")) Or .partGrid.TextMatrix(i, 7) = Trim(rsT("Nama_PART")) Then
                    cekGrid = False
                End If
            End With
        Next
    ElseIf kdGrid = 2 Then
        With rsS
            If .RecordCount > 1 Then
                .MoveFirst
                Do While Not .EOF
                    If rsS!KodeItem <> "" Then
                        If Mid(rsS!KodeItem, 3, Len(rsS!KodeItem) - 2) = Trim(rsT("ITEM_PART")) Then
                            cekGrid = False
                        End If
                    Else
                        cekGrid = True
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
End Function


Function cekGridFOP(rsT As ADODB.Recordset, barispart As Integer, frmP As Object, kdGrid As Integer, Optional rsS As Recordset) As Boolean
Dim i As Integer

    cekGridFOP = True
    If kdGrid = 1 Then
        For i = 1 To barispart
            With frmP
                If .partGrid.TextMatrix(i, 1) = Trim(rsT("ITEM_PART")) Then
                    cekGridFOP = False
                End If
            End With
        Next
    ElseIf kdGrid = 2 Then
        With rsS
            If .RecordCount > 1 Then
                .MoveFirst
                Do While Not .EOF
                    If rsS!KodeItem <> "" Then
                        If Mid(rsS!KodeItem, 3, Len(rsS!KodeItem) - 2) = Trim(rsT("ITEM_PART")) Then
                            cekGridFOP = False
                        End If
                    Else
                        cekGridFOP = True
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
End Function


Function cekMOU(kMdl As String, nRgk As String, Optional tipeService As String) As String
Dim rsDealer As ADODB.Recordset
Dim rsSearch As ADODB.Recordset

    cekMOU = ""
    If Trim(tipeService) <> "W" Then
        Set oRs1 = support.listMOU(kMdl, nRgk)
        If Not oRs1.EOF Then
            If oRs1("FLAGDEALER") = "A" Then
                cekMOU = Trim(oRs1("NO_MOU"))
            Else
                Set rsDealer = support.listDealerMOU(Trim(oRs1("NO_MOU")), gKdCbg) 'oConn.Execute("SELECT * FROM MOUCBG WHERE NO_MOU = '" & oRs1("NO_MOU") & "' AND KODEDEALER = '" & gKdCbg & "'")
                If Not rsDealer.EOF Then
                    cekMOU = Trim(oRs1("NO_MOU"))
                End If
            End If
        End If
    End If
End Function

Function capailimitMOU(limitPercent As Integer, TipeMdl As String, noChassis As String) As Boolean
Dim noMou As String
Dim rsSearch As ADODB.Recordset
    
    capailimitMOU = False
    noMou = Trim(cekMOU(Trim(TipeMdl), Trim(noChassis)))
    If noMou <> "" Then
        
        If Log.RPLHistori = 0 Then
            Set rsSearch = objO.FetchData("SELECT LIMITAR-HUTANG+BAYAR AS SISA, limitar FROM MOUHDR WITH (NOLOCK) WHERE NO_MOU = '" & noMou & "'")
        Else
            Set rsSearch = objO.FetchData("SELECT LIMITAR-HUTANG+BAYAR AS SISA, limitar FROM [" & dbReplikasiHO & "]..MOUHDR WITH (NOLOCK) WHERE NO_MOU = '" & noMou & "'")
        End If
        If Not rsSearch.EOF Then
            If CDbl(rsSearch("SISA")) < (CDbl(rsSearch("LIMITAR")) * 0.01 * limitPercent) Then
                capailimitMOU = True
            End If
        End If
    End If
End Function

Function cekPiutang(nKdPlg As String, kMdl As String, nRgk As String) As Boolean
      Dim rsMOU As ADODB.Recordset

1         cekPiutang = False
2             Set oRs1 = objO.FetchDataACC("SELECT kodelgn, sum(jumlahtrn) as totAR FROM artrn WITH (NOLOCK) WHERE kodeLgn = '" & nKdPlg & "' AND LEFT(NOFAKTUR,1) IN ('E','D') GROUP BY kodeLgn ORDER BY kodeLgn")
3             If Not oRs1.EOF Then
4                 If CDbl(oRs1("TOTAR")) > 0 Then
                      ' ================================================================================================================================
                      ' 01 Mar 2017 | Thania | buka kuncian pencarian mou berdasarkan model hanya berdasarkan noka (req. 13 Feb 2017 - Setyo)
                      ' ================================================================================================================================
                      'Set rsMOU = objO.FetchData("SELECT isnull(limitar,0) as MaxAr, * FROM mouhdr A WITH (NOLOCK) INNER JOIN moudtl B WITH (NOLOCK) ON A.no_mou = B.no_mou and A.FS_LOG_DEL = '' WHERE mdl = '" & kMdl & "' AND no_Rangka = '" & nRgk & "' ORDER BY mdl, no_Rangka")
                      
                      
5                     If Log.RPLHistori = 0 Then
6                         Set rsMOU = objO.FetchData("SELECT isnull(limitar,0) as MaxAr, * FROM mouhdr A WITH (NOLOCK) INNER JOIN moudtl B WITH (NOLOCK) ON A.no_mou = B.no_mou and A.FS_LOG_DEL = '' WHERE no_Rangka = '" & nRgk & "' ORDER BY mdl, no_Rangka")
7                     Else
8                         Set rsMOU = objO.FetchData("SELECT isnull(limitar,0) as MaxAr, * FROM [" & dbReplikasiHO & "]..mouhdr A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..moudtl B WITH (NOLOCK) ON A.no_mou = B.no_mou and A.FS_LOG_DEL = '' WHERE no_Rangka = '" & nRgk & "' ORDER BY mdl, no_Rangka")
9                     End If
                      
                      ' ================================================================================================================================
10                    If Not rsMOU.EOF Then
11                        If CDbl(oRs1("TOTAR")) >= CDbl(rsMOU("MAXAR")) Then
12                            MsgBox "Masih ada AR sejumlah = Rp. " & Format(CDbl(oRs1("TOTAR")), "##,0") & vbCr & "Mohon untuk dilunasi sebelum hari H service.", vbInformation, "Piutang"
13                            cekPiutang = True
14                        End If
15                    End If
16                End If
17            End If
          
End Function

Function hitungLabour(nRt As Double, klas As String, Optional warranty As Boolean = False) As Double
    If oRs1.State Then oRs1.Close
    If warranty Then
        If Log.RPLHistori = 1 Then
            Set oRs1 = objO.FetchData("SELECT PRICE FROM [" & dbReplikasiHO & "]..carclass WITH (NOLOCK) WHERE war = '1' and FS_LOG_DEL = ''")
        Else
            Set oRs1 = objO.FetchData("SELECT PRICE FROM carclass WITH (NOLOCK) WHERE war = '1' and FS_LOG_DEL = ''")
        End If
        ''Set oRs1 = objO.FetchData("SELECT PRICE FROM carclass WITH (NOLOCK) WHERE war = '1' and FS_LOG_DEL = ''")
        If Not oRs1.EOF Then
            hitungLabour = CDbl(Round((oRs1("PRICE") * nRt), 0))
        End If
    Else
        If Log.RPLHistori = 1 Then
            Set oRs1 = objO.FetchData("SELECT PRICE FROM [" & dbReplikasiHO & "]..carclass WITH (NOLOCK) WHERE classcode = '" & Trim(klas) & "' and FS_LOG_DEL = ''")
        Else
            Set oRs1 = objO.FetchData("SELECT PRICE FROM carclass WITH (NOLOCK) WHERE classcode = '" & Trim(klas) & "' and FS_LOG_DEL = ''")
        End If
        ''Set oRs1 = objO.FetchData("SELECT PRICE FROM carclass WITH (NOLOCK) WHERE classcode = '" & Trim(klas) & "' and FS_LOG_DEL = ''")
        If Not oRs1.EOF Then
            If CexChassis_Discon(klas) = True Then
                hitungLabour = CDbl(Round(((oRs1("PRICE") * ((100 - nDiscNRM) / 100)) * nRt), 0))
            Else
                hitungLabour = CDbl(Round(((oRs1("PRICE") * ((100 - nDiscNON) / 100)) * nRt), 0))
            End If
        End If
    End If
End Function

Function hitungPart(nQty As Integer, nHrg As Double) As Double
    hitungPart = CDbl(nQty) * CDbl(nHrg)
End Function

Function hitungWaktu(nRt As Double, jamMulai As String, TglMulai As String, Optional sts As String) As String
      Dim jM As Integer   'Waktu Mulai
      Dim jS As Integer
      Dim mM As Integer   'Waktu Selesai
      Dim mS As Integer
      Dim jam As Integer  'Rate
      Dim menit As Integer
      Dim tglHsl As String

1        On Error GoTo hitungWaktu_Error

2         jM = 0
3         jS = 0
4         mM = 0
5         mS = 0
          
6         mM = (Left(jamMulai, InStr(jamMulai, ":") - 1) * 60) + Mid(jamMulai, InStr(jamMulai, ":") + 1, Len(jamMulai) - InStr(jamMulai, ":"))
7         jam = CDbl(nRt) * 60
          
8         If sts = "progress" Then
9             mS = (mM + jam) Mod 60 + (10 - Right(Format((mM + jam) Mod 60, "00"), 1))
10        Else
11            mS = (mM + jam) Mod 60
12        End If
13        jS = Val((mM + jam) / 60)
14        If mS >= 60 Then
15            jS = jS + 1
16            mS = mS - 60
17        End If
18        tglHsl = TglMulai
19        If jS >= 24 Then
20            If sts = "progress" Then
21                jS = 23
22                mS = 59
23            Else
24                jS = jS - 24
25            End If
26            tglHsl = TglDDMMYYYY(DateAdd("d", 1, Format(TglMulai, "dd/mm/yyyy")))
27        End If
          
28        hitungWaktu = Format(jS, "00") & ":" & Format(mS, "00") & "|" & tglHsl

29       On Error GoTo 0
30       Exit Function

hitungWaktu_Error:

31        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hitungWaktu of Module MainModul")
End Function

Function hitungTotalPenyerahan(total As Integer, jamMulai As String, TglMulai As String) As String
Dim jM As Integer   'Waktu Mulai
Dim jS As Integer
Dim mM As Integer   'Waktu Selesai
Dim mS As Integer
Dim jam As Integer  'Rate
Dim menit As Integer
Dim tglHsl As String

    jM = 0
    jS = 0
    mM = 0
    mS = 0
    
    mM = (Left(jamMulai, InStr(jamMulai, ":") - 1) * 60) + Mid(jamMulai, InStr(jamMulai, ":") + 1, Len(jamMulai) - InStr(jamMulai, ":"))
    jam = total
    
    mS = (mM + jam) Mod 60
    
    jS = Val((mM + jam) / 60)
    If mS >= 60 Then
        jS = jS + 1
        mS = mS - 60
    End If
    tglHsl = TglMulai
    If jS >= 24 Then
        jS = jS - 24
        tglHsl = TglDDMMYYYY(DateAdd("d", 1, Format(TglMulai, "dd/mm/yyyy")))
    End If
    
    hitungTotalPenyerahan = Format(jS, "00") & ":" & Format(mS, "00") & "|" & tglHsl
End Function


Function hitungWaktuJamKerja(nRt As Double, jamMulai As String, TglMulai As String) As String
Dim jM As Integer   'Waktu Mulai
Dim jS As Integer
Dim mM As Integer   'Waktu Selesai
Dim mS As Integer
Dim jam As Integer  'Rate
Dim menit As Integer
Dim tglHsl As String
Dim JamKerja As String
Dim JamKerjaMulai As String
Dim JamKerjaSelesai As String
Dim mMulai As Integer
Dim mSelesai As Integer

    jM = 0
    jS = 0
    mM = 0
    mS = 0
    
    JamKerja = cekJamKerja(Trim(IIf(Weekday(TglMulai) = 1, 6, Weekday(TglMulai) - 2) & WeekdayName(IIf(Weekday(TglMulai) = 1, 8, Weekday(TglMulai)) - 1)))
    
    JamKerjaMulai = Left(JamKerja, InStr(JamKerja, "|") - 1)
    mMulai = (Left(JamKerjaMulai, InStr(JamKerjaMulai, ":") - 1) * 60) + Mid(JamKerjaMulai, InStr(JamKerjaMulai, ":") + 1, Len(JamKerjaMulai) - InStr(JamKerjaMulai, ":"))
    
    JamKerjaSelesai = Mid(JamKerja, InStr(JamKerja, "|") + 1, Len(JamKerja) - InStr(JamKerja, "|"))
    mSelesai = (Left(JamKerjaSelesai, InStr(JamKerjaSelesai, ":") - 1) * 60) + Mid(JamKerjaSelesai, InStr(JamKerjaSelesai, ":") + 1, Len(JamKerjaSelesai) - InStr(JamKerjaSelesai, ":"))
    
    mM = (Left(jamMulai, InStr(jamMulai, ":") - 1) * 60) + Mid(jamMulai, InStr(jamMulai, ":") + 1, Len(jamMulai) - InStr(jamMulai, ":"))
    jam = CDbl(nRt) * 60
    
    mS = (mM + jam + (1440 - mSelesai + mMulai)) Mod 60
    jS = Val((mM + jam + (1440 - mSelesai + mMulai)) / 60)
    If mS >= 60 Then
        jS = jS + 1
        mS = mS - 60
    End If
    tglHsl = TglMulai
    If jS >= 24 Then
        jS = jS - 24
        tglHsl = TglDDMMYYYY(DateAdd("d", 1, Format(TglMulai, "dd/mm/yyyy")))
    End If
    
    hitungWaktuJamKerja = Format(jS, "00") & ":" & Format(mS, "00") & "|" & tglHsl
End Function

Function cekJamKerja(nHari As String) As String
Dim rsJam As ADODB.Recordset
    
    cekJamKerja = "|"
    Set rsJam = objO.FetchData("SELECT WAKTUMULAI,WAKTUSELESAI FROM T_STALL WITH (NOLOCK) WHERE HARI = '" & Trim(nHari) & "'")
    If Not rsJam.EOF Then
        cekJamKerja = Trim(rsJam("WAKTUMULAI")) & "|" & Trim(rsJam("WAKTUSELESAI"))
    End If
End Function

Function hitungWaktuMundur(nRt As Double, jamMulai As String, TglMulai As String, Optional sts As String) As String
Dim jM As Integer   'Waktu Mulai
Dim jS As Integer
Dim mM As Integer   'Waktu Selesai
Dim mS As Integer
Dim jam As Integer  'Rate
Dim menit As Integer
Dim tglHsl As String
Dim JamKerja As String
Dim JamKerjaMulai As String
Dim JamKerjaSelesai As String
Dim mMulai As Integer
Dim mSelesai As Integer

    jM = 0
    jS = 0
    mM = 0
    mS = 0
    
    JamKerja = cekJamKerja(Trim(IIf(Weekday(TglMulai) = 1, 6, Weekday(TglMulai) - 2) & WeekdayName(IIf(Weekday(TglMulai) = 1, 8, Weekday(TglMulai)) - 1)))
    
    JamKerjaMulai = Left(JamKerja, InStr(JamKerja, "|") - 1)
    mMulai = (Left(JamKerjaMulai, InStr(JamKerjaMulai, ":") - 1) * 60) + Mid(JamKerjaMulai, InStr(JamKerjaMulai, ":") + 1, Len(JamKerjaMulai) - InStr(JamKerjaMulai, ":"))
    
    JamKerjaSelesai = Mid(JamKerja, InStr(JamKerja, "|") + 1, Len(JamKerja) - InStr(JamKerja, "|"))
    mSelesai = (Left(JamKerjaSelesai, InStr(JamKerjaSelesai, ":") - 1) * 60) + Mid(JamKerjaSelesai, InStr(JamKerjaSelesai, ":") + 1, Len(JamKerjaSelesai) - InStr(JamKerjaSelesai, ":"))
    
    mM = (Left(jamMulai, InStr(jamMulai, ":") - 1) * 60) + Mid(jamMulai, InStr(jamMulai, ":") + 1, Len(jamMulai) - InStr(jamMulai, ":"))
    jam = CDbl(nRt) * 60
    
    If sts = "progress" Then
        
    Else
        If CDbl(mM - jam - (1440 - mSelesai + mMulai)) < 0 Then
            mM = mM + (24 * 60)
            jS = 24
        End If
        mS = (mM - jam - (1440 - mSelesai + mMulai)) Mod 60
    End If
    jS = jS + Val((mM - jam - (1440 - mSelesai + mMulai)) / 60)
    If mS >= 60 Then
        jS = jS + 1
        mS = mS - 60
    End If
    tglHsl = TglMulai
    If jS >= 24 Then
        If sts = "progress" Then
            jS = 23
            mS = 59
        Else
            jS = jS - 24
        End If
        tglHsl = TglDDMMYYYY(DateAdd("d", -1, Format(TglMulai, "dd/mm/yyyy")))
    End If
    
    hitungWaktuMundur = Format(jS, "00") & ":" & Format(mS, "00") & "|" & tglHsl
End Function


Function hitMenit(jamMulai As String, jamSelesai As String) As Integer
Dim jM As Integer
Dim mM As Integer
Dim jS As Integer
Dim mS As Integer

    If jamSelesai < jamMulai Then
        jamSelesai = (24 + CDbl(Left(jamSelesai, InStr(jamSelesai, ":") - 1))) & ":" & Right(jamSelesai, InStr(jamSelesai, ":") - 1)
    End If
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

Function hitungTotalMenit(jam As String) As Integer
    Dim jM As Integer
    Dim mM As Integer
    
    jM = Left(jam, InStr(jam, ":") - 1) * 60
    mM = Right(jam, InStr(jam, ":") - 1)
    
    hitungTotalMenit = jM + mM
End Function

Function hitungRateMenit(jam As String) As Double
    Dim jM As Integer
    Dim mM As Double
    
    jM = Left(jam, InStr(jam, ":") - 1)
    mM = Right(jam, InStr(jam, ":") - 1) / 60
    
    hitungRateMenit = Round(jM + mM, 2)
End Function

Function cekMRS(nRgk As String, nKm As Double, fRC As String, tipeJb As Integer, Optional rsJb As ADODB.Recordset, Optional TheGrid As String) As String
          Dim rsHlp As ADODB.Recordset
      Dim rsHlp1 As ADODB.Recordset
      Dim gSBE As Boolean

1         If fRC = "" Then
2             If tipeJb = 1 Then
3                 rsJb.MoveFirst
4                 Do While Not rsJb.EOF
5                     If Log.RPLHistori = 0 Then
6                         Set rsHlp = objO.FetchData("SELECT SBE FROM jobOrder WITH (NOLOCK) WHERE fs_item = '" & rsJb("JOB") & "' and FS_LOG_DEL = ''")
7                     Else
8                         Set rsHlp = objO.FetchData("SELECT SBE FROM [" & dbReplikasiHO & "]..jobOrder WITH (NOLOCK) WHERE fs_item = '" & rsJb("JOB") & "' and FS_LOG_DEL = ''")
9                     End If
10                    If Not rsHlp.EOF Then
11                        gSBE = IIf(rsHlp("SBE") = "Y", True, False)
12                        Exit Do
13                    Else
14                        gSBE = False
15                    End If
16                    rsJb.MoveNext
17                Loop
18            ElseIf tipeJb = 2 Then
                  Dim i As Integer
19                For i = 1 To 7
20                  '  If Log.RPLHistori = 0 Then
21                        Set rsHlp = objO.FetchData("SELECT SBE FROM ACTRATE WITH (NOLOCK) WHERE fs_item = '" & TheGrid & "' AND MDL IN (SELECT MDL FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(nRgk) & "') and FS_LOG_DEL = ''")
22                   ' Else
23                   '     Set rsHlp = objO.FetchData("SELECT SBE FROM [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) WHERE fs_item = '" & TheGrid & "' AND MDL IN (SELECT MDL FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(nRgk) & "') and FS_LOG_DEL = ''")
24                   ' End If
25                    If Not rsHlp.EOF Then
26                        gSBE = IIf(rsHlp("SBE") = "Y", True, False)
27                        Exit For
28                    Else
29                        gSBE = False
30                    End If
31                    rsHlp.Close
32                Next
33            End If
              
34            If gSBE Then
35                Set rsHlp1 = objO.FetchData("SELECT noRangka FROM M_ESTIMASI WITH (NOLOCK) WHERE noRangka = '" & nRgk & "' " _
                                           & "AND JT_TGL BETWEEN DATEADD(d, -15,GETDATE()) AND DATEADD(d, 15,GETDATE()) " _
                                           & "AND KMnext BETWEEN '" & nKm - 1500 & "' AND '" & nKm + 1500 & "' ORDER BY JT_Tgl DESC")
36                If Not rsHlp1.EOF Then
37                    cekMRS = "GRP"
38                Else
39                    cekMRS = ""
40                End If
41                rsHlp1.Close
42            Else
43                cekMRS = ""
44            End If
45        Else
46            cekMRS = "SBE"
47        End If
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


Function cekMekanik(nMekanik As String, nRte As Double, nTgl As String, nMnt As String) As Boolean
Dim rsHlp As New ADODB.Recordset
    
    cekMekanik = False
     
    If DateDiff("d", Format(nTgl, "MM/DD/YYYY"), Format(isnow, "MM/DD/YYYY")) = 0 Then
        sSQL = "SELECT A.MECH FROM MEKANIK A WITH (NOLOCK) INNER JOIN T_MEKANIK B WITH (NOLOCK) ON A.MECH = B.MECH " _
             & "WHERE TANGGAL = '" & nTgl & "' AND A.MECH = '" & nMekanik & "' AND JAMMULAI <= '" & nMnt & "' AND JAMSELESAi >= '" & nMnt & "' " _
             & "AND (MENITKERJA-TERPAKAI) > 0 AND (MENITKERJA-TERPAKAI) > " & CDbl(nRte) * 60 & ""
    Else
        sSQL = "SELECT MECH FROM MEKANIK WITH (NOLOCK) WHERE MECH = '" & nMekanik & "'"
    End If
    
    Set rsHlp = objO.FetchData(sSQL)
    If Not rsHlp.EOF Then
        cekMekanik = True
    End If
    If Not cekMekanik Then gBantuan3 = vbCr & vbCr & "Mekanik " & nMekanik & " tidak tersedia pada Tanggal " & nTgl & " ,Jam " & nMnt
    rsHlp.Close
End Function

Function cekDouble(Grd As Object, jOrd As String, drt As Integer, idx As Integer) As Boolean
Dim i As Integer

    cekDouble = False
    With Grd
        If idx = 1 Then
            For i = 1 To drt - 1
                If Trim(.TextMatrix(i, 1)) = Trim(jOrd) Then
                    cekDouble = True
                    gBantuan = i
                    Exit Function
                End If
            Next
        ElseIf idx = 2 Then
            For i = 1 To drt - 1
                If Trim(Mid(.TextMatrix(i, 1), 3, Len(.TextMatrix(i, 1)) - 2)) = Trim(jOrd) Then
                    cekDouble = True
                    gBantuan = i
                    Exit Function
                End If
            Next
        End If
    End With
End Function

Function hitungRate(waktuAwal As String, waktuAkhir As String) As Integer
Dim jA As Integer
Dim jS As Integer
Dim mA As Integer
Dim mS As Integer
Dim dA As Integer
Dim dS As Integer
    
    jA = 0
    jS = 0
    mA = 0
    mS = 0
    jA = Left(waktuAwal, InStr(waktuAwal, ":") - 1)
    jS = Left(waktuAkhir, InStr(waktuAkhir, ":") - 1)
    mA = Right(waktuAwal, InStr(waktuAwal, ":") - 1)
    mS = Right(waktuAkhir, InStr(waktuAkhir, ":") - 1)
    
    dA = (jA * 60) + mA
    dS = (jS * 60) + mS
    
    hitungRate = dS - dA
End Function

Function pecahNoPol(nPol As String, tDpn As String, tTgh As String, tBlk As String, Optional kar As String = "-")
    
On Error GoTo nextProses
    nPol = Replace(nPol, " ", "-")
    If Not IsNumeric(Left(nPol, 1)) Then
        ' KARAKTER PERTAMA HURUF
        tDpn = Left(nPol, InStr(nPol, kar) - 1)
        If cariKarakter(nPol, kar) > 1 Then
            tTgh = Mid(nPol, InStr(nPol, kar) + 1, InStr((Right(nPol, Len(nPol) - InStr(nPol, kar))), kar) - 1)
            tBlk = Mid(Right(nPol, Len(nPol) - InStr(nPol, kar)), InStr(Right(nPol, Len(nPol) - InStr(nPol, kar)), kar) + 1, Len(Right(nPol, Len(nPol) - InStr(nPol, kar))) - InStr(Right(nPol, Len(nPol) - InStr(nPol, kar)), kar))
        Else
            tTgh = Mid(nPol, InStr(nPol, kar) + 1, Len(nPol) - InStr(nPol, kar))
            tBlk = ""
        End If
    Else
        ' KARAKTER PERTAMA ANGKA
        tDpn = ""
        tTgh = Left(nPol, IIf(InStr(nPol, kar) > 0, InStr(nPol, kar) - 1, Len(nPol)))
        tBlk = Mid(nPol, InStr(nPol, kar) + 1, Len(nPol) - InStr(nPol, kar))
    End If
    Exit Function

nextProses:
    tDpn = ""
    tTgh = ""
    tBlk = ""
End Function

Function pecahNoTelp(nPol As String, tDpn As String, tTgh As String, tBlk As String, Optional kar As String = "-")
    
    If InStr(nPol, kar) > 0 Then
        tDpn = Left(nPol, InStr(nPol, kar) - 1)
        If cariKarakter(nPol, kar) > 1 Then
            tTgh = Mid(nPol, InStr(nPol, kar) + 1, InStr((Right(nPol, Len(nPol) - InStr(nPol, kar))), kar) - 1)
            tBlk = Mid(Right(nPol, Len(nPol) - InStr(nPol, kar)), InStr(Right(nPol, Len(nPol) - InStr(nPol, kar)), kar) + 1, Len(Right(nPol, Len(nPol) - InStr(nPol, kar))) - InStr(Right(nPol, Len(nPol) - InStr(nPol, kar)), kar))
        Else
            tTgh = Mid(nPol, InStr(nPol, kar) + 1, Len(nPol) - InStr(nPol, kar))
            tBlk = ""
        End If
    Else
        tDpn = Left(nPol, 10)
        tTgh = ""
        tBlk = ""
    End If
End Function

Function cariKarakter(ntxt As String, nChar As String) As Integer
Dim nPanjang As Integer
Dim i As Integer
    
    cariKarakter = 0
    nPanjang = Len(ntxt)
    For i = 1 To nPanjang
        If Trim(Mid(ntxt, i, 1)) = nChar Then
            cariKarakter = cariKarakter + 1
        End If
    Next
End Function

Function hitungMenit(nJam As String) As Integer
1        On Error GoTo hitungMenit_Error

2         hitungMenit = (CDbl(Left(nJam, InStr(nJam, ":") - 1)) * 60) + CDbl(Mid(nJam, InStr(nJam, ":") + 1, (Len(nJam) - InStr(nJam, ":"))))

3        On Error GoTo 0
4        Exit Function

hitungMenit_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hitungMenit of Module MainModul")
End Function

Public Function countRestTime(mechCode As String, divTime As Double, countDay As Double, nDate As Date)
Dim rsTimeDay As ADODB.Recordset
Dim rsHlp As ADODB.Recordset
Dim nameDay As String
Dim timeS As Double
Dim timeP As Double
Dim newAbsent As Boolean
Dim strSql As String

    On Error GoTo errHandler
    newAbsent = False
    nameDay = IIf(Weekday(nDate) = 1, 6, Weekday(nDate) - 2) & WeekdayName(IIf(Weekday(nDate) = 1, 8, Weekday(nDate)) - 1)
    If LoopMk <> True Then
        Set rsHlp = objO.FetchData_MEKANIK("SELECT * FROM T_MEKANIK WITH (NOLOCK) WHERE TANGGAL = '" & nDate & "' AND MECH = '" & mechCode & "'")
    Else
        Set rsHlp = rsMK
    End If
    If rsHlp.EOF Then
        newAbsent = True
        Set rsTimeDay = objO.FetchData("SELECT WAKTUMULAI, WAKTUSELESAI FROM T_STALL WITH (NOLOCK) WHERE HARI = '" & nameDay & "'")
        If Not rsTimeDay.EOF Then
            timeS = CDbl(hitungRate(Trim(rsTimeDay("WAKTUMULAI")), Trim(rsTimeDay("WAKTUSELESAI"))))
        End If
        timeP = 0
    Else
        newAbsent = False
        timeS = CDbl(rsHlp("MENITKERJA"))
        If timeS = 0 Then
            Set rsTimeDay = objO.FetchData("SELECT WAKTUMULAI, WAKTUSELESAI FROM T_STALL WITH (NOLOCK) WHERE HARI = '" & nameDay & "'")
            If Not rsTimeDay.EOF Then
                timeS = CDbl(hitungRate(Trim(rsTimeDay("WAKTUMULAI")), Trim(rsTimeDay("WAKTUSELESAI"))))
            End If
        End If
        timeP = CDbl(rsHlp("ALOKASI"))
    End If
    If CDbl((timeS - timeP)) > 0 Then
        Dim rsSublet As ADODB.Recordset
        
        'If newAbsent Then
        Set rsSublet = objO.FetchData("SELECT * FROM T_MEKANIK WITH (NOLOCK) WHERE MECH = '" & Trim(mechCode) & "' AND HARI = '" & nameDay & "' AND TANGGAL = '" & Trim(Format(nDate, "DD/MM/YYYY")) & "'")
        Dim nSQL As String
        If rsSublet.EOF Then
            ' INSERT BARU
'            If onTR = True Then
'                Call objHT.InsertData_Bengkel("INSERT INTO T_MEKANIK VALUES ('" & Trim(mechCode) & "', '" & Trim(nameDay) & "', '" & Trim(Format(nDate, "DD/MM/YYYY")) & "', '0', '','',0,0,0,'" & CDbl(timeP + IIf(divTime > (timeS - timeP), (timeS - timeP), divTime)) & "')")
'            Else
'                Call objO.InsertData("INSERT INTO T_MEKANIK VALUES ('" & Trim(mechCode) & "', '" & Trim(nameDay) & "', '" & Trim(Format(nDate, "DD/MM/YYYY")) & "', '0', '','',0,0,0,'" & CDbl(timeP + IIf(divTime > (timeS - timeP), (timeS - timeP), divTime)) & "')")
'            End If
            nSQL = "INSERT INTO T_MEKANIK VALUES ('" & Trim(mechCode) & "', '" & Trim(nameDay) & "', '" & Trim(Format(nDate, "DD/MM/YYYY")) & "', '0', '','',0,0,0,'" & CDbl(timeP + IIf(divTime > (timeS - timeP), (timeS - timeP), divTime)) & "')"
        Else
            ' UPDATE LAMA
'            If onTR = True Then
'                Call objHT.UpdateData_Bengkel("UPDATE T_MEKANIK SET ALOKASI = ALOKASI + '" & IIf(Round(divTime) > (timeS - timeP), timeS - timeP, Round(divTime)) & "' WHERE MECH = '" & Trim(mechCode) & "' AND HARI = '" & nameDay & "' AND TANGGAL = '" & Trim(Format(nDate, "DD/MM/YYYY")) & "'")
'            Else
'                Call objO.UpdateData("UPDATE T_MEKANIK SET ALOKASI = ALOKASI + '" & IIf(Round(divTime) > (timeS - timeP), timeS - timeP, Round(divTime)) & "' WHERE MECH = '" & Trim(mechCode) & "' AND HARI = '" & nameDay & "' AND TANGGAL = '" & Trim(Format(nDate, "DD/MM/YYYY")) & "'")
'            End If
            nSQL = "UPDATE T_MEKANIK SET ALOKASI = ALOKASI + '" & IIf(Round(divTime) > (timeS - timeP), timeS - timeP, Round(divTime)) & "' WHERE MECH = '" & Trim(mechCode) & "' AND HARI = '" & nameDay & "' AND TANGGAL = '" & Trim(Format(nDate, "DD/MM/YYYY")) & "'"
        End If
        UpdateMK(nMK) = nSQL
        nMK = nMK + 1
    End If
    divTime = divTime - Abs(timeS - timeP)
    If divTime > 0 Then
        countDay = countDay + 1
        LoopMk = True
        'Call countRestTime(mechCode, divTime, countDay, DateAdd("d", 1, nDate))
    End If
    LoopMk = False
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
    Exit Function
End Function

Public Function decresRestTime(mechCode As String, divTime As Double, countDay As Integer, nDate As Date)
Dim rsTimeDay As ADODB.Recordset
Dim rsHlp As ADODB.Recordset
Dim nameDay As String
Dim timeS As Double
Dim timeP As Double
Dim newAbsent As Boolean
Dim strSql As String

    On Error GoTo errHandler
    newAbsent = False
    nameDay = IIf(Weekday(nDate) = 1, 6, Weekday(nDate) - 2) & WeekdayName(IIf(Weekday(nDate) = 1, 8, Weekday(nDate)) - 1)
    
    Set rsHlp = objO.FetchData("SELECT * FROM T_MEKANIK WITH (NOLOCK) WHERE MECH = '" & Trim(mechCode) & "' AND HARI = '" & Trim(nameDay) & "' AND TANGGAL = '" & Trim(Format(nDate, "DD/MM/YYYY")) & "'")
    If Not rsHlp.EOF Then
        
        rsMK!Mech = Trim(Nul(rsHlp.Fields(0).value))
        rsMK!hari = Trim(Nul(rsHlp.Fields(1).value))
        rsMK!Tanggal = Trim(Nul(rsHlp.Fields(2).value))
        rsMK!Shift = Val(rsHlp.Fields(3).value)
        rsMK!jamMulai = Trim(Nul(rsHlp.Fields(4).value))
        rsMK!jamSelesai = Trim(Nul(rsHlp.Fields(5).value))
        rsMK!MenitKerja = Val(rsHlp.Fields(6).value)
        rsMK!Terpakai = Val(rsHlp.Fields(7).value)
        rsMK!masuk = Trim(Nul(rsHlp.Fields(8).value))
        rsMK!ALOKASI = Val(rsHlp.Fields(9).value)
        rsMK.Update

        newAbsent = False
        timeS = CDbl(rsHlp("ALOKASI"))
        timeP = CDbl(rsHlp("TERPAKAI"))
        If CDbl((timeS - timeP)) > 0 Then
            UpdateMK(nMK) = "UPDATE T_MEKANIK SET ALOKASI = ALOKASI - '" & IIf(divTime > (timeS - timeP), timeS - timeP, divTime) & "' WHERE MECH = '" & Trim(mechCode) & "' AND HARI = '" & Trim(nameDay) & "' AND TANGGAL = '" & Trim(Format(nDate, "DD/MM/YYYY")) & "' AND ALOKASI > 0"
            nMK = nMK + 1
            rsMK!Mech = Nul(rsHlp.Fields(0).value)
            rsMK!hari = Nul(rsHlp.Fields(1).value)
            rsMK!Tanggal = Nul(rsHlp.Fields(2).value)
            rsMK!Shift = Val(rsHlp.Fields(3).value)
            rsMK!jamMulai = Nul(rsHlp.Fields(4).value)
            rsMK!jamSelesai = Nul(rsHlp.Fields(5).value)
            rsMK!MenitKerja = Val(rsHlp.Fields(6).value)
            rsMK!Terpakai = Val(rsHlp.Fields(7).value)
            rsMK!masuk = Nul(rsHlp.Fields(8).value)
            rsMK!ALOKASI = Val(rsHlp.Fields(9).value) - Val(IIf(divTime > (timeS - timeP), timeS - timeP, divTime))
            rsMK.Update
        End If
        divTime = divTime - Abs(timeS - timeP)
        If divTime > 0 Then
            countDay = countDay + 1
            'Call decresRestTime(mechCode, divTime, countDay, DateAdd("d", 1, nDate))
        End If
    End If
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
Exit Function
End Function

Public Sub kedipWarranty(formO As Object)
    
    With formO
        ' JIKA ADA KETERLIBATAN WARRANTY
        If .cmdWarranty.Enabled Then
            If .cmdWarranty.BackColor = &HDCFCFB Then
                .cmdWarranty.BackColor = &HFF&
            ElseIf .cmdWarranty.BackColor = &HFF& Then
                .cmdWarranty.BackColor = &HDCFCFB
            End If
        End If
        
        ' JIKA STALL BELUM DIINPUT
        If .cmdStall.Tag = "belum" Then
            If .cmdStall.BackColor = &HDCFCFB Then
                .cmdStall.BackColor = &HFFFF&
            ElseIf .cmdStall.BackColor = &HDCFCFB Then
                .cmdStall.BackColor = &HDCFCFB
            End If
        Else
            .cmdStall.BackColor = &HDCFCFB
        End If
    End With
End Sub

Function TglNul(kode As Variant, Optional Data)
    If IsMissing(Data) Then Data = ""
    TglNul = IIf(IsNull(kode) Or kode = "", Data, kode)
End Function

Function TglNul2(kode As Variant, Optional Data)
   If IsNull(kode) Or kode = "" Then
      TglNul2 = ""
   Else
    TglNul2 = Mid(kode, 4, 2) & "/" & Left(kode, 2) & "/" & Mid(kode, 7, 4)
   End If
End Function

Function tglNotNull(Tgl As String)
    If IsNull(Tgl) Or Tgl = "" Then
      tglNotNull = "01/01/1900"
   Else
      tglNotNull = Tgl
   End If
End Function

Function cekKodePart(nSTRING As String) As String
Dim j As String

    j = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(nSTRING, ".", ""), "-", ""), "0", ""), "1", ""), "2", ""), "3", ""), "4", ""), "5", ""), "6", ""), "7", ""), "8", ""), "9", "")
    Dim nS As Integer
    nS = 0
    Dim i As Integer
    For i = 2 To Len(j)
        If Asc(UCase(Mid(j, i, 1))) >= 65 And Asc(UCase(Mid(j, i, 1))) <= 90 Then
            cekKodePart = Mid(j, i, 1)
            Exit For
        End If
    Next i
End Function

Function setKoma(nAngka As Double) As String
    setKoma = Replace(IIf(IsNull(nAngka), "0", nAngka), ",", ".")
End Function

Function setHurufAngka(Angka As String) As String
    setHurufAngka = Replace(Replace(Replace(Angka, ">", ""), " ", ""), ",", ".")
End Function

Function setHuruf(Angka As String) As String
    setHuruf = Replace(Replace(Replace(Angka, ">", ""), " ", ""), ",", ",")
End Function

Public Sub closeAllConnection()
    
    On Error Resume Next
    Set oConn = Nothing
    If oConnAcces.State <> adStateClosed Then oConnAcces.Close
    Set oConnAcces = Nothing
    If oConnAcces2.State <> adStateClosed Then oConnAcces2.Close
    Set oConnAcces2 = Nothing
    If oConnHO.State <> adStateClosed Then oConnHO.Close
    Set oConnHO = Nothing
    If oConnNis.State <> adStateClosed Then oConnNis.Close
    Set oConnNis = Nothing
    If oConnNisSP.State <> adStateClosed Then oConnNisSP.Close
    Set oConnNisSP = Nothing
    If oConnNisSP2.State <> adStateClosed Then oConnNisSP2.Close
    Set oConnNisSP2 = Nothing
    If oConnProfil.State <> adStateClosed Then oConnProfil.Close
    Set oConnProfil = Nothing
    If oConnSPK.State <> adStateClosed Then oConnSPK.Close
    Set oConnSPK = Nothing
    If oConnSPKHO.State <> adStateClosed Then oConnSPKHO.Close
    Set oConnSPKHO = Nothing
    If oConnLog.State <> adStateClosed Then oConnLog.Close
    Set oConnLog = Nothing
End Sub

Public Sub cekTampungan()
Dim rsHlp As ADODB.Recordset
Dim rsHlp2 As ADODB.Recordset
Dim rsSearch As ADODB.Recordset
Dim dBaseHO As String
Dim n

    On Error GoTo errorVB
    Set rsHlp = objO.FetchData("SELECT top 500 * FROM TAMPHISTORY WITH (NOLOCK) ORDER BY NO_CHASSIS")
    If Not rsHlp.EOF Then
        n = rsHlp.RecordCount
        dBaseHO = HODb
        If Log.LogHOConn = 1 Then
        If ConHO = True Then
            Do While Not rsHlp.EOF
                 
                Set rsSearch = objO.FetchDataHO("SELECT * FROM " & dBaseHO & "..HISTORYnew WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsHlp("NO_CHASSIS")) & "' AND No_Wo = '" & Trim(rsHlp("NO_WO")) & "' AND Noinv = '" & Trim(rsHlp("NOINV")) & "'")
                erTranHO = False
                Call objHT.OpenTransactionHO
                n = n - 1
                ' SIMPAN KE HISTORY
                If rsSearch.EOF Then
                    Call objHT.TransDataHO("INSERT INTO " & dBaseHO & "..HISTORY VALUES ('" & rsHlp("CUSTOMER") & "','" & rsHlp("CUST_ADRES") & "','" & rsHlp("KOTA") & "','" & rsHlp("NO_POLISI") & "','" & rsHlp("NO_ENGINE") & "','" & rsHlp("NO_CHASSIS") & "','" & rsHlp("FRONT") & "','" & rsHlp("MECH") & "','" & rsHlp("KM") & "','" & Format(rsHlp("DATE"), "MM/DD/YYYY") & "'," & CDbl(rsHlp("PART_G")) & "," & CDbl(rsHlp("PART_NG")) & "," & CDbl(rsHlp("OIL_G")) & ", " _
                                   & "" & CDbl(rsHlp("OIL_NG")) & "," & CDbl(rsHlp("MTRL")) & "," & CDbl(rsHlp("LABOR")) & "," & CDbl(rsHlp("BODY")) & "," & CDbl(rsHlp("SBLT")) & "," & CDbl(rsHlp("MET")) & "," & CDbl(rsHlp("TOTAL")) & ",'" & rsHlp("JOB") & "','" & rsHlp("REP1") & "','" & rsHlp("REP2") & "','" & rsHlp("REP3") & "','" & rsHlp("REP4") & "','" & rsHlp("REP5") & "','" & rsHlp("REP6") & "','" & rsHlp("REP7") & "','" & rsHlp("MEKANIK1") & "','" & rsHlp("MEKANIK2") & "','" & rsHlp("MEKANIK3") & "', " _
                                   & "'" & rsHlp("MEKANIK4") & " ','" & rsHlp("MEKANIK5") & "','" & rsHlp("MEKANIK6") & "','" & rsHlp("MEKANIK7") & "','" & rsHlp("MEKANIK8") & "','" & rsHlp("KLPKERJA1") & "','" & rsHlp("KLPKERJA2") & "','" & rsHlp("KLPKERJA3") & "','" & rsHlp("KLPKERJA4") & "','" & rsHlp("KLPKERJA5") & "','" & rsHlp("KLPKERJA6") & "','" & rsHlp("KLPKERJA7") & "', " _
                                   & "'" & rsHlp("KLPKERJA8") & " ','" & CDbl(rsHlp("RTE1")) & "','" & CDbl(rsHlp("RTE2")) & "','" & CDbl(rsHlp("RTE3")) & "','" & CDbl(rsHlp("RTE4")) & "','" & CDbl(rsHlp("RTE5")) & "','" & CDbl(rsHlp("RTE6")) & "','" & CDbl(rsHlp("RTE7")) & "','" & CDbl(rsHlp("RTE8")) & "','" & rsHlp("NO_WO") & "','" & rsHlp("NOINV") & "','" & rsHlp("NOFAK") & "','" & rsHlp("STATUS") & "','" & rsHlp("PAID") & "','" & rsHlp("DISC") & "','" & rsHlp("DISCT") & "','" & rsHlp("REMARK1") & "','" & CDbl(rsHlp("FN_RETURN")) & "', " _
                                   & "'" & rsHlp("FS_RC") & " ','" & rsHlp("FS_KD_DEALER") & "','" & rsHlp("FS_NM_DEALER") & "','" & rsHlp("TIME_WO") & "','" & rsHlp("TIME_M") & "','" & rsHlp("TIME_S") & "','" & rsHlp("TIME_A") & "','" & rsHlp("TIME_AT") & "','" & rsHlp("TIME_SA") & "','" & rsHlp("TIPEWO") & "','" & rsHlp("KODETIPE") & "','" & rsHlp("SBE") & "','" & rsHlp("FS_KD_PLG") & "', " _
                                   & "'" & rsHlp("FS_KD_WIL") & " ','" & CDbl(rsHlp("THS")) & "','" & rsHlp("MOU") & "','" & CDbl(rsHlp("PDS")) & "','" & rsHlp("memoPart") & "','" & rsHlp("memoSublet") & "','','','','','" & Format(rsHlp("INVDATE"), "MM/DD/YYYY") & "')")
                Else
                    Call objHT.TransDataHO("UPDATE " & dBaseHO & "..HISTORY SET CUSTOMER = '" & rsHlp("CUSTOMER") & "',CUST_ADRES = '" & rsHlp("CUST_ADRES") & "',KOTA = '" & rsHlp("KOTA") & "',NO_POLISI = '" & rsHlp("NO_POLISI") & "',NO_ENGINE = '" & rsHlp("NO_ENGINE") & "',FRONT = '" & rsHlp("FRONT") & "',MECH = '" & rsHlp("MECH") & "',KM = '" & rsHlp("KM") & "',DATE = '" & Format(rsHlp("DATE"), "MM/DD/YYYY") & "',PART_G = " & CDbl(rsHlp("PART_G")) & ",PART_NG = " & CDbl(rsHlp("PART_NG")) & ",OIL_G = " & CDbl(rsHlp("OIL_G")) & ", " _
                                   & "OIL_NG = " & CDbl(rsHlp("OIL_NG")) & ",MTRL = " & CDbl(rsHlp("MTRL")) & ",LABOR = " & CDbl(rsHlp("LABOR")) & ",BODY = " & CDbl(rsHlp("BODY")) & ",SBLT = " & CDbl(rsHlp("SBLT")) & ",MET = " & CDbl(rsHlp("MET")) & ",TOTAL = " & CDbl(rsHlp("TOTAL")) & ",JOB = '" & rsHlp("JOB") & "',REP1 = '" & rsHlp("REP1") & "',REP2 = '" & rsHlp("REP2") & "',REP3 = '" & rsHlp("REP3") & "',REP4 = '" & rsHlp("REP4") & "',REP5 = '" & rsHlp("REP5") & "',REP6 = '" & rsHlp("REP6") & "',REP7 = '" & rsHlp("REP7") & "', " _
                                   & "MEKANIK1 = '" & rsHlp("MEKANIK1") & "',MEKANIK2 ='" & rsHlp("MEKANIK2") & "',MEKANIK3 = '" & rsHlp("MEKANIK3") & "', MEKANIK4 = '" & rsHlp("MEKANIK4") & " ',MEKANIK5 = '" & rsHlp("MEKANIK5") & "',MEKANIK6 = '" & rsHlp("MEKANIK6") & "',MEKANIK7 = '" & rsHlp("MEKANIK7") & "',MEKANIK8 = '" & rsHlp("MEKANIK8") & "',KLPKERJA1 = '" & rsHlp("KLPKERJA1") & "',KLPKERJA2 = '" & rsHlp("KLPKERJA2") & "',KLPKERJA3 = '" & rsHlp("KLPKERJA3") & "',KLPKERJA4 = '" & rsHlp("KLPKERJA4") & "', " _
                                   & "KLPKERJA5 = '" & rsHlp("KLPKERJA5") & "',KLPKERJA6 = '" & rsHlp("KLPKERJA6") & "',KLPKERJA7 = '" & rsHlp("KLPKERJA7") & "', KLPKERJA8 = '" & rsHlp("KLPKERJA8") & " ',RTE1 = '" & CDbl(rsHlp("RTE1")) & "',RTE2 = '" & CDbl(rsHlp("RTE2")) & "',RTE3 = '" & CDbl(rsHlp("RTE3")) & "',RTE4 = '" & CDbl(rsHlp("RTE4")) & "',RTE5 = '" & CDbl(rsHlp("RTE5")) & "',RTE6 = '" & CDbl(rsHlp("RTE6")) & "',RTE7 = '" & CDbl(rsHlp("RTE7")) & "',RTE8 = '" & CDbl(rsHlp("RTE8")) & "', " _
                                   & "NOFAK = '" & rsHlp("NOFAK") & "',STATUS = '" & rsHlp("STATUS") & "',PAID = '" & rsHlp("PAID") & "',DISC = '" & rsHlp("DISC") & "',DISCT = '" & rsHlp("DISCT") & "',REMARK1 = '" & rsHlp("REMARK1") & "',FN_RETURN = '" & CDbl(rsHlp("FN_RETURN")) & "', " _
                                   & "FS_RC = '" & rsHlp("FS_RC") & " ',FS_KD_DEALER = '" & rsHlp("FS_KD_DEALER") & "',FS_NM_DEALER = '" & rsHlp("FS_NM_DEALER") & "',TIME_WO = '" & rsHlp("TIME_WO") & "',TIME_M = '" & rsHlp("TIME_M") & "',TIME_S = '" & rsHlp("TIME_S") & "', " _
                                   & "TIME_A = '" & rsHlp("TIME_A") & "',TIME_AT = '" & rsHlp("TIME_AT") & "',TIME_SA = '" & rsHlp("TIME_SA") & "',TIPEWO = '" & rsHlp("TIPEWO") & "',KODETIPE = '" & rsHlp("KODETIPE") & "',SBE = '" & rsHlp("SBE") & "',FS_KD_PLG = '" & rsHlp("FS_KD_PLG") & "', " _
                                   & "FS_KD_WIL = '" & rsHlp("FS_KD_WIL") & " ',THS = '" & CDbl(rsHlp("THS")) & "',MOU = '" & rsHlp("MOU") & "',PDS = '" & CDbl(rsHlp("PDS")) & "',memoPart = '" & rsHlp("memoPart") & "',memoSublet = '" & rsHlp("memoSublet") & "',INVDATE = '" & Format(rsHlp("INVDATE"), "MM/DD/YYYY") & "' " _
                                   & "WHERE NO_CHASSIS = '" & Trim(rsHlp("NO_CHASSIS")) & "' AND No_Wo = '" & Trim(rsHlp("NO_WO")) & "' AND Noinv = '" & Trim(rsHlp("NOINV")) & "'")
                End If
                 
                
                ' SIMPAN KE M_ESTIMASI
                 
                Set rsHlp2 = objO.FetchData("SELECT * FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & rsHlp("NO_CHASSIS") & "' AND NOWO = '" & rsHlp("NO_WO") & "'")
                If Not rsHlp2.EOF Then
                    Call objHT.TransDataHO("INSERT INTO " & dBaseHO & "..M_ESTIMASI WITH (NOLOCK) VALUES ('" & rsHlp2("NoRangka") & "','" & rsHlp2("RataKm") & "','" & Format(rsHlp2("JT_Tgl"), "MM/DD/YYYY") & "','" & Format(rsHlp2("JT_Bln"), "MM/DD/YYYY") & "','" & rsHlp2("KmNext") & "','" & rsHlp2("SBENext") & "','" & rsHlp2("KmNow") & "','" & Format(rsHlp2("TglSvcNow"), "MM/DD/YYYY") & "','" & rsHlp2("NoInv") & "','" & rsHlp2("NoWo") & "','" & rsHlp2("NamaCust") & "','" & rsHlp2("AlamatCust") & "','" & rsHlp2("NoPolisi") & "','" & rsHlp2("PSTN") & "','" & rsHlp2("GSM") & "','" & rsHlp2("CDMA") & "','','')")
                End If
                If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                
                ' SIMPAN KE CUSTDAT
                 
                Set rsHlp2 = objO.FetchData("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "'")
                If Not rsHlp2.EOF Then
                    Call objHT.TransDataHO("INSERT INTO " & dBaseHO & "..CUSTDAT VALUES ('" & rsHlp2("CUSTOMER") & "','" & rsHlp2("CUST_ADRES") & "','" & rsHlp2("KOTA") & "','" & rsHlp2("KODEPOS") & "','" & rsHlp2("TELPON") & "','" & rsHlp2("NO_POLISI") & "','" & rsHlp2("NO_ENGINE") & "','" & rsHlp2("WMI") & "','" & rsHlp2("NO_CHASSIS") & "','" & rsHlp2("MDL") & "','" & rsHlp2("MODEL") & "','" & rsHlp2("FMDL") & "','" & rsHlp2("COLOR") & "', " _
                                   & "'" & rsHlp2("FS_KD_DEALER") & " ','" & rsHlp2("FS_NM_DEALER") & "','" & rsHlp2("DDATE") & "','" & rsHlp2("REMARK") & "','" & rsHlp2("CODE") & "','" & rsHlp2("NPWP") & "'," & Replace(rsHlp2("PPN") * 100, ",", ".") & "," & rsHlp2("DISC") & ",'" & rsHlp2("KEYNO") & "','" & rsHlp2("USERADD") & "','" & rsHlp2("SPK") & "','" & Format(rsHlp2("TGL_REP"), "MM/DD/YYYY") & "','" & rsHlp2("JOB") & "','" & rsHlp2("JOBCD") & "'," & rsHlp2("FS_KM_AKHIR") & ", " _
                                   & "'" & rsHlp2("USEJOBDSC") & " ','" & Format(rsHlp2("RCDATE"), "MM/DD/YYYY") & "','" & rsHlp2("KD_K") & "','" & rsHlp2("CR") & "','" & rsHlp2("CDK") & "','" & Format(rsHlp2("TGL_D"), "MM/DD/YYYY") & "','" & rsHlp2("CDD") & "','" & rsHlp2("CRDTCD") & "','" & rsHlp2("FS_TH_MOBIL") & "','" & rsHlp2("USEFS_NM_HOBI") & "','" & rsHlp2("FS_KD_PLG") & "','" & rsHlp2("FS_KD_WIL") & "','" & Format(rsHlp2("DOTAM"), "MM/DD/YYYY") & "','" & rsHlp2("USENAME") & "', " _
                                   & "'" & rsHlp2("USEADDR") & " ','" & rsHlp2("USECITY") & "','" & rsHlp2("USEPOS") & "','" & rsHlp2("USEPSTN") & "','" & rsHlp2("USEGSM") & "','" & rsHlp2("USECDMA") & "','" & rsHlp2("USET4LAHIR") & "','" & Format(rsHlp2("USETGLLAHIR"), "MM/DD/YYYY") & "','" & rsHlp2("ASURANSI") & "','" & rsHlp2("ASR_ADDR") & "','" & rsHlp2("ASR_CITY") & "','" & Format(rsHlp2("ASR_TGLM"), "MM/DD/YYYY") & "', " _
                                   & "'" & Format(rsHlp2("ASR_TGLS"), "MM/DD/YYYY") & " '," & rsHlp2("ASR_LIMIT") & ",'" & rsHlp2("ASR_NOTE") & "','" & rsHlp2("NOTE") & "','" & rsHlp2("GSM") & "','" & rsHlp2("CDMA") & "','" & rsHlp2("USEAGAMA") & "','" & rsHlp2("USERADD") & "'," & rsHlp2("FLAGKM") & ",'" & rsHlp2("HEAVYUSE") & "','" & rsHlp2("PKP") & "','" & Trim(rsHlp2("KDPLG_ASR")) & "','" & Trim(rsHlp2("USEEMAIL")) & "')")
                End If
                If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                
                If rsHlp("MOU") <> "" And rsHlp("PAID") = "T" Then
                    ' SIMPAN KE MOU TRAN
                     
                    Set rsHlp2 = objO.FetchData("SELECT * FROM MOUTRN WITH (NOLOCK) WHERE NO_MOU = '" & rsHlp("MOU") & "'")
                    If Not rsHlp2.EOF Then
                        Call objHT.TransDataHO("INSERT INTO " & dBaseHO & "..MOUTRN VALUES ('" & Format(rsHlp2("TANGGAL"), "MM/DD/YYYY") & "', '" & rsHlp2("NO_MOU") & "', '" & rsHlp2("NO_WO") & "', '" & rsHlp2("NOINV") & "', '" & rsHlp2("NO_CHASSIS") & "', '" & rsHlp2("NO_POLISI") & "', '" & rsHlp2("FS_KD_DEALER") & "', " & CDbl(rsHlp2("TOTAL")) & ", '" & rsHlp2("STATUS") & "', '" & rsHlp2("FS_LOG_DEL") & "', '" & Format(rsHlp2("FD_TGL_DEL"), "MM/DD/YYYY") & "')")
                        Call objHT.TransDataHO("UPDATE " & dBaseHO & "..MOUHDR SET HUTANG = HUTANG + " & CDbl(rsHlp("TOTAL")) & " WHERE NO_MOU = '" & rsHlp("MOU") & "'")
                    End If
                    If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                    
                End If
                Call objO.DeleteData("DELETE TAMPHISTORY WHERE NO_WO = '" & Trim(rsHlp("NO_WO")) & "' AND NOINV = '" & Trim(rsHlp("NOINV")) & "' AND NO_CHASSIS = '" & Trim(rsHlp("NO_CHASSIS")) & "'")
                rsHlp.MoveNext
                
                If erTranHO = False Then
                    Call objHT.CloseTransactionHO
                Else
                    Call objHT.CloseTransactionHO_err
                    statusKoneksi = "                                                            Koneksi belum berhasil ulangi transfer .... "
                    Exit Sub
                End If
            Loop
            statusKoneksi = "                                                            Transfer berhasil .... "
        End If
        End If
    End If
     
    Exit Sub
    
errorVB:
    statusKoneksi = "                                                            Koneksi belum berhasil ulangi transfer .... "
    Call objHT.CloseTransactionHO_err
    Exit Sub
End Sub

Public Function jmlMekanikEM(mekanik As String) As Integer
Dim i As Integer
    
    jmlMekanikEM = 0
    For i = 1 To Len(mekanik)
        If Mid(mekanik, i, 1) = "," Then
            jmlMekanikEM = jmlMekanikEM + 1
        End If
    Next
End Function

Public Function pecahMekanik(mekanik As String, idx As Integer) As String
Dim Nama As String
Dim namaAll As String
Dim i As Integer
    
    namaAll = Trim(mekanik)
    For i = 1 To idx
        pecahMekanik = Replace(Left(namaAll, InStr(namaAll, ",") - 1), "'", "")
        namaAll = Right(namaAll, (Len(namaAll) - InStr(namaAll, ",")))
    Next
End Function

Public Function cekKodeAkun(kodeAkun As String) As String
          Dim rsAkun As ADODB.Recordset
          
1        On Error GoTo cekKodeAkun_Error

2         cekKodeAkun = ""

3             Set rsAkun = objO.FetchDataACC("SELECT KODEGL FROM GLMST WITH (NOLOCK) WHERE KODEGL = '" & Trim(kodeAkun) & "'")
4             If Not rsAkun.EOF Then
5                 cekKodeAkun = Trim(rsAkun("KODEGL"))
6             Else
7                 cekKodeAkun = ""
8             End If

9        On Error GoTo 0
10       Exit Function

cekKodeAkun_Error:
11       erTranAll = True
12       ErrMsg.judul = ""
13       ErrMsg.message = "Error " & err.Number & " (" & err.Description & ") in procedure cekKodeAkun of Module MainModul"
14       ErrMsg.show
         
      '   MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cekKodeAkun of Module MainModul"
End Function

Public Function cekValidasi(noOrder As String) As Boolean
Dim rsValid As ADODB.Recordset
    
    cekValidasi = False
    Set rsValid = objO.FetchData("SELECT NO_WO FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(noOrder) & "' AND USERDEL = ''")
    If Not rsValid.EOF Then
        cekValidasi = True
    End If
End Function

Public Function hargaDealerSBI(nRate As Double, tipeClass As String, noRangka As String) As Double
Dim rsHlp2 As ADODB.Recordset
Dim nStatusDealer As Integer
Dim nHarga As Double
    
    hargaDealerSBI = 0
    
    myChassis = Trim(noRangka)
    Set rsHlp2 = objBilling.takeInvoiceName(Trim(noRangka))
    If Not rsHlp2.EOF Then
        nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
        If nStatusDealer = 1 Then
            nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
        Else
            If nStatusDealer = 2 Then
                nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
            ElseIf nStatusDealer = 3 Then
                nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
            End If
        End If
        hargaDealerSBI = Round(CDbl(nRate) * CDbl(nHarga))
    End If
End Function

Public Function hargaCustomerCBUNonTAM(nRate As Double, tipeClass As String) As Double
Dim rsHlp2 As ADODB.Recordset
Dim nStatusDealer As Integer
Dim nHarga As Double
    
    hargaCustomerCBUNonTAM = 0
    nHarga = objBilling.takeClassPrice(tipeClass, 1)
    hargaCustomerCBUNonTAM = nRate * nHarga
    
End Function

Public Sub saveLogInUser(UserName As String, kompieName As String)
    Dim cm As New Command
    If oConnLog.State = adStateClosed Then oConnLog.Open (SQLConProfilBK)
    cm.CommandText = ("INSERT INTO LogUser (kodeUser,clientName,versiNIS,startLogin) VALUES ('" & Trim(UserName) & "','" & Trim(kompieName) & "','" & Trim(gVersi) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')")
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = oConnLog
    cm.Execute
    Set cm = Nothing
    If oConnLog.State <> adStateClosed Then oConnLog.Close
End Sub

Public Sub saveLogOutUser(UserName As String, kompieName As String)
If allowSaveLog Then
    Dim rsLog As ADODB.Recordset
    Set rsLog = FetchLogUser("SELECT id FROM LogUser WITH (NOLOCK) WHERE KODEUSER = '" & Trim(UserName) & "' AND CLIENTNAME = '" & Trim(kompieName) & "' ORDER BY ID DESC")
    If Not rsLog.EOF Then
        Dim cm As New Command
        If oConnLog.State = adStateClosed Then oConnLog.Open (SQLConProfilBK)
        cm.CommandText = ("UPDATE LogUser SET endLogin = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "' WHERE KODEUSER = '" & Trim(UserName) & "' AND CLIENTNAME = '" & Trim(kompieName) & "' AND ID = '" & Trim(rsLog("ID")) & "'")
        cm.CommandType = adCmdText
        Set cm.ActiveConnection = oConnLog
        cm.Execute
        Set cm = Nothing
        If oConnLog.State <> adStateClosed Then oConnLog.Close
    End If
End If
End Sub

Private Function FetchLogUser(strSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    If oConnLog.State = adStateClosed Then oConnLog.Open (SQLConProfilBK)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConProfilBK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchLogUser = rs
    If oConnLog.State = adStateOpen Then oConnLog.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub saveLogMenu(UserName As String, kompieName As String, versi As String, subMenu As String, nameFunction As String, inState As Integer)
1     On Error GoTo vb_error
      Dim rsLog As ADODB.Recordset
2         If allowSaveLog Then
3                 Select Case inState
                      Case 0
                          ' begintrans log
4                         Call objO.InsertData_Profil("INSERT INTO LogMenu VALUES ('" & Trim(UserName) & "','" & Trim(kompieName) & "','" & Trim(versi) & "','" & Trim(subMenu) & "', '" & Trim(nameFunction) & "', GETDATE(),'','')")
5                     Case 1
                          ' committrans log
6                         Set rsLog = objO.FetchData_Profil("SELECT id FROM LogMenu WITH (NOLOCK) WHERE KODEUSER = '" & Trim(UserName) & "' AND CLIENTNAME = '" & Trim(kompieName) & "' AND ISNULL(ENDCOMAND,'') = '' ORDER BY ID DESC")
7                         If Not rsLog.EOF Then
8                             Call objO.UpdateData_Profil("UPDATE LogMenu SET EndComand = GETDATE() WHERE KODEUSER = '" & Trim(UserName) & "' AND CLIENTNAME = '" & Trim(kompieName) & "' AND ID = '" & Trim(rsLog("ID")) & "'")
9                         End If
10                    Case 2
                          ' rollback log
11                        Set rsLog = objO.FetchData_Profil("SELECT id FROM LogMenu WITH (NOLOCK) WHERE KODEUSER = '" & Trim(UserName) & "' AND CLIENTNAME = '" & Trim(kompieName) & "' AND ISNULL(ROLLCOMAND,'') = '' ORDER BY ID DESC")
12                        If Not rsLog.EOF Then
13                            Call objO.UpdateData_Profil("UPDATE LogMenu SET rollComand = GETDATE() WHERE KODEUSER = '" & Trim(UserName) & "' AND CLIENTNAME = '" & Trim(kompieName) & "' AND ID = '" & Trim(rsLog("ID")) & "'")
14                        End If
15                End Select
16        End If
17        Exit Sub
vb_error:
18    MsgBox ErrMessage(Erl, err.Number, "Procedure : Save Log Menu"), vbInformation
End Sub

Public Function pembulatanJam(timeInput As String) As String

    Dim jInput As Integer
    Dim mInput As Integer
    Dim menitIN As Integer
    
    jInput = Left(timeInput, InStr(timeInput, ":") - 1)
    mInput = Right(timeInput, InStr(timeInput, ":") - 1)
    menitIN = (jInput * 60) + mInput
    
    If (menitIN Mod 10) <> 0 Then
        mInput = mInput + (10 - (menitIN Mod 10))
        If mInput >= 60 Then
            jInput = jInput + Round(mInput / 60)
            mInput = mInput Mod 60
            If jInput >= 24 Then
                jInput = jInput Mod 24
            End If
        End If
    End If
    pembulatanJam = Format(jInput, "00") & ":" & Format(mInput, "00")
End Function

Public Function cekLastIsSBIDealer(noRangka As String) As Boolean
Dim rsHlp As ADODB.Recordset
    
    cekLastIsSBIDealer = False
    Set rsHlp = objO.FetchData("SELECT FS_RC FROM T_INVOICE WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "' AND RPR_COND <> 'K' AND DELIV <> 'BTL' AND TPSVC <> 'W' AND FS_LOG_DEL = '' ORDER BY DATE DESC, NO_WO DESC")
    If Not rsHlp.EOF Then
        'If rsHlp.RecordCount <= 2 Then
            If Trim(rsHlp("FS_RC")) = "R001" Then
                cekLastIsSBIDealer = True
            End If
        'End If
    End If
     
End Function

Public Function cekWoInvoice(noRangka As String) As Boolean
Dim rsHlp As ADODB.Recordset
    
    cekWoInvoice = False
    Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "' AND RPR_COND <> 'K' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' And FS_RC <> 'R001' And PDS <> 1 AND TPSVC <> 'W' AND NOINV = '' ORDER BY DATE DESC")
    If Not rsHlp.EOF Then
        cekWoInvoice = True
    End If
     
End Function

Public Function adaBelumClosing(noRangka As String) As Boolean
Dim rsHlp As ADODB.Recordset

    adaBelumClosing = False
    Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "' AND RPR_COND <> 'K' AND DELIV <> 'BTL'  AND FS_LOG_DEL = '' AND TPSVC <> 'W'")
    If Not rsHlp.EOF Then
        If rsHlp.RecordCount > 0 Then
            adaBelumClosing = True
        End If
    End If
End Function

Public Function cekSBIDealerInvoice(noRangka As String) As Boolean
Dim rsHlp As ADODB.Recordset
    
    cekSBIDealerInvoice = False
    Set rsHlp = objO.FetchData("SELECT NOINV FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "' AND RPR_COND <> 'K' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND FS_RC = 'R001' ORDER BY DATE DESC")
    If Not rsHlp.EOF Then
        If Trim(rsHlp("NOINV")) <> "" Then
            cekSBIDealerInvoice = True
        End If
    Else
        cekSBIDealerInvoice = True
    End If
     
End Function

Public Function cekLastIsPDS(noChs As String) As Boolean
Dim rsSearch As ADODB.Recordset

    cekLastIsPDS = False
     
    Set rsSearch = objO.FetchData("SELECT PDS FROM HISTORY WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noChs) & "' ORDER BY DATE DESC, NO_WO DESC")
    If Not rsSearch.EOF Then
        'If rsSearch.RecordCount = 1 Then
            If Trim(rsSearch("PDS")) = "1" Then
                cekLastIsPDS = True
            End If
        'End If
    End If
     
End Function

Public Function newUnit(noChs As String) As Boolean
Dim rsSearch As ADODB.Recordset

    newUnit = True
     
    Set rsSearch = objO.FetchData("SELECT NO_CHASSIS FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noChs) & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = ''")
    If Not rsSearch.EOF Then
        If rsSearch.RecordCount >= 1 Then
            newUnit = False
        End If
    End If
     
End Function

Public Function adaOnlySublet(noRangka As String, noWO As String) As Boolean
Dim rsHlp As ADODB.Recordset

    adaOnlySublet = False
    Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "' AND NO_WO = '" & Trim(noWO) & "'")
    If Not rsHlp.EOF Then
        adaOnlySublet = True
    End If
End Function

Public Function cekRetur(NoInvoice As String, statusRetur As String) As String
    
    cekRetur = NoInvoice
    If Trim(statusRetur) = "BTL" Then
        cekRetur = NoInvoice & "R"
    End If
End Function

Public Function updateAntrian(nomorPol As String, model As String, Nama As String, tglJanji As Date, jamJanji As String, tipeUpdate As Integer, noBook As String)
      ' ==:: 25 AGUSTUS 2010 ; SIMPAN DATA UNTUK ANTRIAN ::==
1         On Error GoTo errorVB
         ' If Get_ConnectionQueue("Queue" & Right(gDb, 2), gUser, gPwd, gServerDb) Then
2             Select Case tipeUpdate
                  Case 1
                      '$$$$$ SIMPAN $$$$$
3                     Call support.insertBookQueue(Trim(nomorPol), Left(Trim(model), 25), Left(Trim(Nama), 22) & IIf(Len(Trim(Nama)) > 22, "...", ""), Format(tglJanji, "MM/DD/YYYY"), Format(jamJanji, "HH:NN"), 0, Trim(noBook)) 'oConnQueue.Execute ("INSERT INTO DATABOOKING (NOPOL, TIPE, NAMA, TGLBOK, JAMBOK, AKTIV, NOBOK ) VALUES ('" & Trim(nomorPol) & "','" & Left(Trim(model), 25) & "', '" & Left(Trim(Nama), 22) & IIf(Len(Trim(Nama)) > 22, "...", "") & "', '" & Format(tglJanji, "MM/DD/YYYY") & "','" & Format(jamJanji, "HH:NN") & "','0','" & Trim(noBook) & "')")
4                 Case 2
                      '$$$$$ UPDATE $$$$$
5                     Call support.updateBookQueue(Trim(nomorPol), Left(Trim(model), 25), Left(Trim(Nama), 22) & "...", Format(tglJanji, "MM/DD/YYYY"), Format(jamJanji, "HH:NN"), 0, Trim(noBook))   'oConnQueue.Execute ("UPDATE DATABOOKING SET TIPE = '" & Left(Trim(model), 25) & "', NAMA = '" & Left(Trim(Nama), 22) & "..." & "', TGLBOK = '" & Format(tglJanji, "MM/DD/YYYY") & "', JAMBOK = '" & Format(jamJanji, "HH:NN") & "', AKTIV = 0 WHERE NOPOL = '" & Trim(nomorPol) & "' AND NOBOK = '" & Trim(noBook) & "'")
6                 Case 3
                      '$$$$$ DELETE $$$$$
7                     Call support.deleteBookQueue(Trim(nomorPol), Trim(noBook))   'oConnQueue.Execute ("DELETE DATABOOKING WHERE NOPOL = '" & Trim(nomorPol) & "' AND NOBOK = '" & Trim(noBook) & "'")
8             End Select
         ' End If
         ' If oConnQueue.State <> adStateClosed Then oConnQueue.Close
9         Exit Function
errorVB:
10       erTranAll = True
11       ErrMsg.judul = ""
12       ErrMsg.message = "Log Antrian Error : Update Antrian" & Chr(13) & " Code : " & tipeUpdate
13       ErrMsg.show
         
      ''MsgBox ("Log Antrian Error : Update Antrian" & Chr(13) & " Code : " & tipeUpdate)
      'If oConnQueue.State <> adStateClosed Then oConnQueue.Close
End Function

Public Function finishService(noWO As String, nomorPol As String, model As String, callModel As String, Nama As String, tglFinish As String, jamFinish As String, tipeUpdate As Integer)
' ==:: 25 AGUSTUS 2010 ; SIMPAN DATA UNTUK ANTRIAN ::==
    On Error GoTo errorVB
    'If Get_ConnectionQueue("Queue" & Right(gDb, 2), gUser, gPwd, gServerDb) Then
        Select Case tipeUpdate
            Case 1
                '$$$$$ SIMPAN $$$$$
                Call support.setQueue("INSERT INTO " & Log.LogQueue & "..FINISHSERVICE (NoWo, NOPOL, NAMA, MODEL, TIPE, TGLFINISH, JAMFINISH, PANGGIL, TGLPANGGIL) VALUES ('" & Trim(noWO) & "',  '" & Trim(nomorPol) & "', '" & Left(Trim(Nama), 25) & "','" & Left(Trim(model), 25) & "','" & Left(Trim(callModel), 25) & "', '" & Trim(tglFinish) & "','" & Format(jamFinish, "HH:NN") & "','0','')")
                'Call support.setQueue("INSERT INTO " & Log.LogQueue & "..FINISHSERVICE (NOPOL, NAMA, MODEL, TIPE, TGLFINISH, JAMFINISH, PANGGIL, TGLPANGGIL) VALUES ('" & Trim(nomorPol) & "', '" & Left(Trim(Nama), 25) & "','" & Left(Trim(model), 25) & "','" & Left(Trim(callModel), 25) & "', '" & Trim(tglFinish) & "','" & Format(jamFinish, "HH:NN") & "','0','')")
                'oConnQueue.Execute ("INSERT INTO " & Log.LogQueue & "..FINISHSERVICE (NOPOL, NAMA, MODEL, TIPE, TGLFINISH, JAMFINISH, PANGGIL, TGLPANGGIL) VALUES ('" & Trim(nomorPol) & "', '" & Left(Trim(Nama), 25) & "','" & Left(Trim(model), 25) & "','" & Left(Trim(callModel), 25) & "', '" & Trim(tglFinish) & "','" & Format(jamFinish, "HH:NN") & "','0','')")
            Case 2
                '$$$$$ DELETE $$$$$
                Call support.setQueue("DELETE " & Log.LogQueue & "..FINISHSERVICE WHERE nowo = '" & Trim(noWO) & "' and NOPOL = '" & Trim(nomorPol) & "' AND TGLFINISH = '" & Trim(tglFinish) & "'")
                'Call support.setQueue("DELETE " & Log.LogQueue & "..FINISHSERVICE WHERE NOPOL = '" & Trim(nomorPol) & "' AND TGLFINISH = '" & Trim(tglFinish) & "'")
                'oConnQueue.Execute ("DELETE " & Log.LogQueue & "..FINISHSERVICE WHERE NOPOL = '" & Trim(nomorPol) & "' AND TGLFINISH = '" & Trim(tglFinish) & "'")
        End Select
    'End If
    'If oConnQueue.State <> adStateClosed Then oConnQueue.Close
    Exit Function
errorVB:
MsgBox ("Log Antrian Error : Finish Service" & Chr(13) & " Code : " & tipeUpdate)
'If oConnQueue.State <> adStateClosed Then oConnQueue.Close
End Function

Public Function namaModel(model As String) As String

    model = LCase(model)
    If InStr(model, "avanza") <> 0 Then
        namaModel = "Avanza"
    Else
        If InStr(model, "camry") <> 0 Then
            namaModel = "Camry"
        Else
            If InStr(model, "corolla") <> 0 Then
                namaModel = "Corolla"
            Else
                If InStr(model, "crown") <> 0 Then
                    namaModel = "Crown"
                Else
                    If InStr(model, "dyna") <> 0 Then
                        namaModel = "Dyna"
                    Else
                        If InStr(model, "fortuner") <> 0 Then
                            namaModel = "Fortuner"
                        Else
                            If InStr(model, "hilux") <> 0 Then
                                namaModel = "Hilux"
                            Else
                                If InStr(model, "kijang") <> 0 Then
                                    namaModel = "Kijang"
                                Else
                                    If InStr(model, "innova") <> 0 Then
                                        namaModel = "Innova"
                                    Else
                                        If InStr(model, "landcruiser") <> 0 Then
                                            namaModel = "Lancruiser"
                                        Else
                                            If InStr(model, "previa") <> 0 Then
                                                namaModel = "Previa"
                                            Else
                                                If InStr(model, "rush") <> 0 Then
                                                    namaModel = "Rush"
                                                Else
                                                    If InStr(model, "soluna") <> 0 Then
                                                        namaModel = "Soluna"
                                                    Else
                                                        If InStr(model, "starlet") <> 0 Then
                                                            namaModel = "Starlet"
                                                        Else
                                                            If InStr(model, "vios") <> 0 Then
                                                                namaModel = "Vios"
                                                            Else
                                                                If InStr(model, "yaris") <> 0 Then
                                                                    namaModel = "Yaris"
                                                                Else
                                                                    If InStr(model, "alphard") <> 0 Then
                                                                        namaModel = "buildup"
                                                                    Else
                                                                        If InStr(model, "altis") <> 0 Then
                                                                            namaModel = "Corolla"
                                                                        Else
                                                                            If InStr(model, "harrier") <> 0 Then
                                                                                namaModel = "buildup"
                                                                            Else
                                                                                If InStr(model, "lexus") <> 0 Then
                                                                                    namaModel = "buildup"
                                                                                Else
                                                                                    If InStr(model, "prius") <> 0 Then
                                                                                        namaModel = "buildup"
                                                                                    Else
                                                                                        namaModel = "Lain-Lain"
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function cekEXE()
    Dim wmi, QRY, LIST, Item
    Dim CekApp As Boolean
    Dim CekApp2 As Boolean
      'QRY = "SELECT * FROM Win32_Process WHERE Name like 'NIS Service%'"
      QRY = "SELECT * FROM Win32_Process"
      
      Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}\\.\root\cimv2")
      Set LIST = wmi.ExecQuery(QRY)
      Dim i As Integer
      i = 1
      For Each Item In LIST
         Dim f As New FileSystemObject
         Dim hasil As String
         If NulSID(Item.ExecutablePath) <> "" Then
            hasil = GetOriginalName(Item.ExecutablePath)
            
            If NulSID(hasil) = NulSID(oriFileNameNISService) And GetCurrentProcessId <> Item.ProcessID Then ' And App.Path & "\" <> Mid(Trim(Item.ExecutablePath), 1, Len(Trim(Item.ExecutablePath)) - Len(Trim(Item.Name))) Then
               If MsgBox("Aplikasi NIS Service sudah ada yang terbuka, Apakah ingin menutup NIS Service sebelumnya?", vbYesNo) = vbYes Then
                    'Shell "taskkill.exe /f /t /im " & Item.Name
                    Item.Terminate
               Else
                    End
               End If
            ElseIf NulSID(hasil) = NulSID(oriFileNameNISService) And GetCurrentProcessId = Item.ProcessID Then
                i = i + 1
            ElseIf NulSID(hasil) = NulSID(oriFileNameNISService) And GetCurrentProcessId <> Item.ProcessID Then
                Item.Terminate
            End If
         End If
      Next
End Function

Public Function searchPrinter(namePrinter) As Integer
    
    Dim rsPrint As ADODB.Recordset
    Set rsPrint = objO.FetchData("SELECT IDPRINT FROM TIPEPRINTER WITH (NOLOCK) WHERE TPPRINT = '" & Trim(namePrinter) & "'")
    If Not rsPrint.EOF Then
        searchPrinter = Trim(rsPrint("IDPRINT"))
    End If
End Function

Public Function onlyNum(karakter As String) As String
Dim nKar As Integer
Dim i As Integer
    
    onlyNum = ""
    nKar = Len(karakter)
    For i = 1 To nKar
        If Asc(Mid(karakter, i, 1)) >= 48 And Asc(Mid(karakter, i, 1)) <= 57 Then
            onlyNum = onlyNum + Mid(karakter, i, 1)
        End If
    Next
End Function

Public Function ComputerIsOnline(ByVal strComputerName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strResult   'As String
    Dim nDeL        As String
    Dim ShellX      As String
    Dim FileNum     As Integer
    Dim lPid        As Long
    Dim nCounter    As Integer
    
    strResult = ""
    nCounter = 0
    nDeL = Left(PathWindows, Len(PathWindows) - 1) & "\" & Left(McDelIndex(7), 8) & ".txt"
    ShellX = Shell("command.com /c ping -n 2 " & strComputerName & " > " & nDeL, vbHide)
    lPid = ShellX
    If lPid <> 0 Then
        FileNum = FreeFile
        'Open nDeL For Input As #FileNum
        'strResult = Input(LOF(1), 1)
REPEAT:
        strResult = getShell(nDeL)
        
        If Trim(strResult) = "" And nCounter < 3 Then
            nCounter = nCounter + 1
            GoTo REPEAT
        End If
        'Close #FileNum
        'ComputerIsOnline = (InStr(strResult, "Lost = 0") > 0)
        ComputerIsOnline = (InStr(strResult, "Reply") > 0)
        Call DeleteFile(nDeL)
    End If
    Exit Function
ErrorHandler:
    ComputerIsOnline = False
    Exit Function
End Function

Private Sub DeleteFile(nPIC As String)
On Error GoTo ErHandler
   Dim FSO, fileobjAccect
   Set FSO = CreateObject("Scripting.FileSystemobject")
   Set fileobjAccect = FSO.GetFile(nPIC)
   fileobjAccect.Delete
   Set fileobjAccect = Nothing
   Set FSO = Nothing
   Exit Sub
ErHandler:
   Set fileobjAccect = Nothing
   Set FSO = Nothing
End Sub

Private Function GetSQLServerName(ByVal strSQLServer As String, strUser As String, strPass As String) As Boolean
On Error GoTo ErrServer
    Dim oSQLServer As SQLDMO.SQLServer
    Dim oDBase     As SQLDMO.database
    
    Set oSQLServer = New SQLDMO.SQLServer
    oSQLServer.LoginSecure = False
    oSQLServer.BlockingTimeout = 1
    oSQLServer.Connect strSQLServer, strUser, strPass
    GetSQLServerName = True
    
Exit Function
ErrServer:
    GetSQLServerName = False
End Function

Public Function cekIncludePPn(nominal As Double) As Double
Dim rsSearch As ADODB.Recordset
    
     
    Set rsSearch = objO.FetchData("SELECT INCLUDEPPN FROM T_PROFIL WITH (NOLOCK)")
    If Not rsSearch.EOF Then
        Select Case Trim(rsSearch("INCLUDEPPN"))
            Case 0
                ' HARGA NON INCLUDE PPN
                'cekIncludePPn = CDbl(nominal) + Round((CDbl(nominal) * (10 / 100)), 2)
                cekIncludePPn = CDbl(nominal) + ((CDbl(nominal) * (profilACC.nilaippn / profilACC.nilaidpp))) ''(10 / 100)))
            Case 1
                ' HARGA INCLUDE PPN
                cekIncludePPn = CDbl(nominal)
        End Select
    End If
     
End Function

Public Function cekPPn(nominal As Double) As Double
Dim rsSearch As ADODB.Recordset
    
     
    Set rsSearch = objO.FetchData("SELECT INCLUDEPPN FROM T_PROFIL WITH (NOLOCK)")
    If Not rsSearch.EOF Then
        Select Case Trim(rsSearch("INCLUDEPPN"))
            Case 0
                ' HARGA NON INCLUDE PPN
                cekPPn = CDbl(nominal) * (10 / 100)
            Case 1
                ' HARGA INCLUDE PPN
                cekPPn = CDbl(nominal) * (10 / 110)
        End Select
    End If
     
End Function

Public Function trimSS(kata As String) As String
    trimSS = IIf(IsNull(kata), "", Trim(kata))
End Function

Public Sub simpanAntrian(id As Integer, aType As String, aNumber As Integer, aDate As String, aWO As String, aNoPol As String, aDateSelesai As Date)
1     On Error GoTo vb_error

          Dim rsTmp As ADODB.Recordset
          Dim rsTemp As ADODB.Recordset
          Dim rsChassing As ADODB.Recordset
          Dim tanggalIn As String
          Dim tanggalPanggil As String
          Dim NoFak As String
          Dim tanggalWo As String
          Dim NoPolisi, noRangka, firstName, lastName, email, phoneNumber As String
          Dim km As Double
          Dim noBookingId, NoBooking, dealer_customer_ID As String
          Dim tglBooking As Date
          Dim oneAccount As String
          Dim eventID As String
          
          
2         Select Case id
              Case 1
                  '' INSERT
3                 Set rsTmp = objO.FetchData("SELECT * FROM " & Log.LogQueue & "..QUEUE WITH (NOLOCK) WHERE TIPE = '" & Trim(aType) & "' AND NOMOR = '" & Trim(aNumber) & "' AND CONVERT( VARCHAR(10), TGL_PANGGIL,101) = '" & Format(TglYYYYMMDD(aDate), "MM/DD/YYYY") & "'")
4                 If Not rsTmp.EOF Then
5                     tanggalIn = Format(rsTmp("Tgl_In"), "YYYY-MM-DD HH:MM:SS")
6                     Set rsTemp = objO.FetchData("SELECT * FROM " & Log.LogQueue & "..PANGGILAN WITH (NOLOCK) WHERE TIPE = '" & Trim(aType) & "' AND QUEUE = '" & Trim(aNumber) & "' AND CONVERT( VARCHAR(10), TGL_PANGGIL,101) = '" & Format(TglYYYYMMDD(aDate), "MM/DD/YYYY") & "'")
7                     If Not rsTemp.EOF Then
8                         tanggalPanggil = Format(rsTemp("Tgl_panggil"), "YYYY-MM-DD HH:MM:SS")
9                         Call objHT.InsertData_Bengkel("INSERT INTO LISTANTRIAN VALUES ('" & Trim(rsTmp("ID")) & "','" & Format(TglYYYYMMDD(aDate), "MM/DD/YYYY") & "','" & Trim(aType) & "','" & Trim(aNumber) & "','" & Trim(rsTemp("ID")) & "','" & Trim(aWO) & "', '" & Trim(gSA) & "', 'OKE')")
10                        Call objHT.UpdateData_Bengkel("UPDATE " & Log.LogQueue & "..PANGGILAN SET NOPOL = '" & Trim(aNoPol) & "' WHERE TIPE = '" & Trim(aType) & "' AND QUEUE = '" & Trim(aNumber) & "' AND CONVERT( VARCHAR(10), TGL_PANGGIL,101) = '" & Format(TglYYYYMMDD(aDate), "MM/DD/YYYY") & "'")
11                    End If
12                End If
13            Case 2
                  '' BATAL WO
14                Call objHT.UpdateData_Bengkel("UPDATE LISTANTRIAN SET STATUS = 'BTL' WHERE WO = '" & Trim(aWO) & "'")
15            Case 3
                  '' GANTI NO ANTRIAN
16                Set rsTemp = objO.FetchData("SELECT * FROM " & Log.LogQueue & "..PANGGILAN WITH (NOLOCK) WHERE TIPE = '" & Trim(aType) & "' AND QUEUE = '" & Trim(aNumber) & "' AND CONVERT( VARCHAR(10), TGL_PANGGIL,101) = '" & Format(TglYYYYMMDD(aDate), "MM/DD/YYYY") & "'")
                  'revisi 25 juli 2012
                  'Set rsTmp = objO.FetchData("SELECT * FROM " & Log.LogQueue & "..QUEUE a INNER JOIN " & Log.LogQueue & "..PANGGILAN B ON A.TIPE = B.TIPE AND A.NOMOR = B.QUEUE AND CONVERT( VARCHAR(10), A.TGL_PANGGIL,101) = CONVERT( VARCHAR(10), B.TGL_PANGGIL,101) WHERE A.TIPE = '" & Trim(aType) & "' AND NOMOR = '" & Trim(aNumber) & "' AND CONVERT( VARCHAR(10), A.TGL_PANGGIL,101) = '" & Format(TglYYYYMMDD(aDate), "MM/DD/YYYY") & "'")
17                Set rsTmp = objO.FetchData("SELECT A.* FROM " & Log.LogQueue & "..QUEUE A WITH (NOLOCK) INNER JOIN " & Log.LogQueue & "..PANGGILAN B WITH (NOLOCK) ON A.TIPE = B.TIPE AND A.NOMOR = B.QUEUE AND CONVERT( VARCHAR(10), A.TGL_PANGGIL,101) = CONVERT( VARCHAR(10), B.TGL_PANGGIL,101) WHERE A.TIPE = '" & Trim(aType) & "' AND NOMOR = '" & Trim(aNumber) & "' AND CONVERT( VARCHAR(10), A.TGL_PANGGIL,101) = '" & Format(TglYYYYMMDD(aDate), "MM/DD/YYYY") & "'")
                  'di tutup 3/8/2012  Call objHT.UpdateData_Bengkel("UPDATE " & Log.LogQueue & "..PANGGILAN SET NOPOL = '' WHERE NOPOL = '" & Trim(aNoPol) & "'")
18                If Not rsTmp.EOF Then
19                    tanggalIn = Format(rsTmp("Tgl_In"), "YYYY-MM-DD HH:MM:SS")
20                    If Not rsTemp.EOF Then
21                        tanggalPanggil = Format(rsTemp("Tgl_panggil"), "YYYY-MM-DD HH:MM:SS")
22                        Call objHT.UpdateData_Bengkel("UPDATE LISTANTRIAN SET IDPRINT = '" & Trim(rsTmp("ID")) & "', TANGGAL = '" & Format(TglYYYYMMDD(aDate), "MM/DD/YYYY") & "', TIPE = '" & Trim(aType) & "', QUEUE = '" & Trim(aNumber) & "', IDCALL = '" & Trim(rsTemp("ID")) & "' WHERE WO = '" & Trim(aWO) & "'")
23                        Call objHT.UpdateData_Bengkel("UPDATE " & Log.LogQueue & "..PANGGILAN SET NOPOL = '" & Trim(aNoPol) & "' WHERE TIPE = '" & Trim(aType) & "' AND QUEUE = '" & Trim(aNumber) & "' AND CONVERT( VARCHAR(10), TGL_PANGGIL,101) = '" & Format(TglYYYYMMDD(aDate), "MM/DD/YYYY") & "'")
24                    End If
25                End If
26        End Select
27        If Log.DiDxP1 = 1 Then
28            Set rsChassing = objO.FetchData("SELECT no_polisi,nofak,[date] + [time] as tanggalWO,no_chassis,km FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(aWO) & "'")
29            If Not rsChassing.EOF Then
30                NoPolisi = Trim(rsChassing("No_polisi"))
31                NoFak = Trim(rsChassing("Nofak"))
32                tanggalWo = Format(Trim(rsChassing("tanggalwo")), "YYYY-MM-DD HH:MM:DD")
33                noRangka = NulSID(rsChassing!NO_CHASSIS)
34                km = NulSID(rsChassing!km)
35            End If
36            tglBooking = "1900-01-01"
              Dim selisih As String
37            selisih = DateDiff("n", Format(tanggalWo, "YYYY-MM-DD HH:MM:SS"), Format(aDateSelesai, "YYYY-MM-DD HH:MM:SS"))
38            If NoFak <> "" Then
39                noBookingId = objO.FetchData_String("SELECT KODETIPE FROM book WITH (NOLOCK) WHERE nofak = '" & NoFak & "' ")
40                tglBooking = objO.FetchData_String("SELECT convert(VARCHAR(20),DATE + TIME,120) FROM book WITH (NOLOCK) WHERE nofak = '" & NoFak & "'")
41                If NulSID(noBookingId) <> "" Then
                       '' isi nobookingapp dan kodeprogram
42                    If InStr(1, NulSID(noBookingId), "#|#") > 0 Then
43                        noBookingId = Left(NulSID(noBookingId), InStr(1, NulSID(noBookingId), "#|#") - 1)
44                    End If
45                End If
46                If NulSID(noBookingId) = "" Then
47                    noBookingId = Right(gDb, 2) & "." & aWO
48                End If
49                NoBooking = Right(gDb, 2) & "." & NoFak
50            Else
51                NoBooking = Right(gDb, 2) & "." & aWO
52                noBookingId = Right(gDb, 2) & "." & aWO
53            End If
              
54            If noRangka <> "" Then
55                Set rsChassing = objO.FetchData("SELECT * FROM CUSTDAT with (nolock) WHERE no_chassis = '" & Trim(noRangka) & "'")
56                If Not rsChassing.EOF Then
                      Dim outPut() As String
57                    outPut = Split(rsChassing!customer, " ", 2)
58                    firstName = Left(NulSID(outPut(0)), 64)
59                    lastName = Left(NulSID(outPut(1)), 64)
60                    email = NulSID(rsChassing!useEMAIL)
61                    phoneNumber = NulSID(rsChassing!GSM)
62                    dealer_customer_ID = NulSID(rsChassing!FS_KD_PLG)
63                End If
64            End If
              
65            Set rsChassing = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MonitorSvcProggres WHERE booking_number = '" & Trim(NoBooking) & "'")
              
66            If rsChassing.EOF Then
67                eventID = objO.FetchData_String("select event_id from [" & dbReplikasiHO & "]..bookingonline_didx_toSend where nobookingapps = '" & noBookingId & "'")
68                oneAccount = objO.FetchData_String("SELECT one_account_ID FROM [" & dbReplikasiHO & "]..bookingonline_didx_toSend WHERE phone_number = '" & phoneNumber & "'")
69                sSQL = "INSERT INTO [" & dbReplikasiHO & "]..MonitorSvcProggres (one_account_ID,dealer_customer_ID,first_name,last_name,phone_number,email,VIN,police_number,actual_mileage,booking_ID,booking_number,created_datetime,service_location,service_category,job_type,slot_datetime_start,outlet_ID,outlet_name,wo_number,duration_estimation,reception_datetime,wo_created_datetime,clock_on_datetime) "
70                sSQL = sSQL & " VALUES ('" & oneAccount & "','" & dealer_customer_ID & "','" & firstName & "','" & lastName & "','" & phoneNumber & "','" & email & "','" & noRangka & "','" & NoPolisi & "','" & km & "','" & noBookingId & "','" & NoBooking & "','" & Format(tglBooking, "YYYY-MM-DD HH:mm:ss") & "','WORKSHOP','PERIODIC_MAINTENANCE','OTHERS','" & tanggalWo & "','" & Log.kdDealerDidx & "','" & NulSID(Log.LogNamaNas) & "'," & Right(gDb, 2) & "." & Trim(aWO) & "," & selisih & ",'" & tanggalWo & "','" & tanggalWo & "','1900-01-01')"
71                Call objHT.InsertData_Bengkel(sSQL)
                  
72                sSQL = "INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,status,proses,event_no,tglentry) VALUES ('MONITORINGGR','" & noBookingId & "','" & NoBooking & "','','" & Right(gDb, 2) & "','WO','0','" & eventID & "',GETDATE())"
73                Call objHT.InsertData_Bengkel(sSQL)
74            End If
75        End If
76        Set rsChassing = objO.FetchData("SELECT * FROM CHASSINGFORM WHERE NOWO = '" & Trim(aWO) & "'")
          
77        If rsChassing.EOF Then
              'INSERT CHASSING FORM
78            Call objHT.InsertData_Bengkel("INSERT INTO CHASSINGFORM([NoWo], [NoPolisi], [Nofak], [TanggalIn], [TanggalPanggil], [TanggalWO]) VALUES ('" & Trim(aWO) & "','" & NoPolisi & "','" & NoFak & "','" & tanggalIn & "','" & tanggalPanggil & "','" & tanggalWo & "')")
79        Else
              'UPDATE CHASSING FORM
80            Call objHT.UpdateData_Bengkel("UPDATE CHASSINGFORM SET NOPOLISI = '" & NoPolisi & "',NOFAK = '" & NoFak & "',TANGGALIN = '" & tanggalIn & "',TANGGALPANGGIL = '" & tanggalPanggil & "',TANGGALWO = '" & tanggalWo & "' WHERE NOWO = '" & Trim(aWO) & "'")
81        End If
          
          
          'mengambil data chassing form
82        Set rsChassing = objO.FetchData("SELECT * FROM CHASSINGFORM with (nolock) WHERE NOWO = '" & Trim(aWO) & "'")
          
83        If rsChassing.EOF Then
              'INSERT CHASSING FORM
84            Call objHT.InsertData_Bengkel("INSERT INTO CHASSINGFORM([NoWo], [NoPolisi], [Nofak], [TanggalIn], [TanggalPanggil], [TanggalWO]) VALUES ('" & Trim(aWO) & "','" & NoPolisi & "','" & NoFak & "','" & tanggalIn & "','" & tanggalPanggil & "','" & tanggalWo & "')")
85        Else
              'UPDATE CHASSING FORM
86            Call objHT.UpdateData_Bengkel("UPDATE CHASSINGFORM SET NOPOLISI = '" & NoPolisi & "',NOFAK = '" & NoFak & "',TANGGALIN = '" & tanggalIn & "',TANGGALPANGGIL = '" & tanggalPanggil & "',TANGGALWO = '" & tanggalWo & "' WHERE NOWO = '" & Trim(aWO) & "'")
87        End If
88        Exit Sub
vb_error:

89        erTranAll = True
90        ErrMsg.judul = "Simpan Antrian"
91        ErrMsg.message = ErrMessage(Erl, err.Number, "Procedure : Simpan Antrian")
92        ErrMsg.show
          
          ''MsgBox ErrMessage(Erl, err.Number, "Procedure : Simpan Antrian"), vbInformation
End Sub

Public Sub removeNextDayBooking(nBooking As String, isDel As Boolean)
1     On Error GoTo vb_error
2         If isDel = True Then
3             Call objHT.UpdateData_Bengkel("update " & Log.LogQueue & "..DataBooking Set TglBok = DATEADD(day, -30, TglBok) Where  NoBok = '" & nBooking & "' And Aktiv = 0")
4         Else
5             Call objHT.UpdateData_Bengkel("update " & Log.LogQueue & "..DataBooking Set TglBok = DATEADD(day, 30, TglBok) Where  NoBok = '" & nBooking & "' And Aktiv = 0")
6         End If
7         Exit Sub
vb_error:
8         erTranAll = True
9         ErrMsg.judul = ""
10        ErrMsg.message = ErrMessage(Erl, err.Number, "Procedure : Remove Next Day Booking")
11        ErrMsg.show
          
      ''MsgBox ErrMessage(Erl, err.Number, "Procedure : Remove Next Day Booking"), vbInformation
End Sub

Public Sub transferHistoryLama(nRangka As String, rsLama As ADODB.Recordset)
Dim rsSearchLama As ADODB.Recordset
Dim rsSearchHelp As ADODB.Recordset
Dim rsSearchDealer As ADODB.Recordset
Dim rsSearchLama2 As ADODB.Recordset
Dim nCount As Integer
Dim nKdDealer As String
Dim nNmDealer As String
Dim nTot As Double
Dim nKlas As String

    nKdDealer = Trim(gKdCbg)
    nNmDealer = Trim(gNmCbg)
        Set rsSearchLama = objO.FetchDataAAS("SELECT * FROM HISTORY A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_CHASSIS = '" & Trim(Nul(nRangka)) & "' ORDER BY DATE")
        If Not rsSearchLama.EOF Then
            nTot = CDbl(rsSearchLama.RecordCount)
            Do While Not rsSearchLama.EOF
                nCount = nCount + 1
                    ' CARI KLAS MODEL
                    
                    If Log.RPLHistori = 0 Then
                        Set rsSearchHelp = objO.FetchData("SELECT MODEL FROM MODEL WITH (NOLOCK) WHERE MDL = '" & Trim(rsSearchLama("MDL")) & "'")
                    Else
                        Set rsSearchHelp = objO.FetchData("SELECT MODEL FROM [" & dbReplikasiHO & "]..MODEL WITH (NOLOCK) WHERE MDL = '" & Trim(rsSearchLama("MDL")) & "'")
                    End If
                    If Not rsSearchHelp.EOF Then
                        nKlas = Trim(rsSearchHelp("MODEL"))
                    Else
                        nKlas = "ECO"
                    End If
                    
                    ' INPUT HISTORY
                    Set rsSearchDealer = objO.FetchData("SELECT * FROM HISTORY WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsSearchLama("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsSearchLama("NO_WO")) & "'")
                    If rsSearchDealer.EOF Then
                       Call objO.InsertData("INSERT INTO HISTORY (CUSTOMER,CUST_ADRES,KOTA,NO_POLISI,NO_ENGINE,NO_CHASSIS,FRONT,MECH,KM,DATE,PART_G,LABOR,TOTAL,REP1,REP2,REP3,REP4,REP5,REP6,REP7,MEKANIK1,MEKANIK2,MEKANIK3,MEKANIK4,MEKANIK5,MEKANIK6,MEKANIK7,MEKANIK8,NO_WO,NOINV,STATUS,PAID,REMARK1,FS_RC,FS_KD_DEALER,FS_NM_DEALER) " & _
                                            "VALUES ('" & Trim(Nul(rsSearchLama("CUSTOMER"))) & "','" & Trim(Nul(rsSearchLama("CUST_ADRES"))) & "','" & Trim(Nul(rsSearchLama("KOTA"))) & "','" & Trim(Nul(rsSearchLama("NO_POLISI"))) & "','" & Trim(Nul(rsSearchLama("NO_ENGINE"))) & "','" & Trim(Nul(rsSearchLama("NO_CHASSIS"))) & "','" & Trim(Nul(rsSearchLama("FRONT"))) & "','" & Trim(Nul(rsSearchLama("MECH"))) & "', " & _
                                            "'" & Trim(NulA(rsSearchLama("KM"))) & "','" & Format(NulTgl(rsSearchLama("DATE")), "MM/DD/YYYY") & "','" & setKoma(CDbl(NulA(rsSearchLama("PART")))) & "','" & setKoma(CDbl(NulA(rsSearchLama("LABOR")))) & "','" & setKoma(CDbl(NulA(rsSearchLama("TOTAL")))) & "','" & Trim(Nul(rsSearchLama("REP1"))) & "','" & Trim(Nul(rsSearchLama("REP2"))) & "','" & Trim(Nul(rsSearchLama("REP3"))) & "','" & Trim(Nul(rsSearchLama("REP4"))) & "', " & _
                                            "'" & Trim(Nul(rsSearchLama("REP5"))) & "','" & Trim(Nul(rsSearchLama("REP6"))) & "','" & Trim(Nul(rsSearchLama("REP7"))) & "','" & Trim(Nul(rsSearchLama("MEKANIK1"))) & "','" & Trim(Nul(rsSearchLama("MEKANIK2"))) & "','" & Trim(Nul(rsSearchLama("MEKANIK3"))) & "','" & Trim(Nul(rsSearchLama("MEKANIK4"))) & "','" & Trim(Nul(rsSearchLama("MEKANIK5"))) & "','" & Trim(Nul(rsSearchLama("MEKANIK6"))) & "'," & _
                                            "'" & Trim(Nul(rsSearchLama("MEKANIK7"))) & "','" & Trim(Nul(rsSearchLama("MEKANIK8"))) & "','" & Trim(Nul(rsSearchLama("NO_WO"))) & "','" & Trim(Nul(rsSearchLama("NOINV"))) & "','" & Trim(Nul(rsSearchLama("STATUS"))) & "','" & Trim(Nul(rsSearchLama("PAID"))) & "','" & Trim(Nul(rsSearchLama("REMARK1"))) & "','" & Trim(Nul(rsSearchLama("FS_RC"))) & "','" & Trim(nKdDealer) & "','" & Trim(nNmDealer) & "')")
                    End If
                    
                    ' INPUT T_PART
                    Set rsSearchLama2 = objO.FetchDataAAS("SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsSearchLama("NO_WO")) & "'")
                    If Not rsSearchLama2.EOF Then
                        Do While Not rsSearchLama2.EOF
                            Set rsSearchDealer = objO.FetchData("SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsSearchLama("NO_WO")) & "' AND KODEITEM = '" & Trim(rsSearchLama2("KODEITEM")) & "'")
                            If rsSearchDealer.EOF Then
                               Call objO.InsertData("INSERT INTO T_PART (NO_WO,KODEITEM,KETERANGAN,QTY,HARGA,OLI,NOINV,FN_EXT_DISKON) " & _
                                                    "VALUES ('" & Trim(Nul(rsSearchLama2("NO_WO"))) & "','" & Trim(Nul(rsSearchLama2("KODEITEM"))) & "','" & Trim(Nul(rsSearchLama2("KETERANGAN"))) & "', " & setKoma(CDbl(NulA(rsSearchLama2("QTY")))) & "," & setKoma(CDbl(NulA(rsSearchLama2("HARGA")))) & ",'" & Trim(IIf(rsSearchLama2("OLI"), 1, 0)) & "','" & Trim(Nul(rsSearchLama2("NOINV"))) & "','" & Trim(NulA(rsSearchLama2("FN_EXT_DISKON"))) & "')")
                            End If
                            rsSearchLama2.MoveNext
                        Loop
                    End If
                    
                    ' HITUNG NEXTEXTIMASI
                    Dim nKMNext As Double
                    Dim nSbeNext As Double
                    
                    If nCount = nTot Then
                        Dim nRS As ADODB.Recordset
                        Set nRS = objO.FetchData("Select top 1 NoRangka From M_Estimasi WITH (NOLOCK) Where NoRangka = '" & Trim(rsSearchLama("NO_CHASSIS")) & "'")
                        If nRS.RecordCount = 0 Then
                            Set rsSearchLama2 = objO.FetchDataAAS("SELECT FS_RC, DATE, KM, NOINV, NO_WO FROM HISTORY WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsSearchLama("NO_CHASSIS")) & "' AND KM <> '' ORDER BY FS_RC DESC, DATE DESC")
                            If Not rsSearchLama2.EOF Then
                                If Trim(rsSearchLama2("FS_RC")) <> "" Then
                                    nSbeNext = CDbl(Right(rsSearchLama2("FS_RC"), 3)) * 1000
                                Else
                                    nSbeNext = Round(CDbl(rsSearchLama2("KM")) / 10000) * 10000
                                End If
                                'Call objO.UpdateData("EXEC NextEstimasi '" & Trim(rsSearchLama("NO_CHASSIS")) & "', '" & Format(Trim(rsSearchLama2("DATE")), "MM/DD/YYYY") & "', '" & CDbl(Replace(rsSearchLama2("KM"), ".", "")) & "', '" & CDbl(nSbeNext) & "', '10000', '" & Trim(rsSearchLama2("NOINV")) & "', '" & Trim(rsSearchLama2("NO_WO")) & "','" & Trim(Nul(rsSearchLama("CUSTOMER"))) & "', '" & Trim(Nul(rsSearchLama("CUST_ADRES"))) & "','" & Trim(Nul(rsSearchLama("NO_POLISI"))) & "', '" & Trim(Nul(rsSearchLama("TELPON"))) & "', '', ''")
                                Call objO.UpdateData("EXEC " & Trim(hTipePenghitunganDealer) & " '" & Trim(rsSearchLama("NO_CHASSIS")) & "', '" & Format(Trim(rsSearchLama2("DATE")), "MM/DD/YYYY") & "', '" & CDbl(Replace(rsSearchLama2("KM"), ".", "")) & "', '" & CDbl(nSbeNext) & "', '10000', '" & Trim(rsSearchLama2("NOINV")) & "', '" & Trim(rsSearchLama2("NO_WO")) & "','" & Trim(Nul(rsSearchLama("CUSTOMER"))) & "', '" & Trim(Nul(rsSearchLama("CUST_ADRES"))) & "','" & Trim(Nul(rsSearchLama("NO_POLISI"))) & "', '" & Trim(Nul(rsSearchLama("TELPON"))) & "', '', ''")
                            End If
                        End If
                    End If
                rsSearchLama.MoveNext
            Loop
        End If
End Sub

Public Sub dropTempTable()
    
    Dim rsDrop As ADODB.Recordset
    
    Set rsDrop = objO.FetchData("select * from sysobjects WHERE NAME LIKE 'INVOICE%' AND XTYPE  = 'U'")
    If Not rsDrop.EOF Then
        Do While Not rsDrop.EOF
            Call objO.DeleteData("DROP TABLE [DBO].[" & Trim(rsDrop("NAME")) & "]")
            rsDrop.MoveNext
        Loop
    End If
    If rsDrop.State Then rsDrop.Close
End Sub

Public Function birthDayReminder(noRangka As String) As String
    Dim rsBirth As ADODB.Recordset
    birthDayReminder = ""
    Set rsBirth = objO.FetchData("SELECT USETGLLAHIR FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "'")
    If Not rsBirth.EOF Then
        If TglDDMMYYYY(rsBirth("USETGLLAHIR")) <> "01-01-1900" Then
            birthDayReminder = TglDDMMYYYY(rsBirth("USETGLLAHIR"))
        Else
            birthDayReminder = ""
        End If
    End If
End Function

Public Function cariNamaBirth(noRangka As String) As String
    
    Dim rsBirth As ADODB.Recordset
    
    cariNamaBirth = ""
    Set rsBirth = objO.FetchData("SELECT USENAME FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "'")
    If Not rsBirth.EOF Then
        cariNamaBirth = Trim(rsBirth("USENAME"))
    End If
End Function

Public Function okCallFinishService(noWO As String) As Boolean
    
    Dim rsInvoice As ADODB.Recordset
    Dim rsFinish As ADODB.Recordset
    okCallFinishService = False
    
    If gQueue Then
        Set rsInvoice = objO.FetchData("SELECT * FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(noWO) & "' AND NOINV <> ''")
        If Not rsInvoice.EOF Then
            Set rsFinish = objO.FetchData("SELECT * FROM " & Log.LogQueue & "..FINISHSERVICE WITH (NOLOCK) WHERE NOPOL = '" & Trim(rsInvoice("NO_POLISI")) & "' AND TGLFINISH = '" & Format(TglYYYYMMDD(rsInvoice("INVDATE")), "MM/DD/YYYY") & "'")
            If Not rsFinish.EOF Then
                If Trim(rsFinish("PANGGIL")) <> "0" Then
                    okCallFinishService = True
                End If
            Else
                'If Trim(rsInvoice("THS")) <> "0" Or Trim(rsInvoice("PDS")) <> "0" Or Trim(rsInvoice("TPSVC")) = "W" Then
                okCallFinishService = True
                'End If
            End If
        End If
    Else
        okCallFinishService = True
    End If
End Function

Public Function newTipeFakturPart(noWO As String) As Boolean
    
    On Error GoTo vbError
    Dim rsSearch As ADODB.Recordset
    
    newTipeFakturPart = False
     
    Set rsSearch = objO.FetchData("SELECT * From syscolumns where id in (select id from sysobjects where id = object_id(N'[dbo].[JOBDTA]')  ) AND [NAME] = 'NOFK'")
    If Not rsSearch.EOF Then
        newTipeFakturPart = True
    End If
     
    Set rsSearch = Nothing
    Exit Function
    
vbError:
    newTipeFakturPart = False
    Exit Function
End Function

Public Function FakturDb(noWO As String) As String
    
    On Error GoTo vbError
    Dim rsSearch As ADODB.Recordset
    FakturDb = ""
    
    Set rsSearch = objO.FetchData("SELECT ISNULL(DBASEFK,'') AS DBASEFK FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(noWO) & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND IsNull(NOFKRETUR,'') = ''")
    If Not rsSearch.EOF Then
        FakturDb = Trim(rsSearch("DBASEFK"))
    End If
    Set rsSearch = Nothing
    Exit Function
    
vbError:
    FakturDb = ""
    Exit Function
End Function

Public Function partOkInvoice(noWO As String) As Boolean
    
    On Error GoTo vbError
    Dim rsSearch As ADODB.Recordset
    partOkInvoice = False
    
    Set rsSearch = objO.FetchData("SELECT isnull(NOFK,'') AS NOFK, isnull(NOFKRETUR,'') AS NOFKRETUR FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(noWO) & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = ''")
    If Not rsSearch.EOF Then
        If Trim(rsSearch("NOFKRETUR")) = "" Then
            partOkInvoice = True
        End If
    End If
    Set rsSearch = Nothing
    Exit Function
    
vbError:
    partOkInvoice = False
    Exit Function
End Function

'Public Function AdaPemakaianSublet(NoWo As String) As Boolean
'Dim rsSearch As New ADodb.Recordset
'Dim strSQL As String
'
'    On Error GoTo vbError
'    AdaPemakaianSublet = False
'
'        strSQL = "SELECT No_Wo FROM jobSublet WHERE no_WO = '" & NoWo & "' and valid = '1'"
'        Set rsSearch = objO.FetchData(strSQL)
'        If Not rsSearch.EOF Then
'            AdaPemakaianSublet = True
'        End If
'
'        Set rsSearch = Nothing
'
'    Exit Function
'
'vbError:
'    AdaPemakaianSublet = False
'    Exit Function
'End Function

Public Function AdaPemakaianPart(noWO As String) As Boolean
Dim rsSearch As New ADODB.Recordset
Dim strSql As String

    On Error GoTo vbError
    AdaPemakaianPart = False
         
        strSql = "Select Z.* FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN  where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                "(INMST.HARGAJUALA ) as HARGAJUALA, " & _
                "INMST.NORACK,0 as DisItem, 0 as ExtDis FROM INMST " & _
                "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN  where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                "Where Z.Norack = '" & Trim(noWO) & "'"
        Set rsSearch = objO.FetchDataACC(strSql)
        If Not rsSearch.EOF Then
            AdaPemakaianPart = True
        End If
         
        Set rsSearch = Nothing
    
    Exit Function
    
vbError:
    AdaPemakaianPart = False
    Exit Function
End Function


Public Function AdaPemakaianSublet(noWO As String) As Boolean
Dim rsSearch As New ADODB.Recordset
Dim strSql As String

    On Error GoTo vbError
    AdaPemakaianSublet = False
         
        strSql = "Select valid from JobSublet WITH (NOLOCK) Where Valid = 1 and No_WO = '" & Trim(noWO) & "'"
        Set rsSearch = objO.FetchData(strSql)
        If Not rsSearch.EOF Then
            AdaPemakaianSublet = True
        End If

        Set rsSearch = Nothing
    Exit Function
    
vbError:
    AdaPemakaianSublet = False
    Exit Function
End Function

Public Function AdaInvoicePart(noWO As String) As Boolean
    
    On Error GoTo vbError
    Dim rsSearch As New ADODB.Recordset
    AdaInvoicePart = False
    
    Dim nFDB As String
    nFDB = FakturDb(Trim(noWO))
    If nFDB <> "" Then
        If Get_ConnectionNisSP(nFDB, NisSPUSER, NisSPPass, NisSPKomp) Then
             
             
            'Set rsSearch = objO.FetchDataACC("SELECT NOITEM AS KODEITEM, NAMABARANG, KODESATUAN, ABS(qtysaldoawal+qtyterimabln+qtykeluarbln+abs(qtyjualbln)) AS JML, CASE WHEN ABS(JLHJUALBLN) > 0 THEN ROUND((ABS(JLHJUALBLN)/ABS(qtysaldoawal+qtyterimabln+qtykeluarbln+abs(qtyjualbln))) * 110/100,0) ELSE HARGAJUALA END AS HARGAJUALA, ABS(JLHJUALBLN) AS JLHJUALBLN FROM INMST WHERE NORACK = '" & Trim(NoWo) & "' and KODEGDG = '" & Trim(gGudang) & "' AND (qtysaldoawal+qtyterimabln+qtykeluarbln+abs(qtyjualbln)) > 0")
            Set rsSearch = objO.FetchDataACCX("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
                                           & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(noWO) & "'", SQLConACC1)
            If Not rsSearch.EOF Then
                AdaInvoicePart = True
            End If

        End If
    End If
    Exit Function
    
vbError:
    AdaInvoicePart = False
    Exit Function
End Function

Public Function cariNisDb(noWO As String) As String
    
    On Error GoTo vbError
    Dim rsDb As ADODB.Recordset
    
    cariNisDb = ""
    Set rsDb = objO.FetchData("SELECT DBASEFK FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(noWO) & "'")
    If Not rsDb.EOF Then
        cariNisDb = Trim(rsDb("DBASEFK"))
    End If
    If rsDb.State <> adStateClosed Then rsDb.Close
    Set rsDb = Nothing
    Exit Function
    
vbError:
    cariNisDb = ""
    Exit Function
End Function

Public Function updateGLtrn()
On Error GoTo vbError
Dim strSql As String
Dim rsSearch As New ADODB.Recordset
Set rsSearch = obj.searchGLtrnNonWo

If Not rsSearch.EOF Then
    Dim cm As New Command
    Dim uConn As New ADODB.Connection
    If uConn.State = adStateClosed Then uConn.Open (SQLConACC)

    cm.CommandText = "UPDATE GLTRN SET NOWO = B.NO_WO, NOFPS = C.NOFPS FROM GLTRN A WITH (NOLOCK) LEFT JOIN " & Log.LogBengkel & "..T_INVOICE B WITH (NOLOCK) ON B.NOINV = SUBSTRING(REPLACE(A.NOBUKTI,'',''),3, LEN(REPLACE(A.NOBUKTI,'',''))-2) " & _
                              "LEFT JOIN " & Log.LogFaktur & "..INVOICE C WITH (NOLOCK) ON REPLACE(A.NOBUKTI,'','') = C.NOINVOICE " & _
                              "WHERE KODESUMBER = 'SV' AND NOBUKTI LIKE 'E-%' AND (ISNULL(A.NOFPS,'') = '' OR ISNULL(A.NOWO,'') = '')"
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = uConn
    cm.Execute
    
    Set cm = Nothing
    If uConn.State = adStateOpen Then uConn.Close
End If
vbError:
    Exit Function
End Function

Public Function JTBayar(kodePelanggan As String, vTgl As String) As String
Dim rsSearch As New ADODB.Recordset
Dim nRangeSyarat As Double
    On Error GoTo vbError
    JTBayar = ""
        Set rsSearch = obj.searchSyaratBayar(kodePelanggan)
        If Not rsSearch.EOF Then
        Select Case Left(Trim(rsSearch("SYARATBAYAR")), 1)
            Case "D"
                nRangeSyarat = Val(Mid(Trim(rsSearch("SYARATBAYAR")), 2, 5))
            Case "W"
                nRangeSyarat = Val(Mid(Trim(rsSearch("SYARATBAYAR")), 2, 5)) * 7
            Case "M"
                nRangeSyarat = Val(Mid(Trim(rsSearch("SYARATBAYAR")), 2, 5)) * 30
        End Select
            'JTBayar = TglDDMMYYYY(DateAdd(LCase(Left(rsSearch("SYARATBAYAR"), 1)), Val(Right(rsSearch("SYARATBAYAR"), Len(rsSearch("SYARATBAYAR")) - 1)), vTgl))
            JTBayar = TglDDMMYYYY(DateAdd("D", nRangeSyarat, vTgl))
        End If
    Exit Function
vbError:
    JTBayar = ""
    Exit Function
End Function

Public Function panggilLastPDS(noRgk As String) As Boolean
Dim rsSearch As ADODB.Recordset

    On Error GoTo vbError
    'panggilLastSBE = objO.FetchData_Double("SELECT TOP 1 FS_NILAI FROM HISTORY A INNER JOIN T_REGULAR B ON A.FS_RC = B.FS_RC WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND a.FS_RC <> '' ORDER BY DATE DESC")  '11/05/12
    If Log.RPLHistori = 0 Then
        panggilLastPDS = objO.FetchData_Boolean("SELECT TOP 1 A.PDS FROM HISTORY A WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND a.PDS = '1' AND FS_LOG_DEL = '' ORDER BY DATE DESC")
    Else
        panggilLastPDS = objO.FetchData_Boolean("SELECT TOP 1 A.PDS FROM [" & dbReplikasiHO & "]..HISTORYnew A WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND a.PDS = '1' AND FS_LOG_DEL = '' ORDER BY DATE DESC")
    End If
    Exit Function
    
vbError:
    panggilLastPDS = False
    Exit Function
End Function

Public Function panggilLastSBE(noRgk As String, noPols As String) As Double
Dim rsSearch As ADODB.Recordset

    On Error GoTo vbError
    'panggilLastSBE = objO.FetchData_Double("SELECT TOP 1 FS_NILAI FROM HISTORY A INNER JOIN T_REGULAR B ON A.FS_RC = B.FS_RC WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND a.FS_RC <> '' ORDER BY DATE DESC")  '11/05/12
    If Log.RPLHistori = 0 Then
        panggilLastSBE = objO.FetchData_Double("SELECT TOP 1 FS_NILAI FROM HISTORY A WITH (NOLOCK) INNER JOIN T_REGULAR B WITH (NOLOCK) ON A.FS_RC = B.FS_RC WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND a.FS_RC <> '' AND FS_LOG_DEL = '' ORDER BY DATE DESC")
    Else
        ''panggilLastSBE = objO.FetchData_Double("SELECT TOP 1 FS_NILAI FROM [" & dbReplikasiHO & "]..HISTORYnew A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..T_REGULAR B WITH (NOLOCK) ON A.FS_RC = B.FS_RC WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND a.FS_RC <> '' AND FS_LOG_DEL = '' ORDER BY DATE DESC")
        panggilLastSBE = objO.FetchData_Double("SELECT TOP 1 FS_NILAI FROM [" & dbReplikasiHO & "]..T_INVOICEnew A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..T_REGULAR B WITH (NOLOCK) ON A.FS_RC = B.FS_RC WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND a.FS_RC <> '' AND FS_LOG_DEL = '' and NOINV NOT LIKE '%-B' ORDER BY INVDATE DESC")
    End If
    Exit Function
    
vbError:
    panggilLastSBE = 0
    Exit Function
End Function

Public Function dataCustomerLengkap(noRgk As String) As Boolean
Dim rsSearch As ADODB.Recordset

    On Error GoTo vbError
    dataCustomerLengkap = False
    Set rsSearch = objO.FetchData("SELECT isnull(UsePSTN,'') as PSTN, isnull(UseGSM,'') as GSM, isnull(UseCDMA,'') as CDMA from Custdat WITH (NOLOCK) where No_Chassis = '" & Trim(noRgk) & "'")
    If Not rsSearch.EOF Then
        If Trim(rsSearch("PSTN")) = "" And Trim(rsSearch("GSM")) = "" And Trim(rsSearch("CDMA")) = "" Then
            dataCustomerLengkap = False
        Else
            dataCustomerLengkap = True
        End If
    End If
    Exit Function
    
vbError:
    dataCustomerLengkap = False
    Exit Function
End Function

Public Function cekEvent(noRgk As String, noPols As String) As String

    On Error GoTo vbError
    Dim rsEvent As New ADODB.Recordset
    Dim rsUnit As New ADODB.Recordset
    Dim kriteria As String
    Dim strSql As String
    Dim i As Integer
    
    cekEvent = ""
    If rsEvent.State Then rsEvent.Close
    Set rsEvent = support.searchEvent(Format(isnow, "MM/DD/YYYY")) 'oConn.Execute("SELECT * FROM HO_EVENT WHERE DATEDIFF(d,BERLAKU1,'" & Format(isnow, "MM/DD/YYYY") & "') > 0 AND DATEDIFF(d,'" & Format(isnow, "MM/DD/YYYY") & "',BERLAKU2) > 0 and AKTIF = 1 ORDER BY noEvent")
    If Not rsEvent.EOF Then
        Do While Not rsEvent.EOF
            For i = 0 To 8
                kriteria = ""
                Select Case i
                    Case 0
                        Select Case rsEvent("tipeOrder")
                            Case "Booking"
                                kriteria = " And Nofak <> ''"
                            Case "Non Booking"
                                kriteria = " And Nofak = ''"
                            Case "All", ""
                                kriteria = ""
                        End Select
                    Case 1
                        If CDbl(rsEvent("thVehicle1")) <> 0 And CDbl(rsEvent("thVehicle2")) <> 0 Then
                            kriteria = " And (FS_TH_MOBIL between '" & CDbl(rsEvent("thVehicle1")) & "' And '" & CDbl(rsEvent("thVehicle2")) & "')"
                        End If
                    Case 2
                        If CDbl(rsEvent("KM1")) <> 0 And CDbl(rsEvent("KM2")) <> 0 Then
                            kriteria = " And (FS_KM_AKHIR between '" & CDbl(rsEvent("KM1")) & "' And '" & CDbl(rsEvent("KM2")) & "')"
                        End If
                    Case 3
                        If CDbl(rsEvent("lastService1")) <> 0 Or CDbl(rsEvent("lastService2")) <> 0 Then
                            kriteria = " AND (ISNULL((SELECT TOP 1 CASE ISNULL(DATEDIFF(day, [DATE], GETDATE()),0) WHEN '' THEN 0 ELSE ISNULL(DATEDIFF(day, [DATE], GETDATE()),0) END FROM HISTORY WITH (NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS ORDER BY [DATE] DESC),0) BETWEEN " & CDbl(rsEvent("lastService1")) & " AND " & CDbl(rsEvent("lastService2")) & ")"
                        End If
                    Case 4
                        If Format(rsEvent("birthDay1"), "MM/DD/YYYY") <> "01/01/1900" And Format(rsEvent("birthDay2"), "MM/DD/YYYY") <> "01/01/1900" Then
                            kriteria = " And (USETGLLAHIR between '" & Format(rsEvent("birthDay1"), "MM/DD/YYYY") & "' And '" & Format(rsEvent("birthDay2"), "MM/DD/YYYY") & "')"
                        End If
                    Case 5
                        If Trim(rsEvent("hobby")) <> "" Then
                            kriteria = " And USEFS_NM_HOBI = '" & Trim(rsEvent("HOBBY")) & "'"
                        End If
                    Case 6
                        If Format(rsEvent("delivDay1"), "MM/DD/YYYY") <> "01/01/1900" And Format(rsEvent("delivDay2"), "MM/DD/YYYY") <> "01/01/1900" Then
                            kriteria = " And (CAST((SUBSTRING(ddate,4,2)+'/'+SUBSTRING(ddate,1,2)+'/'+SUBSTRING(ddate,7,4)) AS DATETIME) between '" & Format(rsEvent("delivDay1"), "MM/DD/YYYY") & "' And '" & Format(rsEvent("delivDay2"), "MM/DD/YYYY") & "')"
                        End If
                    Case 7
                        Select Case Trim(rsEvent("modelVehichle"))
                            Case "V"
                                kriteria = " And MDL IN (" & support.searchModelEvent(rsEvent("noEvent")) & ")"
                            Case "A", ""
                                kriteria = ""
                        End Select
                    Case 8
                        kriteria = ""
                End Select
                strSql = strSql + kriteria
            Next
            If rsUnit.State Then rsUnit.Close
            Set rsUnit = support.searchKriteriaEvent(strSql, Trim(noRgk))
            If Not rsUnit.EOF Then
                ' MASUK KRITERIA EVENT
                cekEvent = cekEvent & Trim(rsEvent("TEXTEVENT")) & "|" & Trim(rsEvent("noEVENT")) & vbCrLf
            End If
            rsEvent.MoveNext
        Loop
    End If
    If rsEvent.State <> adStateClosed Then rsEvent.Close
    Set rsEvent = Nothing
    Exit Function
    
vbError:
    cekEvent = ""
    Exit Function
End Function

Public Sub adaNoShow()
      Dim rsHlp As ADODB.Recordset
      Dim rsAct As ADODB.Recordset
      Dim rsHelpJpcb As ADODB.Recordset
      Dim pesan As String
      Dim hlp As String
      Dim nTotAct As Integer
      Dim i As Integer
          Dim rsSO As ADODB.Recordset
          Dim rsSOL As ADODB.Recordset
          Dim noBookingAppOnline As String
          Dim nobookingdidx As String
          Dim evenID As String
          
          
1         On Error GoTo vbError
2         pesan = ""
3         Set rsHlp = support.searchNoShow(gNoShow)  'oConn.Execute("SELECT * FROM BOOK WHERE [DATE] <= DateAdd(n, (" & gNoShow & " * -1), getdate()) AND [TIME] <= convert(varchar(5), DateAdd(n, (" & gNoShow & " * -1), getdate()),8) AND DELIV <> 'BTL' AND DELIV <> 'OKE' AND NOFAK NOT IN (SELECT NOBOK FROM " & Log.LogQueue & "..DATABOOKING WHERE AKTIV <> 0)")
4         If Not rsHlp.EOF Then
5             Do While Not rsHlp.EOF
6                 If Log.isFill = 1 And Log.startAppRate <= Format(rsHlp("DATE"), "DD/MM/YYYY") Then
7                     Set rsSO = objO.FetchDataACC("select Sum(QtyReserve) from SOORD WITH (NOLOCK) Where NOSPK = '" & Trim(rsHlp("NOFAK")) & "'")
8                     If Val(NulA(rsSO.Fields(0).value)) > 0 Then
9                         Set rsSOL = objO.FetchData("Select NoFak from DelBook WITH (NOLOCK) Where Hapus = 0 and NoFak = '" & Trim(rsHlp("NOFAK")) & "'")
10                        If rsSOL.RecordCount = 0 Then
11                            Call objO.InsertData("Insert Into DelBook values ('" & Trim(rsHlp("NOFAK")) & "',getdate(),0,'01/01/1900')")
12                        End If
13                    End If
                      
                      
14                    noBookingAppOnline = ""
15                    If NulSID(rsHlp("KODETIPE")) <> "" Then
                          '' isi nobookingapp dan kodeprogram
16                        If InStr(1, NulSID(rsHlp("KODETIPE")), "#|#") > 0 Then
17                            noBookingAppOnline = Left(NulSID(rsHlp("KODETIPE")), InStr(1, NulSID(rsHlp("KODETIPE")), "#|#") - 1)
18                        End If
19                    End If
                      
                      
20                    If Get_ConnectionProfil_Bengkel99 = True And Log.DiDxP1 = 1 Then
21                        nobookingdidx = ""
22                        If noBookingAppOnline <> "" Then
23                            nobookingdidx = objO.FetchData_String("SELECT booking_ID from [" & dbReplikasiHO & "]..bookingonline_didx WHERE nobookingapps = '" & noBookingAppOnline & "'")
24                            evenID = objO.FetchData_String("select event_id from [" & dbReplikasiHO & "]..bookingonline_didx_toSend where nobookingapps = '" & NulSID(noBookingAppOnline) & "'")
                          
25                        End If
26                        If evenID <> "" Then
27                            Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & noBookingAppOnline & "','" & noBookingAppOnline & "','','" & Right(gDb, 2) & "','CANCELLED','0','" & evenID & "',getdate())")
28                            Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..bookingonline_didx_toSend SET booking_status = 'CANCELLED',cancellation_reason = 'NO SHOW' WHERE event_id = '" & evenID & "' ")
29                        Else
30                            Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & Right(gDb, 2) & "." & NulSID(Trim(rsHlp("NOFAK"))) & "','" & Right(gDb, 2) & "." & NulSID(Trim(rsHlp("NOFAK"))) & "','','" & Right(gDb, 2) & "','CANCELLED','0','',getdate())")
31                            Call objO.InsertData("UPDATE [" & dbReplikasiHO & "]..bookingonline_didx_toSend SET booking_status = 'CANCELLED',cancellation_reason = 'NO SHOW' WHERE booking_id = '" & Right(gDb, 2) & "." & NulSID(Trim(rsHlp("NOFAK"))) & "' ")
                          
32                        End If
33                    End If
                      
                      
34                    Call objO.DeleteData("Update DelBook Set Hapus = 1, TglHapus = getdate() Where NOFAK = '" & Trim(rsHlp("NOFAK")) & "'")
35                    Call support.canceledNoShow(gNoShow, Trim(rsHlp("NOFAK")))
                      ''---------------------------------------------------------------------------
                      ''2017-04-26 || penambahan help jpcb
                      ''---------------------------------------------------------------------------
                     'Set rsHelpJpcb = objO.FetchData("SELECT FROM WITH (NOLOCK) WHERE TANGGAL <= convert(varchar(10), DateAdd(n, (" & minuteNoShow & " * -1), getdate()),1) AND KETERANGAN LIKE '%|" & Trim(rsHlp("NOFAK")) & "|%'")
36                    If Log.helpJpcb = 1 Then
                          Dim rsHlpJPCB As ADODB.Recordset
                          Dim rsLogJPCB As ADODB.Recordset
37                        Set rsHelpJpcb = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE TANGGAL <= convert(varchar(10), DateAdd(n, (" & gNoShow & " * -1), getdate()),1) AND KETERANGAN LIKE '%|" & Trim(rsHlp("NOFAK")) & "|%'")
                          
38                        Set rsLogJPCB = objO.FetchData("SELECT * FROM PROFILBENGKEL.DBO.LOGJPCB WITH (NOLOCK)")
39                        Set rsHlpJPCB = objO.FetchData("SELECT TOP 1 * FROM HELPJPCB WITH (NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(rsHlp("NOFAK")) & "|%' AND TANGGAL <= convert(varchar(10), DateAdd(n, (" & gNoShow & " * -1), getdate()),1)")
40                        If rsHlpJPCB.EOF Then
41                             If rsLogJPCB.RecordCount > 0 Then
42                                rsLogJPCB.MoveFirst
43                                Do While Not rsLogJPCB.EOF
                                      'lakukan pengisian data help jpcb berdasarkan mech
                                      ''2016-06-15 no booking yang ada distall yang diinsert
                                      
44                                   If Not rsHelpJpcb.EOF Then Call objO.InsertData("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('" & Trim(rsHlp("NOFAK")) & "','T','S','" & UCase(Trim(rsHelpJpcb("nostall"))) & "','" & Format(rsHelpJpcb("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsHelpJpcb("JAMMULAI")) & "','" & Trim(rsHelpJpcb("JAMSELESAI")) & "'," & Replace(rsHelpJpcb("rate"), ",", ".") & ",'" & Trim(rsHelpJpcb("mech")) & "','','" & Trim(rsHelpJpcb!keterangan) & "','" & Trim(rsHelpJpcb("adjust")) & "','DELETE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
45                                   rsLogJPCB.MoveNext
46                                Loop
47                             End If
48                         Else
49                            If rsLogJPCB.RecordCount > 0 Then
50                                rsLogJPCB.MoveFirst
51                                Do While Not rsLogJPCB.EOF
                                      'lakukan pengisian data help jpcb berdasarkan mech
52                                   If Not rsHelpJpcb.EOF Then Call objO.InsertData("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('" & Trim(rsHlp!NoFak) & "','T','S','" & UCase(Trim(rsHelpJpcb("nostall"))) & "','" & Format(rsHelpJpcb("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsHelpJpcb("JAMMULAI")) & "','" & Trim(rsHelpJpcb("JAMSELESAI")) & "'," & Replace(rsHelpJpcb("rate"), ",", ".") & ",'" & Trim(rsHelpJpcb("mech")) & "','','" & Trim(rsHelpJpcb!keterangan) & "','" & Trim(rsHelpJpcb("adjust")) & "','DELETE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
53                                   rsLogJPCB.MoveNext
54                                Loop
55                             End If
       
56                        End If
57                    End If
                      ''-------- akhir penambahan helpjpcb ----------------------------------------
58                    Call updateAntrian(Trim(rsHlp("NO_POLISI")), Trim(rsHlp("FMDL")), Trim(rsHlp("CUSTOMER")), Format(rsHlp("DATE"), "DD/MM/YYYY"), Trim(rsHlp("TIME")), 3, Trim(rsHlp("NOFAK")))
59                    nTotAct = 0
60                    For i = 1 To 7
61                        hlp = "REP" & i
62                        If rsHlp(hlp) <> "" Then
                              'If rsAct.State Then rsAct.Close
63                            Set rsAct = support.searchRate(Trim(rsHlp(hlp)), Trim(rsHlp("MDL"))) 'oConn.Execute("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WHERE FS_ITEM = '" & Trim(rsHlp(hlp)) & "' AND MDL = '" & Trim(rsHlp("MDL")) & "'")
64                            If Not rsAct.EOF Then
65                                nTotAct = nTotAct + CDbl(rsAct("RATE"))
66                            End If
67                            If rsAct.State <> adStateClosed Then rsAct.Close
68                            Set rsAct = Nothing
69                        End If
70                    Next
71                    Call objO.UpdateDataACC("UPDATE SOHDR SET TUTUPORDER = 1 WHERE NOSO = '" & Trim(rsHlp("NOFAK")) & "' ")
72                    Call createMekanik
                     ' Call decresRestTime(rsHlp("MECH"), CDbl(nTotAct), 1, Format(rsHlp("DATE"), "MM/DD/YYYY"))
73                    pesan = pesan + "No. Polisi : " & rsHlp("NO_POLISI") & " No. Book : " & rsHlp("NOFAK") & vbCr
                  
74                Else
75                    Call support.canceledNoShow(gNoShow, Trim(rsHlp("NOFAK")))
                      
76                    Call updateAntrian(Trim(rsHlp("NO_POLISI")), Trim(rsHlp("FMDL")), Trim(rsHlp("CUSTOMER")), Format(rsHlp("DATE"), "DD/MM/YYYY"), Trim(rsHlp("TIME")), 3, Trim(rsHlp("NOFAK")))
77                    nTotAct = 0
78                    For i = 1 To 7
79                        hlp = "REP" & i
80                        If rsHlp(hlp) <> "" Then
                              'If rsAct.State Then rsAct.Close
81                            Set rsAct = support.searchRate(Trim(rsHlp(hlp)), Trim(rsHlp("MDL"))) 'oConn.Execute("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WHERE FS_ITEM = '" & Trim(rsHlp(hlp)) & "' AND MDL = '" & Trim(rsHlp("MDL")) & "'")
82                            If Not rsAct.EOF Then
83                                nTotAct = nTotAct + CDbl(rsAct("RATE"))
84                            End If
85                            If rsAct.State <> adStateClosed Then rsAct.Close
86                            Set rsAct = Nothing
87                        End If
88                    Next
89                    Call objO.UpdateDataACC("UPDATE SOHDR SET TUTUPORDER = 1 WHERE NOSO = '" & Trim(rsHlp("NOFAK")) & "' ")
90                    Call createMekanik
                      'Call decresRestTime(rsHlp("MECH"), CDbl(nTotAct), 1, Format(rsHlp("DATE"), "MM/DD/YYYY"))
91                    pesan = pesan + "No. Polisi : " & rsHlp("NO_POLISI") & " No. Book : " & rsHlp("NOFAK") & vbCr
92                End If
93                rsHlp.MoveNext
94            Loop
95        End If
96        Exit Sub

vbError:
97        Exit Sub
End Sub
Public Function gTglServer() As Date
    On Error GoTo vbError
         Set rs = objO.FetchData("Select GetDate() as TGL")
         isnow = rs.Fields(0).value
         gTglServer = Format(rs("TGL"), "DD/MM/YYYY HH:NN:SS")
    Exit Function
vbError:
    isnow = TglDDMMMMYYYY(Date)
    Exit Function
End Function

Public Function gTglServerBilling() As Date
    On Error GoTo vbError
         Set rs = objO.FetchData("Select GetDate() as TGL")
         isNowBilling = rs.Fields(0).value
         gTglServerBilling = Format(rs("TGL"), "DD/MM/YYYY HH:NN:SS")
    Exit Function
vbError:
    isNowBilling = TglDDMMMMYYYY(Date)
    Exit Function
End Function

Function setClose(rsT As Recordset)
    
    If rsT.State Then
        rsT.Close
        Set rsT = Nothing
    End If
End Function

Public Function Get_Connectiontemp(DBname As String, _
                               UsrName As String, _
                               Pass As String, _
                               SerName As String) As Boolean
1      On Error GoTo vb_error
    
3      Get_Connectiontemp = True
4      If oConnTemp.State Then oConnTemp.Close
5      oConnTemp.CursorLocation = adUseClient
6      oConnTemp.ConnectionString = SQLCon
       oConnTemp.ConnectionTimeout = 10
8      oConnTemp.Open
9      oConnTemp.Close
     Exit Function

vb_error:
    Get_Connectiontemp = False
End Function


Public Sub createMekanik()
          
1        On Error GoTo createMekanik_Error

2         If rsMK.State Then rsMK.Close
3         Set rsMK = Nothing
4         With rsMK
5             .Fields.Append "MECH", adVarChar, 5
6             .Fields.Append "HARI", adVarChar, 10
7             .Fields.Append "TANGGAL", adVarChar, 10
8             .Fields.Append "SHIFT", adDouble
9             .Fields.Append "JamMulai", adVarChar, 5
10            .Fields.Append "JamSelesai", adVarChar, 5
11            .Fields.Append "MenitKerja", adDouble
12            .Fields.Append "Terpakai", adDouble
13            .Fields.Append "Masuk", adVarChar, 1
14            .Fields.Append "Alokasi", adDouble
15            .Open
16        End With
          
17        rsMK.AddNew

18       On Error GoTo 0
19       Exit Sub

createMekanik_Error:
20       erTranAll = True
21       ErrMsg.judul = ""
22       ErrMsg.message = "Error " & err.Number & " (" & err.Description & ") in procedure createMekanik of Module MainModul"
23       ErrMsg.show
         
      '   MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createMekanik of Module MainModul"
End Sub

Public Function CexChassis_Discon(nameModel As String) As Boolean
1         isJualNRM = False
2         CexChassis_Discon = False
3         tglJualNRM = Format("01/01/1900", "MM/DD/YYYY")
                                
          Dim rsC As New ADODB.Recordset
          Dim mConn As New ADODB.Connection
4         On Error GoTo errHandler
5         If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
6         rsC.CursorLocation = adUseClient
7         If Log.RPLHistori = 0 Then
8             rsC.Open "SELECT noRangka,tglentry FROM salesUnit WITH (NOLOCK) Where NoRangka = '" & myChassis & "'", SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
9         Else
10            rsC.Open "select ISNULL(b.fs_kd_dealer,'') as FS_KD_DEALER,ISNULL(b.fs_nm_dealer,'') as FS_NM_DEALER,a.tglentry from [" & dbReplikasiHO & "]..Summary a with (nolock) left join [" & dbReplikasiHO & "]..t_dealer b with (nolock) on a.KodeDealer = b.kodecabang where a.tipetr = 'ND' And a.[Status] = '1' AND a.NoRangka = '" & Trim(myChassis) & "'", SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
11        End If
12        Set rsC.ActiveConnection = Nothing
          '' ===========================================================
          '' 20190907 JIKA NORANGKA KOSONG DIANGGAP NON NCM
          '' If rsC.RecordCount > 0 Then
          '' ===========================================================
13        If rsC.RecordCount > 0 And NulSID(myChassis) <> "" Then
14            CexChassis_Discon = True
15            isJualNRM = True
16            tglJualNRM = Format(rsC("TGLENTRY"), "MM/DD/YYYY")
17        End If
18        If mConn.State <> adStateClosed Then mConn.Close
19        Set mConn = Nothing
20        Set rsC = Nothing
21        Call LoadtipeDisc(nameModel)
22        Exit Function
errHandler:
23        CexChassis_Discon = False
24        Call LoadtipeDisc(nameModel)
End Function

Public Function CekUnitTerlibatFOP(noKa As String, MDL As String, ByRef tagihanFOP As String) As Boolean
          Dim hasil           As Boolean
          Dim rsCekUnit       As ADODB.Recordset
          Dim tagihan         As String
          Dim tglRetailAwal   As Date
          Dim tglRetailAkhir  As Date
          
1         hasil = True
          ''cek kode nik pada noka dengan master fop
          ''-------------------------------------------------
2         sSQL = "select FOP,RETAILAWAL,RETAILAKHIR from [" & dbReplikasiHO & "]..HO_FOP_PROGRAM where mdl = '" & NulSID(MDL) & "' and NIK = '" & Mid(NulSID(noKa), 10, 1) & "' and fs_log_del = ''"
3         Set rsCekUnit = objO.FetchData(sSQL)
4         If rsCekUnit.RecordCount > 0 Then
5             tagihan = NulSID(rsCekUnit!FOP)
6             tglRetailAwal = Format(rsCekUnit!retailawal, "dd/mm/yyyy")
7             tglRetailAkhir = Format(rsCekUnit!retailakhir, "dd/mm/yyyy")
8             tagihanFOP = tagihan
9         Else
10            hasil = False
11            GoTo ReturnHasil
12        End If

13        If tagihan = "NRM" Then
              ''cek tanggal retail unit apakah masuk kriteria
              ''---------------------------------------------------------------------
14            sSQL = "select ISNULL(b.fs_kd_dealer,'') as FS_KD_DEALER,ISNULL(b.fs_nm_dealer,'') as FS_NM_DEALER,a.tglentry from [" & dbReplikasiHO & "]..Summary a with (nolock) left join [" & dbReplikasiHO & "]..t_dealer b with (nolock) on a.KodeDealer = b.kodecabang where a.tipetr = 'ND' And a.[Status] = '1' AND a.NoRangka = '" & NulSID(noKa) & "' and a.tglentry between '" & Format(tglRetailAwal, "yyyy-mm-dd") & " 00:00:00' and '" & Format(tglRetailAkhir, "yyyy-mm-dd") & " 23:59:59'"
15            If objO.FetchData_Count(sSQL) > 0 Then
16                If tagihan = "NRM" Then
                      ''cek ke buklet apakah noka sudah terdaftar
                      ''----------------------------------------------------
17                    sSQL = "select * from [" & dbReplikasiHO & "]..buklet_sales where no_rangka = '" & NulSID(noKa) & "'"
18                    If objO.FetchData_Count(sSQL) > 0 Then
19                        hasil = True
20                        GoTo ReturnHasil
21                    Else
22                        tagihanFOP = ""
23                        hasil = False
24                        GoTo ReturnHasil
25                    End If
26                Else
27                    hasil = True
28                    GoTo ReturnHasil
29                End If
30            Else
                  ''jika bukan penjualan nrm ?
31                hasil = False
32                tagihanFOP = ""
33                GoTo ReturnHasil
34            End If
35        ElseIf tagihan = "TAM" Then
              '' tagihan ke TAM = kode NIK produksi ada di norangka
              '' -----------------------------------------------------------
36            hasil = True
37            GoTo ReturnHasil
38        Else '
39            tagihanFOP = ""
40            hasil = False
41            GoTo ReturnHasil
42        End If
43       On Error GoTo CekUnitTerlibatFOP_Error

ReturnHasil:
44    CekUnitTerlibatFOP = hasil
45       On Error GoTo 0
46       Exit Function

CekUnitTerlibatFOP_Error:
47    CekUnitTerlibatFOP = False
End Function

Public Sub LoadtipeDisc(nameModel As String)
    Dim rsCus As ADODB.Recordset
    If Log.RPLHistori = 1 Then
        Set rsCus = objO.FetchData("Select DiscNrm, DiscNon FROM [" & dbReplikasiHO & "]..CarClass WITH (NOLOCK) Where Classcode = '" & nameModel & "'")
    Else
        Set rsCus = objO.FetchData("Select DiscNrm, DiscNon FROM CarClass WITH (NOLOCK) Where Classcode = '" & nameModel & "'")
    End If
    ''Set rsCus = objO.FetchData("Select DiscNrm, DiscNon FROM CarClass WITH (NOLOCK) Where Classcode = '" & nameModel & "'")
    If rsCus.RecordCount > 0 Then
        nDiscNRM = Val(rsCus.Fields(0).value)
        nDiscNON = Val(rsCus.Fields(1).value)
        isOnCarClass = True
    Else
        nDiscNRM = 0
        nDiscNON = 0
        isOnCarClass = False
    End If
End Sub

Public Function Dekripsi(KodePassword As String, KataKunci As String) As String

End Function

Public Function KataKunci(Panjang As Long) As String

End Function

Public Function DekripsiGamaLama(KodePassword As String, KataKunci As String) As String
    Dim hasil As String
    hasil = Crypt(KodePassword, KataKunci, False)
    DekripsiGamaLama = hasil
End Function

Public Function KataKunciGamaLama(Panjang As Long) As String
    Dim hasil As String
    If Panjang = 3 Then
        hasil = "M4d462"
    ElseIf Panjang = 5 Then
        hasil = "G4j4H"
    End If
    KataKunciGamaLama = hasil
End Function

Public Function userKon() As String

End Function

Public Function pwdKon() As String

End Function

'Public Function pwdKon() As String
'   pwdKon = "5"
'   pwdKon = pwdKon & "3"
'   pwdKon = pwdKon & "7"
'   pwdKon = pwdKon & "0"
'   pwdKon = pwdKon & "5"
'   pwdKon = pwdKon & "@"
'   pwdKon = pwdKon & "O"
'   pwdKon = pwdKon & "6"
'End Function


Public Function SingleToDouble(kata As String) As String
    SingleToDouble = Replace(Trim(kata), "'", "''")
End Function

'select A.FRONT,A.NO_WO, A.NO_POLISI, B.DATESERAH, B.TIMESERAH,
'(Left(B.TIMESERAH,2)*60)+(Right(B.TIMESERAH,2)) as WaktuPlan
'from JOBDTA A INNER JOIN JOBDTL B ON A.NO_WO = B.NO_WO
'Left JOIN QUEUE05..FINISHSERVICE C ON A.NO_POLISI = C.NOPOL AND A.INVDATE = C.TGLFINISH
'Where B.DateSerah = '2012-12-11 00:00:00.000' AND isNull(C.PANGGIL,0) = 0
'AND (Left(B.TIMESERAH,2)*60)+(Right(B.TIMESERAH,2)) < 970
'Order By  B.TIMESERAH



Public Sub SetIcon(frm As Form, _
                   Optional navigator As Boolean = True)
1     On Error GoTo vb_error
2      With frmRepoBulanan.ImageList1
3       frm.cmdAdd.PictureNormal = .ListImages("cmdadd").Picture
4       frm.cmdAdd.ToolTipText = ""
5       frm.cmdAdd.Tag = "Tambah Record"
6        frm.cmdAdd.Caption = "&Add"
7     End With


8     Exit Sub
vb_error:
9       MsgBox ErrMessage(Erl, err.Number, "Procedure : SetIcon")
10      Resume Next
End Sub

Public Sub RemoveMenus(frm As Form, Optional brestore As Boolean, Optional bmove As Boolean, Optional bsize As Boolean, Optional bminimize As Boolean, Optional bmaximize As Boolean, Optional bseperator As Boolean, Optional bclose As Boolean)
hMenu = GetSystemMenu(frm.hWnd, False)
If bclose Then DeleteMenu hMenu, 6, MF_BYPOSITION
If bseperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
If bmaximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
If bminimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
If bsize Then DeleteMenu hMenu, 2, MF_BYPOSITION
If bmove Then DeleteMenu hMenu, 1, MF_BYPOSITION
If brestore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

Public Function SpyRound(dNumber As Double, Optional doNotRoundUpIfLessThan As Double = 0.5) As Double
    Dim sNumber As String
    Dim arVal() As String
    If doNotRoundUpIfLessThan = 0 Then
        doNotRoundUpIfLessThan = 0.5
    End If
    sNumber = dNumber
    If InStr(1, sNumber, ",") = 0 Then
        SpyRound = dNumber
    Else
        arVal = Split(sNumber, ",")
        sNumber = "0," & arVal(1)
        dNumber = sNumber 'Val(sNumber)
        If dNumber < doNotRoundUpIfLessThan Then
            SpyRound = Val(arVal(0))
        Else
            SpyRound = Val(arVal(0)) + 1
        End If
    End If
End Function

Public Function updateHarian(xDate As Date)
1     On Error GoTo vbError
      Dim rsHlp As ADODB.Recordset

2     If Log.rsA = 1 And Log.tglRSA <= Format(isnow, "dd/mm/yyyy") Then
3         Call objO.InsertData("update profilbengkel..profil set tglUpdtEvd = CONVERT(VARCHAR(10), GETDATE(),102)")
          Dim rsRSA   As ADODB.Recordset
          
4         sSQL = "select DATEDIFF(HOUR,DATE+TIME,GETDATE()) TOTALJAM,NO_WO,NO_POLISI,FMDL,A.MECH,b.[GROUP] FO"
5         sSQL = sSQL & " from JOBDTA a WITH (NOLOCK)"
6         sSQL = sSQL & " LEFT JOIN MEKANIK b WITH (NOLOCK) ON a.MECH = b.MECH"
7         sSQL = sSQL & " where NOINV = '' and FS_LOG_DEL = '' and DATE >= '" & Format(Log.tglRSA, "yyyy-mm-dd") & "' and DATEDIFF(HOUR,DATE+TIME,GETDATE()) >= 24"
8         sSQL = sSQL & " and NO_WO not in (select NO_WO from [" & dbReplikasiHO & "]..NeedSupportTL WHERE KodeDealer = '" & Right(gDb, 2) & "')"
9         sSQL = sSQL & " order by DATE desc"
10        Set rsRSA = objO.FetchData(sSQL)
11        If Not rsRSA.EOF Then
12            rsRSA.MoveFirst
13            Do While Not rsRSA.EOF
14                sSQL = "INSERT INTO [" & dbReplikasiHO & "]..NeedSupportTL (KodeDealer,no_wo,noPol,model,keterangan,tech,FO,tglEntry,userEntry,sumberTL) "
15                sSQL = sSQL & " VALUES ('" & Right(gDb, 2) & "','" & NulSID(rsRSA!no_wo) & "','" & NulSID(rsRSA!NO_POLISI) & "','" & NulSID(rsRSA!Fmdl) & "','','" & NulSID(rsRSA!Mech) & "','" & NulSID(rsRSA!FO) & "',GETDATE(),'SYSTEM','TIME')"
16                Call objO.InsertData(sSQL)
17                rsRSA.MoveNext
18            Loop
19        End If
20    Else
      ' 31 jan 2017, tambahan pembacaan field tgl utk update yg dilakukan 1 hr 1x
21        If DateDiff("d", xDate, Log.tglUpdtEvd) <> 0 Then
              '----------------------------------------------------'
              '- 2015-05-28 | RUDI | KIRIM SUPPORT TL KE HO       -'
              '----------------------------------------------------'
22            Set rsHlp = objO.FetchData("select DATEDIFF(day,a.tglentry,getdate()) hari,a.*,b.no_polisi,b.fmdl from NeedSupportTL A LEFT JOIN JOBDTA B ON a.no_wo = b.no_wo where isnull(a.TRANSFERHO,0) = 0 AND a.tglInfo = '01/01/1900' and isnull(a.userInfo,'') = '' and DATEDIFF(day,a.tglentry,getdate()) >= 2")
23            If Not rsHlp.EOF Then
24                rsHlp.MoveFirst
25                Do While Not rsHlp.EOF
26                    If Log.LogHOConn = 1 Then
27                        If ConHO = True Then
28                            Call objO.InsertDataHO("INSERT INTO NEEDSUPPORTTL (KodeDealer,no_wo,noPol,model,keterangan,tech,FO,tglEntry,userEntry,sumberTL) VALUES ('" & Trim(Right(gDb, 2)) & "','" & Trim(rsHlp!no_wo) & "','" & NulSID(rsHlp!NO_POLISI) & "','" & Left(NulSID(rsHlp!Fmdl), 50) & "','" & Trim(rsHlp!keterangan) & "','" & Trim(rsHlp!tech) & "','','" & Format(Trim(rsHlp!tglentry), "YYYY-MM-DD HH:MM:SS") & "','" & Trim(rsHlp!userentry) & "','" & Trim(rsHlp!userInfo) & "')")
29                            Call objO.InsertData("UPDATE NEEDSUPPORTTL SET TRANSFERHO = '1',TGLTRANSFER = GETDATE() WHERE  ID = '" & Trim(rsHlp!id) & "' ")
30                        End If
31                    End If
32                    rsHlp.MoveNext
33                Loop
34            End If
35            Call objO.InsertData("update profilbengkel..profil set tglUpdtEvd = CONVERT(VARCHAR(10), GETDATE(),102)")
36        End If
37    End If
38    Exit Function

vbError:
39        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in function updateHarian of main")
40        Exit Function
End Function

Public Function s2d(kata As String) As String
   s2d = Replace(Trim(kata), "'", "''")
End Function

Public Function setNpwpSID(value As String) As String
    value = Left(NulSID(value), 16)
    If value <> "" And Val(value) <> "0" Then
        Dim i As Integer
        i = 1
        Dim cek As String
        cek = ""
        Do While i <= 16
            If IsNumeric(Mid(value, i, 1)) = False Then
                cek = cek & "0"
            Else
                cek = cek & Mid(value, i, 1)
            End If
            i = i + 1
        Loop
        value = cek
        Select Case Len(value)
            Case 1
                value = value & "000000000000000"
            Case 2
                value = value & "00000000000000" ' & value
            Case 3
                value = value & "0000000000000" ' & value
            Case 4
                value = value & "000000000000" ' & value
            Case 5
                value = value & "00000000000" ' & value
            Case 6
                value = value & "0000000000" '& value
            Case 7
                value = value & "000000000" ' & value
            Case 8
                value = value & "00000000" '& value
            Case 9
                value = value & "0000000" '& value
            Case 10
                value = value & "000000" ' & value
            Case 11
                value = value & "00000" '& value
            Case 12
                value = value & "0000" ' & value
            Case 13
                value = value & "000"
            Case 14
                value = value & "00"
            Case 15
                value = value & "0"
        End Select
        
        setNpwpSID = NulSID(Replace(Format(value, "000=000=000=0-000=000"), "=", "."))
    Else
        setNpwpSID = "000.000.000.0-000.000"
    End If
                           
End Function

'J.rev20161020 for windows 8.1
Public Sub Sendkeys(Text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), wait

                            Set WshShell = Nothing
End Sub

Public Function t2k(ByVal value As String) As String
  t2k = Replace(value, ".", ",")
End Function

Public Function k2t(ByVal value As String) As String
  k2t = Replace(value, ",", ".")
End Function


Public Sub GetProfileScoring()
On Error GoTo vb_error
    Dim RsBengkel As New ADODB.Recordset
    RsBengkel.CursorLocation = adUseClient
    RsBengkel.Open "SELECT * FROM profilscoring WITH (NOLOCK)", SQLConSID, adOpenStatic, adLockBatchOptimistic, adCmdText
     
    Set RsBengkel.ActiveConnection = Nothing
    
    If RsBengkel.RecordCount > 0 Then
        
        ProfilScoring.SearchingKTP = NulSID(RsBengkel!SearchingKTP)
        ProfilScoring.SearchingHP = NulSID(RsBengkel!SearchingHP)
        ProfilScoring.SearchingNama = NulSID(RsBengkel!SearchingNama)
        ProfilScoring.ScoringNamaKhusus = NulSID(RsBengkel!ScoringNamaKhusus)
        ProfilScoring.nScoringNamaKhusus = CSng(NulSID(RsBengkel!nScoringNamaKhusus))
        ProfilScoring.LockNama1Khusus = CSng(NulSID(RsBengkel!LockNama1Khusus))
        ProfilScoring.ScoringNamaSimultan = NulSID(RsBengkel!ScoringNamaSimultan)
        ProfilScoring.LockNama1 = CSng(NulSID(RsBengkel!LockNama1))
        ProfilScoring.LockNama2 = CSng(NulSID(RsBengkel!LockNama2))
        ProfilScoring.nScoringNama2 = CSng(NulSID(RsBengkel!nScoringNama2))
        ProfilScoring.ScoringNoHP = NulSID(RsBengkel!ScoringNoHP)
        ProfilScoring.ScoringTglLahir = NulSID(RsBengkel!ScoringTglLahir)
        ProfilScoring.ScoringKelurahan = NulSID(RsBengkel!ScoringKelurahan)
        ProfilScoring.ScoringKecamatan = NulSID(RsBengkel!ScoringKecamatan)
         
    End If
    RsBengkel.Close
    'If DBConBengkel.State = adStateOpen Then DBConBengkel.Close
    Set RsBengkel = Nothing
    Set DBConBengkel = Nothing
 Exit Sub
vb_error:
  MsgBox ErrMessage(Erl, err.Number, "GetProfileScoring")
  MsgBox "Error Setting !!!"
  End
End Sub

Public Function getAngkaRomawi(Angka As Integer) As String
    Select Case Angka
        Case 1
            getAngkaRomawi = "I"
        Case 2
            getAngkaRomawi = "II"
        Case 3
            getAngkaRomawi = "III"
        Case 4
            getAngkaRomawi = "IV"
        Case 5
            getAngkaRomawi = "V"
        Case 6
            getAngkaRomawi = "VI"
        Case 7
            getAngkaRomawi = "VII"
        Case 8
            getAngkaRomawi = "VIII"
        Case 9
            getAngkaRomawi = "IX"
        Case 10
            getAngkaRomawi = "X"
        Case 11
            getAngkaRomawi = "XII"
        Case 12
            getAngkaRomawi = "XIII"
    End Select
End Function

'
'Public Function GetFileInformation(ByVal fileFullPath As String, ByRef FileInformation As FILE_INFORMATION, Optional ByVal showMsgBox As Boolean = False) As Boolean
'    Dim lDummy As Long, lsize As Long, rc As Long
'    Dim lVerbufferLen As Long
'    Dim sBuffer() As Byte
'    Dim udtVerBuffer As VS_FIXEDFILEINFO
'    Dim hFile As Integer
'    Dim FileStruct As OF_STRUCT
'    Dim CreationTime As FILE_TIME
'    Dim LastAccessTime As FILE_TIME
'    Dim LastWriteTime As FILE_TIME
'    Dim LocalFileTime As SYSTEM_TIME
'    Dim MessageString As String
'
'    Dim lBufferLen As Long
'    Dim bytebuffer(255) As Byte
'    Dim Lang_Charset_String As String
'    Dim HexNumber As Long
'    Dim i As Integer
'    Dim strTemp As String
'    Dim buffer As String
'    Dim lVerPointer As Long
'    Dim strVersionInfo(7) As String
'
'    On Error GoTo e_HandleFileInformationError
'
'    With FileInformation
'        lsize = GetFileVersionInfoSize(fileFullPath, lDummy)
'        If lsize >= 1 Then
'            ReDim sBuffer(lsize)
'            rc = GetFileVersionInfo(fileFullPath, 0&, lsize, sBuffer(0))
'            rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
'            MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
'        End If
'
'        buffer = String(255, 0)
'
'        '*** Get size ****
'        lBufferLen = GetFileVersionInfoSize(fileFullPath, lDummy)
'        If lBufferLen >= 1 Then
'            ReDim sBuffer(lBufferLen)
'            rc = GetFileVersionInfo(fileFullPath, 0&, lBufferLen, sBuffer(0))
'            If rc <> 0 Then
'                rc = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
'                If rc <> 0 Then
'                    MoveMemory bytebuffer(0), lVerPointer, lBufferLen
'                    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + _
'                    bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
'                    Lang_Charset_String = Hex(HexNumber)
'
'                    Do While Len(Lang_Charset_String) < 8
'                        Lang_Charset_String = "0" & Lang_Charset_String
'                    Loop
'
'                    With FileInformation
'                        .sFileVersion = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "FileVersion", lVerPointer, lBufferLen, sBuffer)
'                        .sInternalName = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "InternalName", lVerPointer, lBufferLen, sBuffer)
'                        .sLegalCopyright = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "LegalCopyright", lVerPointer, lBufferLen, sBuffer)
'                        .sOriginalFileName = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "OriginalFileName", lVerPointer, lBufferLen, sBuffer)
'                    End With
'                End If
'            End If
'        End If
'
'    End With
'
'    OriName = FileInformation.sOriginalFileName
'    FileVersion = FileInformation.sFileVersion
'    GetFileInformation = True
'    Exit Function
'
'e_HandleFileInformationError:
'    GetFileInformation = False
'    Exit Function
'End Function


Public Function GetOriginalName(ByVal fileFullPath As String) As String
    Dim lDummy As Long, lsize As Long, rc As Long
    Dim lVerbufferLen As Long
    Dim sBuffer() As Byte
    Dim udtVerBuffer As VS_FIXEDFILEINFO
    Dim hFile As Integer
    Dim FileStruct As OF_STRUCT
    Dim CreationTime As FILE_TIME
    Dim LastAccessTime As FILE_TIME
    Dim LastWriteTime As FILE_TIME
    Dim LocalFileTime As SYSTEM_TIME
    Dim MessageString As String

    Dim lBufferLen As Long
    Dim bytebuffer(255) As Byte
    Dim Lang_Charset_String As String
    Dim HexNumber As Long
    Dim i As Integer
    Dim strTemp As String
    Dim buffer As String
    Dim lVerPointer As Long
    Dim strVersionInfo(7) As String
    Dim sOriginalFileName As String
    On Error GoTo e_HandleFileInformationError
 
    lsize = GetFileVersionInfoSize(fileFullPath, lDummy)
    If lsize >= 1 Then
        ReDim sBuffer(lsize)
        rc = GetFileVersionInfo(fileFullPath, 0&, lsize, sBuffer(0))
        rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
        MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    End If

    buffer = String(255, 0)

    '*** Get size ****
    lBufferLen = GetFileVersionInfoSize(fileFullPath, lDummy)
    If lBufferLen >= 1 Then
        ReDim sBuffer(lBufferLen)
        rc = GetFileVersionInfo(fileFullPath, 0&, lBufferLen, sBuffer(0))
        If rc <> 0 Then
            rc = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
            If rc <> 0 Then
                MoveMemory bytebuffer(0), lVerPointer, lBufferLen
                HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + _
                bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
                Lang_Charset_String = Hex(HexNumber)

                Do While Len(Lang_Charset_String) < 8
                    Lang_Charset_String = "0" & Lang_Charset_String
                Loop
 
                sOriginalFileName = GetStringValue("\StringFileInfo\" & Lang_Charset_String & "\" & "OriginalFileName", lVerPointer, lBufferLen, sBuffer)
                
            End If
        End If
    End If
  
    GetOriginalName = sOriginalFileName
    Exit Function

e_HandleFileInformationError:
    GetOriginalName = ""
    Exit Function
End Function

Private Function GetStringValue(ByRef searchString As String, ByVal lVerPointer As Long, ByVal lBufferLen As Long, ByRef sBuffer() As Byte) As String
    Dim buffer As String
    Dim strTemp As String
    Dim rc As Long
    
    GetStringValue = ""
    buffer = String(255, 0)
    rc = VerQueryValue(sBuffer(0), searchString, lVerPointer, lBufferLen)
    
    If rc <> 0 Then
    lstrcpy buffer, lVerPointer
    GetStringValue = Mid$(buffer, 1, InStr(buffer, Chr(0)) - 1)
    End If
End Function



''===================================================================================================
'' DAFTAR REVISIAN
''===================================================================================================
'' 20180515REV01'INSPECTION = REVISI INSPECTION DIAMBIL TOP 1 PALING TERAKHIR
'' 20180515REV02'INSPECTION = REVISI JOIN INSPECTION DIAMBIL TOP 1 PALING TERAKHIR
'' 20180515DEV01'INSPECTION = PROSES MENAMPILKAN INSPECTION DAN INPUT GAMBAR
'' 20180518DEV01'INSPECTION = PENAMBAHAN FLAG DARI PROFILBENGKEL TABLE PROFIL
'' 20180524DEV01'INSPECTION = PENAMBAHAN FTP UNTUK UPLOAD FILEIMAGE KE HOST CR7
'' 20180606DEV01'NOTIFSCC = PENAMBAHAN NOTIF SCC PADA MENU ORDER

'' 20180611DEV01'CETAKANBARU = GANTI FORMAT CETAKAN WO DAN INVOICE
'' 20180611DEV02'TAMBAHMATERAI = PENAMBAHAN PROSES MATERAI
'' 20180724DEV01' Tambahan field NoktpKetFP di profil database, jika 1 maka saat insert/update FakturXX..sifpstandar, NoKTP tidak masuk di field keterangan. Jika 0 NoKTP masuk di field keterangan
                ' Tambahan field NoktpNonPerorangan di profil database, jika 1 maka saat insert/update FakturXX..sifpstandar khusus customer non perorangan, NoKTP tidak masuk di field NoKTP. Jika 0 NoKTP masuk di field NoKTP.
'' 20180927DEV  ' DEV | DEVELOP      | SCORING SINGLE ID
'' 20181002REV  ' DEV | MB THANIA    | PENAMBAHAN CETAKAN SIMPLIFY PREDIAG
'' 20181002REV2 ' REQ | MAS ADIT     | MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
'' 20181003REV1 ' REQ | MB THANIA    | MENAMPILKAN TEXTBOX UNTUK PREDIAGNOSE, KODE WARRANTY DAN NAMA WARRANTY
'' 20181008DEV1 ' DEV | MB THANIA    | MENAMBAHKAN PADA MENU MAINTENANCE PERBAIKAN FI
''                                     MENGAMBIL DARI STALL DATA WO DI STALL FINAL DAN FLAG S
''                                     MENGAMBIL TEMPPROGRESS NOORDER DAN FLAGP <> M
''                                     UPDATE TABLE STALL DENGAN NOSTALL SEBELUMNYA DAN FLAG P
'' 20181011REV1 | REV | PK M.WASISTO | PENAMBAHAN SCORING SINGLE ID
''                                     - PINBB PADA NEWCUSTOMER UNTUK PENANDA ADA / TIDAK SUGGEST , JIKA ADA DIISI SCORING
''                                     - TABLE UPLOAD KTP DITAMBAHKAN FIELD UPDATEKE DEFAULT 0,VALID (VARCHAR(10) DEFAULT '' DIISI USER YANG VALIDASI) ,TGLVALID ,USERUPDATE
''                                     - ADD SCORE NO HP BUKAN NOKTP
''                                     - EDIT CUSTOMER
''                                          = JIKA ADA OUTSTANDING WO ATAS USER TERSEBUT MAKA BISA EDIT KTP + UPLOAD KTP
''                                          = JIKA TIDAK ADA OUTSTANDING WO MAKA HANYA BISA EDIT KTP
''                                     - SCORING HANYA UNTUK PERORANGAN
'' 20181018DEV1 | DEV | DEVELOPMENT  | PENAMBAHAN FLAG UNTUK STANDARISASI SUBLET
'' 20181217DEV1 | DEV | DEVELOPMENT  | PENAMBAHAN HARGAJUAL DISOORD UNTUK KEPERLUAN SIMPLIFY PROSES'' 20181218REV1 | REV | MS ADIT      | JAM SELESAI STALL DAN HELPJPCB > 24:00 DIGANTI 23:39 DAN ADJUST > 24:00 DIGANTI 23:59
'' 20181219DEV1 | DEV | DEVELOPMENT  | UR Revisi project Scoring.
'' 20190625DEV1 | DEV | DEVELOPMENT  | FREE SERVICE DIGANTI KE HO BUKAN DEALER PENJUAL
'' 20190701REV1 | REV | MAS ADIT     | PENAMBAHAN FLAG UNTUK NEW SSC LOCK
'' 20190816DEV1 | DEV | DEVELOPMENT  | PENAMBAHAN FLAG UNTUK PROJEK HANSA RENKEI
'' 20191115DEV1 | DEV | DEVELOP      | Repair Support Activity
'' 20200110DEV1 | DEV | DEVELOP      | Lock retur Invoice Materai
'' 20200113DEV1 | DEV | DEVELOPMENT  | Penyamaan Chassing Form (NIS) dengan Lead time (Ohbeya)
'' 20200210DEV  | DEV | DEVELOPMENT  | Multiple marketing program
'' 20200303DEV  | DEV | DEVELOPMENT  | Vin by Vin Retention
'' 20200408DEV  | DEV | DEVELOPMENT  | Sentralisasi Stars
'' 20200416DEV  | DEV | DEVELOPMENT  | FREE PART SERVICE
'' 20200512DEV  | DEV | DEVELOPMENT  | PENAMBAHAN FLAG UNTUK DISABLE DOWNLOAD MENU MRA
'' 20200915DEV  | DEV | DEVELOPMENT  | PROJECT HANSA RENKEI BP
'' 20201022DEV  | DEV | DEVELOPMENT  | Dimana tiap cabang hanya bisa membaca supplier & sublet yang dibuat HO atau cabang itu sendiri,
''                                     Master supplier & master sublet yang dibuat cabang lain tidak bisa dibaca
'' 20201124DEV  | DEV | DEVELOPMENT  | Develop vin 2 vin rev 1 metode pengambilan data berbeda
'' 20201218DEV  | DEV | DEVELOPMENT  | Menambahkan flag materai 2021
'' 20201222DEV  | DEV | DEVELOPMENT  | PROJECT NASMOCO HOME SERVICE
'' 20210115REV  | REV | REVISI       | PENAMBAHAN MENU PENGISIAN NOCLAIM DAN MENUTUP PROSES SIMPAN DI BILLING
'' 20210520DEV  | DEV | DEVELOPMENT  | Menambahkan flag untuk simpan replikasi invoice karena saat replikasi down akan error
'' 20211125DEV  | DEV | DEVELOPMENT  | menambahkan flag untuk dealer non sales untuk membaca data spk ho
'' 20220210DEV  | DEV | DEVELOPMENT  | hide history service km < 10km
'' 20240119DEV  | DEV | DEVELOPMENT  | develop inputan untuk didx project tam
''=======================================================================================================================================================================================================================================
