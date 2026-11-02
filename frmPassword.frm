VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "Gif89.dll"
Begin VB.Form frmPassword 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6720
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPassword.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PictLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   4215
      TabIndex        =   5
      Top             =   2120
      Visible         =   0   'False
      Width           =   4215
      Begin GIF89LibCtl.Gif89a gLoad 
         Height          =   5655
         Left            =   -1970
         OleObjectBlob   =   "frmPassword.frx":57FE
         TabIndex        =   6
         Top             =   -2640
         Width           =   4455
      End
   End
   Begin VB.Timer TimerLoading 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6240
      Top             =   2160
   End
   Begin VB.TextBox txtEntry 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   400
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2520
      PasswordChar    =   "O"
      TabIndex        =   2
      Top             =   2140
      Width           =   4095
   End
   Begin VB.TextBox txtEntry 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1700
      Width           =   4095
   End
   Begin VB.TextBox txtEntry 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1240
      Width           =   975
   End
   Begin AIFCmp1.asxPowerButton cmdLogin 
      Height          =   600
      Left            =   0
      TabIndex        =   3
      Top             =   2720
      Width           =   6740
      _ExtentX        =   11880
      _ExtentY        =   1058
      BackColor       =   14482683
      BorderStyle     =   0
      Picture         =   "frmPassword.frx":5840
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Motion  Capture  Data  Provided  By  Thania  /  Acep  IT   @  Nasmoco  2011"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3390
      Width           =   6735
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fPassword As String
Public fLevel As String
Dim VKd_Sama As String
Private nLogin As Integer

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub cmdLogin_Click()
    If UCase(fPassword) <> UCase(txtEntry(2).text) Then
        MsgBox "INCORECT USER / PASSWORD !!!", vbInformation + vbOKOnly, "Login NIS Service"
        nLogin = nLogin + 1
        If nLogin = 3 Then End
        Exit Sub
    Else
        TimerLoading.Enabled = True
        PictLoad.Visible = True
        gLoad.FileName = Gif
    End If
End Sub

Private Sub Go_Login()
    If (UCase(fPassword) = UCase(fPassword) And txtEntry(1) <> "") _
    Or fPassword = Left(gTanggal, 2) + Right(gTanggal, 4) Then
        If fPassword = Left(gTanggal, 2) + Right(gTanggal, 4) Then
            gLevel_Ptg = 1
        Else
            gLevel_Ptg = 2
        End If
        gNama_SA = txtEntry(1)
        gSA = txtEntry(0)
        gKd_Sama = VKd_Sama
        gLoket = GetSetting(appname:="SERVICE", Section:="Loket", Key:="Top", Default:="0")
        Set oRs = objO.FetchData("Select * from T_PROFIL")
        If Not oRs.BOF Then
            gPrinter = oRs("FS_PRINTER")
            gRateCOM = oRs("FN_RATE_COM")
            gRatePAS = oRs("FN_RATE_PAS")
            gRateIRC = oRs("FN_RATE_IRC")
            gPerush = oRs("FS_NM_PERUSH")
            gAlamat = oRs("FS_ALM_PERUSH")
            gKotaCbg = oRs("FS_KOTA")
            gCabang = oRs("FS_CABANG")
            gSecmanager = oRs("FS_SEC_MGR")
            gADH = oRs("FS_KA_ADH")
            gNpwp = oRs("FS_NPWP")
            gPkp = oRs("FS_PKP")
            gTelp = oRs("FS_TELP")
            gBank = oRs("FS_BANK")
            gAcc = oRs("FS_ACC")
            gRefreshInfo = oRs("FN_REFRESH_INFO")
            gUkuranOrder = oRs("FN_ORDER")
            gUkuranInvoice = oRs("FN_INVOICE")
            gUkuranRetur = oRs("FN_RETUR")
            gPreviewCetak = oRs("FN_PREVIEW")
            gFPart = oRs("FN_PART")
            gIERC = oRs("FN_IERC")
            gCASH = oRs("FN_CASH")
            gReturn = oRs("FS_KALI_RET_RATE")
            gKodeFaktur = oRs("FS_KD_FAKTUR")
            gGenerasi = oRs("FS_GENERASI")
            gFiskal = oRs("FS_FISKAL")
            gConnecAAS = oRs("FN_AAS")
            gGudang = oRs("FS_GUDANG")
            gDiscMax = CDbl(oRs("DISCMAX"))
            gUpSublet = CDbl(oRs("UPSUBLET"))
            gNoShow = CDbl(oRs("NOSHOW"))
            gFBackup = Trim(oRs("FOLDERBACKUP"))
            gLimitFUDealer = CDbl(oRs("LIMITDEALER"))
            hTipePenghitunganDealer = IIf(Trim(oRs("ESTIMASIDEALER")) <> "", Trim(oRs("ESTIMASIDEALER")), "NextEstimasi")
            hLimitMou = CDbl(oRs("LIMITMOU"))
            gQueue = IIf(oRs("FN_QUEUE") = "1", True, False)
            
            gJurnalPart = False
            Dim rsSearch As ADODB.Recordset
            Set rsSearch = objO.FetchData("SELECT top 1 * From syscolumns where id in (select id from sysobjects where id = object_id(N'[dbo].[T_PROFIL]')  ) AND [NAME] = 'INCLUDEPPN'")
            If Not rsSearch.EOF Then
                gJurnalPart = True
            End If
            If gUstat = 0 Then
                gFont = CDbl(oRs("NOFONT"))
                gHeight = CDbl(oRs("NOHEIGHT"))
            End If
            
            'STATUS KONEKSI HO
            gStatKoneksi = IIf(Trim(oRs("KONEKSIHO")) = "OK", True, False)
            fgCabang = True
            gHo = False
            If gStatKoneksi Then
                If Not Get_ConnectionHO_New Then
                    MsgBox "KONEKSI KE HO GAGAL !!!", vbCritical + vbOKOnly, "KONEKSI HO"
                    gStatKoneksi = False
                End If
                If gUstat = 0 Then
                    If oConnHO.State <> adStateClosed Then oConnHO.Close
                End If
            End If
            If gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
                gHo = True
                fgCabang = False
            End If
            
            ' DATA HO
            gNmHO = oRs("FN_HO")
            gAlmtHO = oRs("FN_HOADRES")
            gKotaHO = oRs("FN_HOKOTA")
            gTelpHO = oRs("FN_HOTELP")
            gNpwpHO = oRs("FN_HONPWP")
            gKdPlgHo = oRs("FN_HOKDPLG")
            
            gManPart = oRs("FN_MAN_PART")
            gTglManAw = TglDDMMYYYY(oRs("FD_MAN_AW"))
            gTglManAk = TglDDMMYYYY(oRs("FD_MAN_AK"))
            gKdCbg = oRs("FS_KD_DEALER")
            gNmCbg = oRs("FS_NM_DEALER")
            gTglKukuh = TglDDMMMMYYYY(oRs("TGLPKP"))
        Else
            gPrinter = "LPT1"
        End If
        Call saveLogInUser(gSA, ComptName)
        oRs.Close
        Unload Me
        If Trim(gTipeUser) = "BILLING" Then
            Layar.pic4text.Visible = True
        Else
            Layar.pic4text.Visible = False
        End If
        Layar.Caption = Trim(gSA) & " ..|" & Layar.Caption
        Layar.Show
    Else
        End
    End If
End Sub

Private Sub Form_Load()
    App.Title = "SERVICE"
    nLogin = 0
End Sub

Private Sub TimerLoading_Timer()
    TimerLoading.Enabled = False
    Call Go_Login
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If txtEntry(Index).BackColor = &HFFC0C0 Then
        DiTab = True
    Else
        If gSendTab Then
            txtEntry(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
            gSendTab = False
        End If
    End If
    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    txtEntry(Index) = LTrim(txtEntry(Index))
    Select Case Index
    Case 0
        Set oRs = objO.FetchData("SELECT * from t_petugas where fs_kd_petugas='" & txtEntry(0).text & "'")
        If Not oRs.BOF Then
            txtEntry(0) = RTrim(oRs("fs_kd_petugas"))
            txtEntry(1) = RTrim(oRs("fs_nm_petugas"))
            fPassword = RTrim(Mid(oRs("fs_kd_pass"), 3, 10))
            VKd_Sama = RTrim(oRs("fs_kd_group"))
            gDiv = Trim(oRs("DIVISI"))
            gPrintCetak = searchPrinter(Trim(oRs("TPPRINT")))
            fLevel = RTrim(Mid(oRs("fs_kd_pass"), 1, 2))
            gTipeUser = Trim(oRs("TIPE"))
        Else
            txtEntry(1) = ""
            fPassword = ""
            VKd_Sama = ""
            gPrintCetak = ""
            gDiv = ""
        End If
    End Select
End Sub

