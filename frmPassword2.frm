VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
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
   Picture         =   "frmPassword2.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6000
      Top             =   840
   End
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
      Begin VB.PictureBox gLoad 
         Height          =   5655
         Left            =   -1970
         ScaleHeight     =   5595
         ScaleWidth      =   4395
         TabIndex        =   6
         Top             =   -2640
         Width           =   4455
      End
   End
   Begin VB.Timer TimerLoading 
      Enabled         =   0   'False
      Interval        =   1
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
      BackColor       =   &H00C0FFFF&
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
      Top             =   2760
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1058
      BackColor       =   14482683
      BorderStyle     =   0
      Picture         =   "frmPassword2.frx":5A97
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
      Caption         =   "COPYRIGHT @ IT PT. NEW RATNA MOTOR 2010"
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
Dim nMov As Integer

Private Sub Timer1_Timer()
    If 0 <= nMov And nMov <= 5 Then
        nMov = nMov + 1
        Me.Left = Me.Left + 50
    ElseIf 6 <= nMov And nMov <= 15 Then
        nMov = nMov + 1
        Me.Left = Me.Left - 50
    ElseIf 16 <= nMov And nMov <= 20 Then
        nMov = nMov + 1
        Me.Left = Me.Left + 50
    End If
    If nMov = 21 Then
        nMov = 0
        Timer1.Enabled = False
        MsgBox "INCORECT USER / PASSWORD !!!", vbInformation + vbOKOnly, "Login NIS Service"
        txtEntry(2).SetFocus
    End If
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub cmdLogin_Click()
    Timer1.Enabled = True
    If NulSID(txtEntry(0)) = "" Then
        Exit Sub
    End If
    If UCase(fPassword) <> UCase(txtEntry(2).Text) Then
        nLogin = nLogin + 1
        txtEntry(2).Text = ""
        If nLogin = 3 Then End
        Exit Sub
    Else
        If Len(txtEntry(2).Text) < 6 Then
            gNama_SA = txtEntry(1)
            gSA = txtEntry(0)
            If NulSID(gSA) = "" Then
                End

            End If
            Unload Me
            MsgBox "Password Anda kurang dari 6 karakter, silahkan ganti password anda !", vbOKOnly, "Informasi"
            frmSetupPasswordLogin.show
        Else
            TimerLoading.Enabled = True
        End If
        
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
    
        If gTipeUser <> "HO" Then
            Dim rsDisc As ADODB.Recordset
            Set rsDisc = objSt.FetchDiscNRM
            nDiscNRM = rsDisc.Fields(0).value
            nDiscNON = rsDisc.Fields(1).value
        End If
        
        Set oRs = objO.FetchData("Select * from T_PROFIL WITH (NOLOCK)")
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
            
            gMaxHariBooking = CInt(oRs("maxHariBooking"))
            
            Log.kdDealerDidx = Trim(oRs("kdDealerDidx"))
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
            fgCabang = True
            gHo = False
            If gStatKoneksi Then
                If Log.LogHOConn = 1 Then
                    If ConHO = False And Log.RPLHistori = 0 Then
                        MsgBox "KONEKSI KE HO GAGAL !!!", vbCritical + vbOKOnly, "KONEKSI HO"
                        gStatKoneksi = False
                    Else
                        gStatKoneksi = True
                    End If
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
            
            
            ' DATA TAM
            gNmTAM = oRs("FN_TAM")
            gAlmtTAM = oRs("FN_TAMADRES")
            gKotaTAM = oRs("FN_TAMKOTA")
            gTelpTAM = oRs("FN_TAMTELP")
            gNpwpTAM = oRs("FN_TAMNPWP")
            gKdPlgTAM = oRs("FN_TAMKDPLG")
            
            gManPart = oRs("FN_MAN_PART")
            gTglManAw = TglDDMMYYYY(oRs("FD_MAN_AW"))
            gTglManAk = TglDDMMYYYY(oRs("FD_MAN_AK"))
            gKdCbg = oRs("FS_KD_DEALER")
            gNmCbg = oRs("FS_NM_DEALER")
            gTglKukuh = TglDDMMMMYYYY(oRs("TGLPKP"))
        Else
            gPrinter = "LPT1"
        End If
        
        'lakukan pengecekan password update terakhir
        If Log.expPwd > 0 Then
            If DateDiff("d", gTglUpdate, isnow) > Log.expPwd Then
                oRs.Close
                Unload Me
                
                MsgBox "Password Anda telah usang, silahkan ganti password anda !", vbOKOnly, "Informasi"
                frmSetupPasswordLogin.show
            Else
                Call saveLogInUser(gSA, ComptName)
                oRs.Close
                Unload Me
                If Trim(gTipeUser) = "BILLING" Then
                    Layar.pic4text.Visible = True
                Else
                    Layar.pic4text.Visible = False
                End If
                Layar.Caption = Trim(gSA) & " ..|" & Layar.Caption
                Layar.show
            End If
        Else
            Call saveLogInUser(gSA, ComptName)
            oRs.Close
            Unload Me
            If Trim(gTipeUser) = "BILLING" Then
                Layar.pic4text.Visible = True
            Else
                Layar.pic4text.Visible = False
            End If
            Layar.Caption = Trim(gSA) & " ..|" & Layar.Caption
            Layar.show
        End If
    Else
        End
    End If
End Sub

Private Sub Form_Load()
    'App.Title = "SERVICE"
    nLogin = 0
    nMov = 0
End Sub


Private Sub TimerLoading_Timer()
    TimerLoading.Enabled = False
    Call Go_Login
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If txtEntry(Index).BackColor = &HC0FFFF Then
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
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
 
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
    If KeyCode = 27 Then End
    
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    txtEntry(Index) = LTrim(txtEntry(Index))
    Select Case Index
    Case 0
        If NulSID(txtEntry(0)) = "" Then
            txtEntry(1) = ""
            fPassword = ""
            VKd_Sama = ""
            gPrintCetak = ""
            gDiv = ""
            Exit Sub
        End If
        Set oRs = objO.FetchData("SELECT * from t_petugas WITH (NOLOCK) where fs_kd_petugas='" & txtEntry(0).Text & "'")
        If Not oRs.BOF Then
            txtEntry(0) = RTrim(oRs("fs_kd_petugas"))
            txtEntry(1) = RTrim(oRs("fs_nm_petugas"))
            fPassword = RTrim(Mid(oRs("fs_kd_pass"), 3, 100))
            VKd_Sama = RTrim(oRs("fs_kd_group"))
            gDiv = Trim(oRs("DIVISI"))
            gPrintCetak = searchPrinter(Trim(oRs("TPPRINT")))
            fLevel = RTrim(Mid(oRs("fs_kd_pass"), 1, 2))
            gTipeUser = Trim(oRs("TIPE"))
            gTglUpdate = Format(oRs("TGLUPDATE"), "DD/MM/YYYY HH:NN:SS")
            If gUstat = 0 Then
                gNIK = NulSID(oRs!nik)
            End If
        Else
            txtEntry(1) = ""
            fPassword = ""
            VKd_Sama = ""
            gPrintCetak = ""
            gDiv = ""
        End If
    End Select
End Sub

