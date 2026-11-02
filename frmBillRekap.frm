VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmBillRekap 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekap Billing"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   12195
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10695
      Left            =   0
      Picture         =   "frmBillRekap.frx":0000
      ScaleHeight     =   10695
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   -120
      Width           =   495
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   600
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   7575
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.ComboBox cboRekap 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   2145
      TabIndex        =   0
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   661
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   2
      Left            =   2145
      TabIndex        =   4
      Top             =   1440
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   6870
      TabIndex        =   11
      Top             =   1035
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   3
      Top             =   1035
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin ucXPButton.XPButton TblKeluar 
      Height          =   495
      Left            =   3840
      TabIndex        =   14
      Top             =   2040
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&x"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBillRekap.frx":23C3
      PICN            =   "frmBillRekap.frx":23DF
      PICH            =   "frmBillRekap.frx":28B3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton TblSimpan 
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   2040
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&s"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBillRekap.frx":2DFE
      PICN            =   "frmBillRekap.frx":2E1A
      PICH            =   "frmBillRekap.frx":3271
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmBillRekap.frx":375B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   840
      X2              =   6960
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   840
      X2              =   6960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   1485
      Width           =   1110
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipe Rekap"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   1080
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   225
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   600
      Width           =   570
   End
End
Attribute VB_Name = "frmBillRekap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 28 Mei 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================
Dim rpt As Object
Dim rpt2 As Object

Private Sub hapusForm()
    
    text(0) = TglDDMMYYYY(Format(isnow, "DD/MM/YYYY"))
    text(1) = TglDDMMYYYY(Format(isnow, "DD/MM/YYYY"))
    text(2) = ""
    text(3) = ""
    text(3).Visible = False
    MaskEdBox1.Visible = False
    
    ' ISI COMBO
    cboRekap.Clear
    cboRekap.AddItem ("PDS")
    cboRekap.AddItem ("SBI")
    cboRekap.AddItem ("RETURN JOB")
    cboRekap.AddItem ("MOU")
    cboRekap.AddItem ("PROGRAM")
    cboRekap.AddItem ("WARRANTY")
End Sub

Private Sub Cetak(idx As Integer, SQL As String, cust As String, almt As String, kt As String, npwp As String, CP As String, judul As String)
    Set rpt = New RptBillRekap
    Set rpt2 = New RptBillRekapDealer
    Dim rsRepo As ADODB.Recordset
    
    If idx = 1 Then
        Set rsRepo = objO.FetchData(SQL)
        rpt.database.SetDataSource rsRepo
        rpt.txtJudul.SetText judul
        rpt.txtTujuan.SetText cust
        rpt.txtAlmt.SetText gInvAlmt 'almt
        rpt.txtKt.SetText gInvKota 'kt
        rpt.txtNpwp.SetText gInvNPWP  'npwp
        rpt.txtKota.SetText gKotaCbg & "  , "
        rpt.txtContact.SetText CP
        rpt.txtKabeng.SetText gSecmanager
        rpt.txtTanggal.SetText "Tanggal : " & Format(text(0), "DD-MM-YYYY") & " s/d " & Format(text(1), "DD-MM-YYYY")
    ElseIf idx = 2 Then
        Set rsRepo = objO.FetchData(SQL)
        rpt2.database.SetDataSource rsRepo
'        rpt2.txtTujuan.SetText cust
'        rpt2.txtAlmt.SetText almt
'        rpt2.txtKt.SetText kt
'        rpt2.txtNpwp.SetText npwp
        rpt2.txtContact.SetText CP
        rpt2.txtKota.SetText gKotaCbg & "  , "
        rpt2.txtKabeng.SetText gSecmanager
        rpt2.txtTanggal.SetText "Tanggal : " & Format(text(0), "DD-MM-YYYY") & " s/d " & Format(text(1), "DD-MM-YYYY")
    End If
    
    With CRViewer1
        .ReportSource = IIf(idx = 1, rpt, rpt2)
        .ViewReport
        .Zoom 100
        
        cmdSettingPrinter.Top = cmdTutup.Top
        cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
        cmdSettingPrinter.Visible = True
        
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub cboRekap_KeyUp(KeyCode As Integer, Shift As Integer)
    text(3).Visible = False
    MaskEdBox1.Visible = False
    If cboRekap = "SBI" Then
        text(3).Visible = True
        MaskEdBox1.Visible = True
        text(3).SelLength = Len(text(3))
        text(3).SetFocus
    End If
End Sub

Private Sub cmdSettingPrinter_Click()
    If idx = 1 Then
        rpt.PrinterSetup Me.hWnd
    Else
        rpt2.PrinterSetup Me.hWnd
    End If
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblSimpan_Click()
Dim rsHlp As ADODB.Recordset
Dim idn As Integer
Dim gJudul As String
Dim Txt As String
    
    If cboRekap <> "" Then
        Select Case cboRekap
            Case "PDS"
                ' TUJUAN REKAPAN TAGIHAN KE BENGKEL
                gInvCust = gPerush
                gInvAlmt = gAlamat
                gInvKota = gKota
                gInvNPWP = "NPWP : " & gNpwp
                gJudul = "PRE DELIVERY SERVICE [PDS]"
                
                ' QUERY REKAPAN WHERE PDS = 1
                sSQL = "SELECT B.FS_KD_DEALER, B.FS_NM_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, A.FS_KD_PLG, A.NO_CHASSIS, A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, NO_WO, [DATE] AS TGLWO, NOINV, INVDATE AS TGLINV, A.FMDL, " _
                     & "REP1+' '+REP2+' '+REP3+' '+REP4+' '+REP5+' '+REP6+' '+REP7+' '+REP8 AS PEKERJAAN, LABOR AS JASA, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, SBLT + MET AS SUBLET, MTRL, BODY, TOTAL, FS_RC , TPSVC, FN_RETURN, MOU, PDS " _
                     & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN T_DEALER C WITH(NOLOCK) ON B.FS_KD_DEALER = C.FS_KD_DEALER AND C.FS_LOG_DEL = '' WHERE A.FS_LOG_DEL = '' AND PDS = '1'" _
                     & "AND INVDATE BETWEEN '" & Format(text(0), "MM/DD/YYYY") & "' AND '" & Format(text(1), "MM/DD/YYYY") & "' ORDER BY A.NO_CHASSIS"
                idn = 1
            Case "SBI"
                ' TUJUAN REKAPAN TAGIHAN KE DEALER LAIN
                ''GROUP BY DEALER, PISAH HALAMAN
                gJudul = "SERVICE BERKALA INTERNAL [SBI]"
                Txt = ""
                If Trim(text(3)) <> "" Then
                    Txt = "AND B.FS_KD_DEALER = '" & Trim(text(3)) & "'"
                End If
                
                ' QUERY REKAPAN WHERE FS_RC = 1
                sSQL = "SELECT B.FS_KD_DEALER, B.FS_NM_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, A.FS_KD_PLG, A.NO_CHASSIS, A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, NO_WO, [DATE] AS TGLWO, NOINV, INVDATE AS TGLINV, A.FMDL, " _
                     & "REP1+' '+REP2+' '+REP3+' '+REP4+' '+REP5+' '+REP6+' '+REP7+' '+REP8 AS PEKERJAAN, LABOR AS JASA, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, SBLT + MET AS SUBLET, MTRL, BODY, TOTAL, FS_RC , TPSVC, FN_RETURN, MOU, PDS " _
                     & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN T_DEALER C WITH(NOLOCK) ON B.FS_KD_DEALER = C.FS_KD_DEALER AND C.FS_LOG_DEL = '' WHERE A.FS_LOG_DEL = '' AND RIGHT(FS_RC,3) = '001'" _
                     & "AND INVDATE BETWEEN '" & Format(text(0), "MM/DD/YYYY") & "' AND '" & Format(text(1), "MM/DD/YYYY") & "' " & Txt & " ORDER BY A.NO_CHASSIS"
                idn = 2
            Case "RETURN JOB"
                gJudul = "RETURN JOB [RTJ]"
                sSQL = "SELECT B.FS_KD_DEALER, B.FS_NM_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, A.FS_KD_PLG, A.NO_CHASSIS, A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, NO_WO, [DATE] AS TGLWO, NOINV, INVDATE AS TGLINV, A.FMDL, " _
                     & "REP1+' '+REP2+' '+REP3+' '+REP4+' '+REP5+' '+REP6+' '+REP7+' '+REP8 AS PEKERJAAN, '0' AS JASA, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, SBLT + MET AS SUBLET, MTRL, BODY, TOTAL, FS_RC , TPSVC, FN_RETURN, MOU, PDS " _
                     & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN T_DEALER C WITH(NOLOCK) ON B.FS_KD_DEALER = C.FS_KD_DEALER AND C.FS_LOG_DEL = '' WHERE A.FS_LOG_DEL = '' AND FN_RETURN = '1'" _
                     & "AND INVDATE BETWEEN '" & Format(text(0), "MM/DD/YYYY") & "' AND '" & Format(text(1), "MM/DD/YYYY") & "' ORDER BY A.NO_CHASSIS"
                idn = 1
            Case "MOU"
                ' TUJUAN REKAPAN TAGIHAN KE BENGKEL
                gInvCust = gPerush
                gInvAlmt = gAlmt
                gInvKota = gKota
                gInvNPWP = "NPWP : " & gNpwp
                gJudul = "MOU"
                
                ' QUERY REKAPAN WHERE MOU <> ""
                sSQL = "SELECT B.FS_KD_DEALER, B.FS_NM_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, A.FS_KD_PLG, A.NO_CHASSIS, A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, NO_WO, [DATE] AS TGLWO, NOINV, INVDATE AS TGLINV, A.FMDL, " _
                     & "REP1+' '+REP2+' '+REP3+' '+REP4+' '+REP5+' '+REP6+' '+REP7+' '+REP8 AS PEKERJAAN, LABOR AS JASA, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, SBLT + MET AS SUBLET, MTRL, BODY, TOTAL, FS_RC , TPSVC, FN_RETURN, MOU, PDS " _
                     & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN T_DEALER C WITH(NOLOCK) ON B.FS_KD_DEALER = C.FS_KD_DEALER AND C.FS_LOG_DEL = '' WHERE A.FS_LOG_DEL = '' AND MOU <> ''" _
                     & "AND INVDATE BETWEEN '" & Format(text(0), "MM/DD/YYYY") & "' AND '" & Format(text(1), "MM/DD/YYYY") & "' ORDER BY A.NO_CHASSIS"
                idn = 1
            Case "PROGRAM"
                ' TUJUAN REKAPAN TAGIHAN KE BENGKEL
                gInvCust = gPerush
                gInvAlmt = gAlmt
                gInvKota = gKota
                gInvNPWP = "NPWP : " & gNpwp
                gJudul = "PROGRAM KHUSUS"
                
                ' QUERY REKAPAN WHERE TPSVC = 'P'
                sSQL = "SELECT B.FS_KD_DEALER, B.FS_NM_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, A.FS_KD_PLG, A.NO_CHASSIS, A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, NO_WO, [DATE] AS TGLWO, NOINV, INVDATE AS TGLINV, A.FMDL, " _
                     & "REP1+' '+REP2+' '+REP3+' '+REP4+' '+REP5+' '+REP6+' '+REP7+' '+REP8 AS PEKERJAAN, LABOR AS JASA, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, SBLT + MET AS SUBLET, MTRL, BODY, TOTAL, FS_RC , TPSVC, FN_RETURN, MOU, PDS " _
                     & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN T_DEALER C WITH(NOLOCK) ON B.FS_KD_DEALER = C.FS_KD_DEALER AND C.FS_LOG_DEL = '' WHERE A.FS_LOG_DEL = '' AND TPSVC = 'P'" _
                     & "AND INVDATE BETWEEN '" & Format(text(0), "MM/DD/YYYY") & "' AND '" & Format(text(1), "MM/DD/YYYY") & "' ORDER BY A.NO_CHASSIS"
                idn = 1
            Case "WARRANTY"
                ' TUJUAN REKAPAN TAGIHAN KE HO
                Set rsHlp = objO.FetchData("SELECT FN_HO, FN_HOADRES, FN_HOKOTA, FN_HOTELP, FN_HONPWP FROM T_PROFIL WITH(NOLOCK) ")
                If Not rsHlp.EOF Then
                    gInvCust = rsHlp("FN_HO")
                    gInvAlmt = rsHlp("FN_HOADRES")
                    gInvKota = rsHlp("FN_HOKOTA")
                    gInvNPWP = "NPWP : " & rsHlp("FN_HONPWP")
                    gJudul = "WARRANTY"
                End If
                
                ' QUERY REKAPAN WHERE TPSVC = 'W'
                sSQL = "SELECT B.FS_KD_DEALER, B.FS_NM_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, A.FS_KD_PLG, A.NO_CHASSIS, A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, NO_WO, [DATE] AS TGLWO, NOINV, INVDATE AS TGLINV, A.FMDL, " _
                     & "REP1+' '+REP2+' '+REP3+' '+REP4+' '+REP5+' '+REP6+' '+REP7+' '+REP8 AS PEKERJAAN, LABOR AS JASA, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, SBLT + MET AS SUBLET, MTRL, BODY, TOTAL, FS_RC , TPSVC, FN_RETURN, MOU, PDS " _
                     & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN T_DEALER C WITH(NOLOCK) ON B.FS_KD_DEALER = C.FS_KD_DEALER AND C.FS_LOG_DEL = '' WHERE A.FS_LOG_DEL = '' AND TPSVC = 'W'" _
                     & "AND INVDATE BETWEEN '" & Format(text(0), "MM/DD/YYYY") & "' AND '" & Format(text(1), "MM/DD/YYYY") & "' ORDER BY A.NO_CHASSIS"
                idn = 1
        End Select
        Call Cetak(idn, sSQL, gInvCust, gInvAlmt, gInvKota, gInvNPWP, Trim(text(2)), gJudul)
    Else
        MsgBox "Isi Pilihan Rekap !"
        cboRekap.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
            text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
        
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 3
                gMainKeyCari = text(Index)
                gSearchOption = "DEALERINV"
                gBantuan = Trim(text(0))
                gBantuan2 = Trim(text(1))
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    Select Case Index
        Case 0, 1
            KeyAscii = CekNumerik(KeyAscii)
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 0, 1
            If Not CekTgl(text(Index)) Then
                If text(Index) <> ">>->>->>>>" Then
                    MsgBox "Format tanggal salah"
                End If
                text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
                Exit Sub
            End If
        Case 3
            If text(Index) <> "" Then
                sSQL = "SELECT B.FS_KD_DEALER, B.FS_NM_DEALER FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS " _
                     & "WHERE A.FS_LOG_DEL = '' AND RIGHT(FS_RC,3) = '001' AND INVDATE BETWEEN '" & Format(text(0), "MM/DD/YYYY") & "' AND '" & Format(text(1), "MM/DD/YYYY") & "'"
                'rsSearch.Open sSQL, oConn, adOpenDynamic, adLockBatchOptimistic
                Set rsSearch = objO.FetchData(sSQL)
                If rsSearch.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "DEALERINV"
                    gBantuan = Trim(text(0))
                    gBantuan2 = Trim(text(1))
                    frmCari.Show (vbModal)
                End If
            End If
    End Select
End Sub
