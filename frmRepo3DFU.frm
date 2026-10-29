VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepo3DFU 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H000000FF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4755
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1455
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4755
      Width           =   1440
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4755
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCetak 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8385
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4755
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5265
      TabIndex        =   15
      Top             =   4785
      Visible         =   0   'False
      Width           =   3015
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   4155
      TabIndex        =   14
      Top             =   4770
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      BTYPE           =   8
      TX              =   "Printer :"
      ENAB            =   0   'False
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepo3DFU.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   600
      TabIndex        =   5
      Top             =   4440
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1440
      TabIndex        =   11
      Top             =   1680
      Width           =   3855
      Begin VB.OptionButton OpRepo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Voice Of Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton OpRepo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fix It Right"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton TblSimpan 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1890
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
   End
   Begin MSComCtl2.DTPicker tanggal1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   681836545
      CurrentDate     =   39859
   End
   Begin MSComCtl2.DTPicker tanggal2 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   681836545
      CurrentDate     =   39859
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   270
      Left            =   2235
      TabIndex        =   8
      Top             =   1320
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
      Index           =   1
      Left            =   2355
      TabIndex        =   9
      Top             =   1320
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
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
   Begin VB.Image asxPowerBanner1 
      Height          =   9135
      Left            =   0
      Picture         =   "frmRepo3DFU.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SA"
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
      Left            =   975
      TabIndex        =   10
      Top             =   1365
      Width           =   195
   End
   Begin VB.Label Label10 
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
      Left            =   600
      TabIndex        =   7
      Top             =   870
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S/d"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   870
      Width           =   240
   End
End
Attribute VB_Name = "frmRepo3DFU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As Object

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

Private Sub Form_Load()
    tanggal1 = Format(Month(isnow) & "/01/" & year(isnow), "MM/DD/YYYY")
    tanggal2 = GetTglAkhir(Month(isnow), year(isnow), 0)
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblSimpan_Click()

    If OpRepo(0) Then
        Call repoFIR
    ElseIf OpRepo(1) Then
        Call repoVoiceOfCustomer
    End If
End Sub

Private Sub repoVoiceOfCustomer()
Dim oRs As Object
'Dim Report As New repoVoiceOfCustomer
Set Report = New repoVoiceOfCustomer

Dim Txt As String
Dim txtTemp As String
Screen.MousePointer = vbHourglass

    Txt = ""
    txtTemp = ""
    
    ''' REPO VOC
    If Trim(text(0)) <> "" Then
        Txt = "AND B.FRONT = '" & Trim(text(0)) & "'"
        txtTemp = "AND SA = '" & Trim(text(0)) & "'"
    End If
         
    sSQL = "SELECT DISTINCT B.DATE, B.TIME, A.NO_WO, A.NOINV, B.CUSTOMER, B.CUST_ADRES, B.NO_CHASSIS, B.NO_POLISI, " _
         & "B.REQ1,B.REQ2,B.REQ3,B.REQ4,B.REQ5,D.CMPL1,D.CMPL2,D.CMPL3,D.CMPL4,D.CMPL5,B.REP1,B.REP2,B.REP3,B.REP4,B.REP5,B.REP6,B.REP7,B.REP8, " _
         & "B.FRONT , B.MECH, E.invDATE AS TGLINV, E.invTIME AS JAMINV, KELUHAN1, KELUHAN2, KELUHAN3, KODEITEM, KETERANGAN, QTY, F.TELPON, F.GSM, F.CDMA " _
         & "FROM DTLFU A WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS " _
         & "INNER JOIN CUSTDAT F WITH(NOLOCK) ON A.NO_CHASSIS = F.NO_CHASSIS " _
         & "INNER JOIN JOBDTL D WITH(NOLOCK) ON A.NO_WO = D.NO_WO AND A.NO_CHASSIS = D.NO_CHASSIS INNER JOIN T_INVOICE E WITH(NOLOCK) ON A.NO_WO = E.NO_WO AND A.NOINV = LEFT(E.NOINV,8) " _
         & "LEFT JOIN T_PART C WITH(NOLOCK) ON A.NO_WO = C.NO_WO AND A.NOINV = LEFT(C.NOINV,8) " _
         & "WHERE SAMBUNG = 'Y' AND B.DELIV <> 'BTL' AND E.FS_LOG_DEL = '' AND (KUALITAS <> 'P' OR TEPAT <> 'P' OR LAMA <> 'P') AND " _
         & "B.DATE BETWEEN '" & Format(tanggal1, "MM/DD/YYYY") & "' AND '" & Format(tanggal2, "MM/DD/YYYY") & "' " & Txt & " ORDER BY B.DATE, A.NO_WO"
    Set oRs = objO.FetchData(sSQL)
                
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtSA.SetText IIf(Trim(text(0)) <> "", "SA. " & Trim(text(1)), "")
    Report.txtTanggal.SetText "TANGGAL : " & TglDDMMMMYYYY(Format(tanggal1, "DD/MM/YYYY")) & " s/d " & TglDDMMMMYYYY(Format(tanggal2, "DD/MM/YYYY"))
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub

Private Sub repoFIR()
Dim oRs As Object
Set Report = New rptRepo3DFU
Dim Txt As String
Dim txtTemp As String
Screen.MousePointer = vbHourglass

    Txt = ""
    txtTemp = ""
    
    ''' REPO FIR
    If Trim(text(0)) <> "" Then
        Txt = "AND FRONT = '" & Trim(text(0)) & "'"
        txtTemp = "AND SA = '" & Trim(text(0)) & "'"
    End If
         
    sSQL = "SELECT DATE, (SELECT COUNT(*) FROM JOBDTA WITH(NOLOCK) WHERE DATE = A.DATE AND FS_LOG_DEL = '' " & Txt & ") AS TUS, " _
         & "(SELECT COUNT(*) FROM JOBDTA WITH(NOLOCK) WHERE DATE = A.DATE AND FN_RETURN = '0' AND TPSVC <> 'W' AND PDS = 0 AND FS_RC <> 'R001' AND FS_RC <> '' AND FS_LOG_DEL = '' " & Txt & ") AS CPUS, " _
         & "(SELECT COUNT(*) FROM JOBDTA WITH(NOLOCK) WHERE DATE = A.DATE AND FN_RETURN = '0' AND TPSVC <> 'W' AND PDS = 0 AND FS_RC = '' AND FS_LOG_DEL = '' " & Txt & ") AS GRP, " _
         & "(SELECT COUNT(*) FROM JOBDTA WITH(NOLOCK) WHERE DATE = A.DATE AND TPSVC = 'W' AND FS_LOG_DEL = '' " & Txt & ") AS TWC, (SELECT COUNT(*) FROM JOBDTA WITH(NOLOCK) WHERE DATE = A.DATE AND FN_RETURN = '1' AND FS_LOG_DEL = '' " & Txt & ") AS RTJ, " _
         & "(SELECT COUNT(*) FROM JOBDTL K WITH(NOLOCK) INNER JOIN JOBDTA L WITH(NOLOCK) ON K.NO_WO = L.NO_WO WHERE TGLFU = A.DATE AND L.FS_LOG_DEL = '' " & Txt & ") AS LISTFU, " _
         & "(SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' " & txtTemp & ") AS OKCALL, " _
         & "(SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' AND KUALITAS <> 'P' " & txtTemp & ") AS YQ1, " _
         & "(SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' AND KUALITAS = 'P' AND TEPAT <> 'P' " & txtTemp & ") AS YQ2, " _
         & "(SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' AND KUALITAS = 'P' AND TEPAT = 'P' AND LAMA <> 'P' " & txtTemp & ") AS YQ3, " _
         & "CASE WHEN (SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y') = 0 THEN 0 ELSE ROUND((1 - ((CAST((SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' AND KUALITAS <> 'P' " & txtTemp & ") AS FLOAT)+CAST((SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' AND KUALITAS = 'P' AND TEPAT <> 'P' " & txtTemp & ") AS FLOAT)+CAST((SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' AND KUALITAS = 'P' AND TEPAT = 'P' AND LAMA <> 'P' " & txtTemp & ") AS FLOAT))/(CAST(CASE (SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' " & txtTemp & ") WHEN 0 THEN 1 ELSE (SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' " & txtTemp & ") END AS FLOAT)))) * 100,0) END AS FIR, " _
         & "ROUND((CAST((SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' AND KUALITAS <> 'P' " & txtTemp & ") AS FLOAT) / CAST(CASE (SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' " & txtTemp & ") WHEN 0 THEN 1 ELSE (SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE TANGGAL = A.DATE AND SAMBUNG = 'Y' " & txtTemp & ") END AS FLOAT)) * 100,2) AS Q1RATIO, " _
         & "(SELECT COUNT(*) FROM DTLFU WITH(NOLOCK) WHERE NO_WO IN (SELECT NO_WO FROM JOBDTL WITH(NOLOCK) WHERE TGLFU = A.DATE) AND SAMBUNG = 'Y' " & txtTemp & ") AS TRGTCALL " _
         & "FROM JOBDTA A WITH(NOLOCK) WHERE DATE BETWEEN '" & Format(tanggal1, "MM/DD/YYYY") & "' AND '" & Format(tanggal2, "MM/DD/YYYY") & "' GROUP BY DATE ORDER BY DATE"
         
    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtSA.SetText IIf(Trim(text(0)) <> "", "SA. " & Trim(text(1)), "")
    Report.txtTanggal.SetText "TANGGAL : " & TglDDMMMMYYYY(Format(tanggal1, "DD/MM/YYYY")) & " s/d " & TglDDMMMMYYYY(Format(tanggal2, "DD/MM/YYYY"))
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height + 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > (asxPowerBanner1.Width + 100) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
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
            Case 0
                gMainKeyCari = text(Index)
                gSearchOption = "SA"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT FS_NM_PETUGAS FROM T_PETUGAS WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & text(Index) & "'")
                If oRs.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "SA"
                    frmCari.Show (vbModal)
                Else
                    text(Index + 1) = oRs("FS_NM_PETUGAS")
                End If
            Else
                text(Index + 1) = ""
            End If
    End Select
End Sub


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 345
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdCetak.Top = CRViewer1.Top + 345
    cmdCetak.Left = cboPrinter.Left + cboPrinter.Width + 30
End Sub

Private Sub DisableExport()
    Dim objPrinter As Printer
    Set objPrinter = Printer
    If InStr(UCase(objPrinter.DeviceName), "PDF") <> 0 Then
        cboPrinter = ""
    Else
        cboPrinter = objPrinter.DeviceName
    End If
    cboPrinter.Clear
    
    For Each pr In Printers
        If InStr(UCase(pr.DeviceName), "PDF") = 0 Then cboPrinter.AddItem pr.DeviceName
    Next pr
    If Log.PdfOnly = 0 Then
        CRViewer1.EnableExportButton = True
        CRViewer1.EnablePrintButton = True
        cmdExport.Visible = False
        cboPrinter.Visible = False
        cmdCetak.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdCetak_Click()
    For Each pr In Printers
        If UCase(pr.DeviceName) = UCase(cboPrinter) Then
            GoTo Cetak
        End If
    Next pr
    MsgBox "Printer tidak ditemukan Pilih Printer dahulu", vbInformation, "Proses Cetak Gagal"
    cboPrinter.SetFocus
    Exit Sub
    
Cetak:
    Dim Psize As Integer
    Dim POr As Integer
    
    Psize = Report.PaperSize
    POr = Report.PaperOrientation
    
    Set Printer = Printers(cboPrinter.ListIndex)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    Report.PaperOrientation = POr
    Report.PaperSize = Psize
    Report.PrintOut
    
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Salah
    Dim NamaFile As String
    NamaFile = ""
    
    Dim f As New FormExplorer
    f.Show (vbModal)
    Dim MyPath As String
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    
    If Log.PdfOnly = 1 Then
         
        NamaFile = MyPath & f.FileName
        
        If UCase(Right(Trim(f.FileName), 5)) = ".XLSX" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 5)
        If UCase(Right(Trim(f.FileName), 4)) = ".XLS" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 4)
         
        
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = NamaFile & ".Doc"
        Report.ExportOptions.RTFExportAllPages = True
        Report.ExportOptions.FormatType = crEFTExactRichText
        Report.Export False
        MsgBox NamaFile & ".Doc" & "  ... Data sudah di Export !", vbInformation
        Exit Sub
        With msg
            .PassingMsg = NamaFile & ".Doc" & "  ... Data sudah di Export !"
            .passingTipe = 2
            .Show vbModal
        End With
        Exit Sub
    ElseIf onExpo = 0 Then
        Report.Export
    End If
    Exit Sub
Salah:
    If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
    MsgBox "Err.Line : " & Erl & "Err.Numb : " & err.Number & "Err.Desc. : " & err.Description & vbCr & "Procedure : " & Me.Name & " export"
Keluar:
End Sub

