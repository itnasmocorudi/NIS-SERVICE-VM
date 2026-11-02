VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoMrsDatang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan MRS untuk Kedatangan Service"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   12735
   Begin VB.CommandButton cmdPrintOut 
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
      Left            =   9465
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2355
      Visible         =   0   'False
      Width           =   1185
   End
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
      Left            =   1455
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2325
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
      Left            =   2565
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2325
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
      Left            =   4050
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2325
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   6375
      TabIndex        =   10
      Top             =   2355
      Visible         =   0   'False
      Width           =   3015
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   270
      Left            =   5325
      TabIndex        =   11
      Top             =   2325
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   476
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
      MICON           =   "frmRepoMrsDatang.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoMrsDatang.frx":001C
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   5
      Top             =   0
      Width           =   445
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   6120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   7095
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin TDBDate6Ctl.TDBDate DtDOAwal 
      Height          =   330
      Left            =   1785
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmRepoMrsDatang.frx":2364
      Caption         =   "frmRepoMrsDatang.frx":2490
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoMrsDatang.frx":24FC
      Keys            =   "frmRepoMrsDatang.frx":251A
      Spin            =   "frmRepoMrsDatang.frx":2578
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   2118189057
      Value           =   2.12482692446619E-314
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate DtDOAkhir 
      Height          =   330
      Left            =   3825
      TabIndex        =   2
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmRepoMrsDatang.frx":25A0
      Caption         =   "frmRepoMrsDatang.frx":26CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoMrsDatang.frx":2738
      Keys            =   "frmRepoMrsDatang.frx":2756
      Spin            =   "frmRepoMrsDatang.frx":27B4
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   2118189057
      Value           =   2.12482692446619E-314
      CenturyMode     =   0
   End
   Begin ucXPButton.XPButton cmdCetak 
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&p"
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
      MICON           =   "frmRepoMrsDatang.frx":27DC
      PICN            =   "frmRepoMrsDatang.frx":27F8
      PICH            =   "frmRepoMrsDatang.frx":2C4F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdKeluar 
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   1200
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
      MICON           =   "frmRepoMrsDatang.frx":3139
      PICN            =   "frmRepoMrsDatang.frx":3155
      PICH            =   "frmRepoMrsDatang.frx":3629
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   5520
      X2              =   480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   975
      Left            =   480
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: KEDATANGAN ::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2040
      TabIndex        =   6
      Top             =   180
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3345
      TabIndex        =   4
      Top             =   645
      Width           =   345
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   0
      Picture         =   "frmRepoMrsDatang.frx":3B74
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmRepoMrsDatang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As Object

Private Sub cmdCetak_Click()
Dim rsRepo As ADODB.Recordset
Set Report = New RptMRSKedatangan
    
    If Trim(DtDOAwal) <> "" And Trim(DtDOAkhir) <> "" Then
        
        If Log.RPLHistori = 0 Then
            sSQL = "SELECT DATE, " & _
               "(SELECT COUNT(DISTINCT NO_CHASSIS) FROM JOBDTA A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC INNER JOIN M_ESTIMASI C WITH(NOLOCK) ON A.NO_CHASSIS = C.NORANGKA AND C.SBENEXT = B.FS_NILAI " & _
               "WHERE ISNULL(JT_BLN,'') <> '' AND ISNULL(JT_TGL,'')<> ''  AND A.DATE = Q.DATE AND (DATEDIFF(d, (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), DATE) > 15 OR (KM-KMNEXT) > 1000)) AS early, " & _
               "(SELECT COUNT(DISTINCT NO_CHASSIS) FROM JOBDTA A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC INNER JOIN M_ESTIMASI C WITH(NOLOCK) ON A.NO_CHASSIS = C.NORANGKA AND C.SBENEXT = B.FS_NILAI " & _
               "WHERE ISNULL(JT_BLN,'') <> '' AND ISNULL(JT_TGL,'')<> ''  AND A.DATE = Q.DATE AND (ABS(DATEDIFF(d, (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), DATE)) <= 15 AND ABS(KMNEXT-KM) <= 1000)) AS akurat, " & _
               "(SELECT COUNT(DISTINCT NO_CHASSIS) FROM JOBDTA A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC INNER JOIN M_ESTIMASI C WITH(NOLOCK) ON A.NO_CHASSIS = C.NORANGKA AND C.SBENEXT = B.FS_NILAI " & _
               "WHERE ISNULL(JT_BLN,'') <> '' AND ISNULL(JT_TGL,'')<> ''  AND A.DATE = Q.DATE AND (DATEDIFF(d, (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), DATE) < 0 OR (KMNEXT-KM) > 1000)) AS late " & _
               "FROM JOBDTA Q WITH(NOLOCK) WHERE DATE BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' GROUP BY DATE ORDER BY DATE"
        Else
            sSQL = "SELECT DATE, " & _
               "(SELECT COUNT(DISTINCT NO_CHASSIS) FROM JOBDTA A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC INNER JOIN M_ESTIMASI C WITH(NOLOCK) ON A.NO_CHASSIS = C.NORANGKA AND C.SBENEXT = B.FS_NILAI " & _
               "WHERE ISNULL(JT_BLN,'') <> '' AND ISNULL(JT_TGL,'')<> ''  AND A.DATE = Q.DATE AND (DATEDIFF(d, (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), DATE) > 15 OR (KM-KMNEXT) > 1000)) AS early, " & _
               "(SELECT COUNT(DISTINCT NO_CHASSIS) FROM JOBDTA A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC INNER JOIN M_ESTIMASI C WITH(NOLOCK) ON A.NO_CHASSIS = C.NORANGKA AND C.SBENEXT = B.FS_NILAI " & _
               "WHERE ISNULL(JT_BLN,'') <> '' AND ISNULL(JT_TGL,'')<> ''  AND A.DATE = Q.DATE AND (ABS(DATEDIFF(d, (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), DATE)) <= 15 AND ABS(KMNEXT-KM) <= 1000)) AS akurat, " & _
               "(SELECT COUNT(DISTINCT NO_CHASSIS) FROM JOBDTA A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC INNER JOIN M_ESTIMASI C WITH(NOLOCK) ON A.NO_CHASSIS = C.NORANGKA AND C.SBENEXT = B.FS_NILAI " & _
               "WHERE ISNULL(JT_BLN,'') <> '' AND ISNULL(JT_TGL,'')<> ''  AND A.DATE = Q.DATE AND (DATEDIFF(d, (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), DATE) < 0 OR (KMNEXT-KM) > 1000)) AS late " & _
               "FROM JOBDTA Q WITH(NOLOCK) WHERE DATE BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' GROUP BY DATE ORDER BY DATE"
        End If
        Set rsRepo = objO.FetchData(sSQL)
        Report.database.SetDataSource rsRepo
        Report.txtPerusahaan.SetText gPerush
        Report.txtAlamat.SetText gAlamat
        Report.txtTanggal.SetText "dari Tanggal " & TglDDMMMMYYYY(Format(DtDOAwal, "DD/MM/YYYY"), 1) & "    ~    " & TglDDMMMMYYYY(Format(DtDOAkhir, "DD/MM/YYYY"), 1)
        
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 100
            .Top = 0
            .Left = Image1.Left + Image1.Width + 20
            .Width = Me.Width - 500
            .Height = Image1.Height
            
            cmdTutup.Top = .Top + 430
            cmdTutup.Left = .Left
            cmdTutup.Visible = True
            
            cmdSettingPrinter.Top = cmdTutup.Top
            cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 20
            cmdSettingPrinter.Visible = True
            .Visible = True
        End With
        cmdTutup.Visible = True
        Screen.MousePointer = vbDefault
        Call DisableExport
        Call setLokasiTombol
    End If
End Sub

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
    cmdPrintOut.Visible = False
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DtDOAwal = Format(isnow, "DD/MM/YYYY")
    DtDOAkhir = Format(isnow, "DD/MM/YYYY")
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
End Sub

Private Sub Form_Resize()
    Image1.Height = Me.Height - 50
End Sub


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 445
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 445
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 445
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 445
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 445
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdPrintOut.Top = CRViewer1.Top + 445
    cmdPrintOut.Left = cboPrinter.Left + cboPrinter.Width + 30
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
        cmdPrintOut.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdPrintOut_Click()
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

