VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmInfoHistory 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Informasi Repair History"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton TblConvert 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Simpan Text"
      Height          =   390
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Export Text"
      Height          =   390
      Left            =   4170
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblLayar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Cetak &Layar"
      Height          =   390
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Cetak &Printer"
      Height          =   390
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1575
      TabIndex        =   0
      Top             =   150
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   3000
      TabIndex        =   5
      Top             =   165
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   2655
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   645
      TabIndex        =   8
      Top             =   3210
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmInfoHistory.frx":0000
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   11880
      TabIndex        =   10
      Top             =   240
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8895
      Left            =   0
      Picture         =   "frmInfoHistory.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Polisi"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   180
      Width           =   660
   End
End
Attribute VB_Name = "frmInfoHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub hapusForm()
    Text(0) = ""
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptInfoHistory
    Dim keteranganHide As String
    
    If Text(0) = "" Then
        MsgBox "Masukkan No. Polisi !"
    Else
        If Log.hideService = 1 Then
            keteranganHide = " and km > 10"
        End If
        
        gTglTemp = Text(0)
        gKetTemp = ""
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        sSQL = "select NOINV, NO_WO, NO_POLISI,[DATE],KM,FRONT, MECH AS MEKANIK1,MECH AS MEKANIK2,MECH AS MEKANIK3,MECH AS MEKANIK4,MECH AS MEKANIK5,MECH AS MEKANIK6,MECH AS MEKANIK7, " _
             & " CUSTOMER , REP1, REP2, REP3, REP4, REP5, REP6, REP7 " _
             & " from HISTORY WITH(NOLOCK) where no_chassis in (SELECT NO_CHASSIS FROM CUSTDAT WITH(NOLOCK) WHERE NO_POLISI= '" & Text(0) & "') AND FS_LOG_DEL = '' " & keteranganHide & "" _
             & " order by [DATE] DESC, NOINV DESC, KM DESC "
        
        rpt.DataControl1.ConnectionString = gReportConn
        rpt.DataControl1.Source = sSQL
        rpt.Printer.PaperSize = 14
        rpt.Printer.Orientation = ddOPortrait
        Select Case Layar
            Case 1
                rpt.show
            Case 2
                ARViewer1.ReportSource = rpt
            Case Else
                rpt.PrintReport False
                Unload rpt
        End Select
    End If
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > 100 Then
        CRViewer1.Height = ScaleHeight - 100
    End If
    If ScaleWidth > (asxPowerBanner1.Width + 100) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub TblCetak_Click()
    Call Cetak(0)
End Sub

Private Sub TblConvert_Click()
    Dim txt As New ActiveReportsTextExport.ARExportText

    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.ShowSave
    
    txt.FileName = CommonDialog1.FileName
    txt.Export ARViewer1.Pages
End Sub

Private Sub Command2_Click()
    Call Cetak(2)
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblLayar_Click()
    Call Cetak(1)
    'Call cetakHistory
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index) = gMainKeyCari
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
                gMainKeyCari = Text(Index)
                gSearchOption = "HISTORY"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub


Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset
    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                Set rsSearch = objO.FetchData("select * from HISTORY WITH(NOLOCK) where NO_CHASSIS IN (SELECT NO_CHASSIS FROM CUSTDAT WITH(NOLOCK) WHERE NO_POLISI='" & Text(0) & "') AND FS_LOG_DEL = ''")
                If rsSearch.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "HISTORY"
                    frmCari.show (vbModal)
                End If
            End If
    End Select
End Sub

Private Sub cetakHistory(noKa As String, noWO As String)
    Dim rpt As Object
    Set rpt = New RptInfoHistoryUnit
    
    Dim sHlp As Object  'New ADODB.Recordset
    Dim sHlp2 As Object
    Dim sHlp3 As Object
    Dim sHlp4 As Object
    Dim crxSubReportJobOrder As CRAXDRT.Report
    Dim crxSubReportPart As CRAXDRT.Report
    Dim crxSubReportBooking2 As CRAXDRT.Report
    Dim crxSubReportBookingFU As CRAXDRT.Report
    
    Dim rRs As ADODB.Recordset
    Set rRs = objO.FetchData("SELECT * FROM CUSTDAT A WITH(NOLOCK) INNER JOIN HISTORY B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS " _
                          & "LEFT JOIN JOBDTL C WITH(NOLOCK) ON B.NO_WO = C.NO_WO AND A.NO_CHASSIS = C.NO_CHASSIS " _
                          & "WHERE A.NO_CHASSIS = '" & Trim(noKa) & "' AND B.FS_LOG_DEL = ''")
    rpt.database.SetDataSource rRs
    
    sSQL = "SELECT X.TIPE, X.KERJA FROM (SELECT KLPKERJA1 AS TIPE, REP1 AS KERJA FROM JOBDTA WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noKa) & "' AND NO_WO = '" & Trim(noWO) & "' " _
         & "Union SELECT KLPKERJA2 AS TIPE, REP2 AS KERJA FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noKa) & "' AND NO_WO = '" & Trim(noWO) & "' " _
         & "Union SELECT KLPKERJA3 AS TIPE, REP3 AS KERJA FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noKa) & "' AND NO_WO = '" & Trim(noWO) & "' " _
         & "Union SELECT KLPKERJA4 AS TIPE, REP4 AS KERJA FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noKa) & "' AND NO_WO = '" & Trim(noWO) & "' " _
         & "Union SELECT KLPKERJA5 AS TIPE, REP5 AS KERJA FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noKa) & "' AND NO_WO = '" & Trim(noWO) & "' " _
         & "Union SELECT KLPKERJA6 AS TIPE, REP6 AS KERJA FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noKa) & "' AND NO_WO = '" & Trim(noWO) & "' " _
         & "Union SELECT KLPKERJA7 AS TIPE, REP7 AS KERJA FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noKa) & "' AND NO_WO = '" & Trim(noWO) & "' " _
         & "Union SELECT KLPKERJA AS TIPE, REP AS KERJA FROM HIST_JOBADD WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noKa) & "' AND NO_WO = '" & Trim(noWO) & "' AND FS_LOG_DEL = '' " _
         & "Union SELECT 'SUB' AS TIPE, NAMABARANG AS KERJA FROM HIST_SUBLET WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noKa) & "' AND NO_WO = '" & Trim(noWO) & "' " _
         & ") AS X WHERE X.KERJA <> ''"
         
    Set sHlp = objO.FetchData(sSQL)
    Set crxSubReportJobOrder = rpt.Subreport2.OpenSubreport
    crxSubReportJobOrder.database.SetDataSource sHlp
    
    sSQL = "SELECT KODEITEM, KETERANGAN, QTY, HARGA FROM T_PART WITH(NOLOCK) WHERE NO_WO = '" & Trim(noWO) & "'"
    Set sHlp2 = objO.FetchData(sSQL)
    Set crxSubReportPart = rpt.Subreport1.OpenSubreport
    crxSubReportPart.database.SetDataSource sHlp2

    With CRViewer1
        .ReportSource = rpt
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub
