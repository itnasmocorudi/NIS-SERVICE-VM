VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoMonitoringMOU 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Monitoring Supply Slip"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   13875
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
      Left            =   1665
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4305
      Width           =   1440
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
      Left            =   585
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4305
      Width           =   1050
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5475
      TabIndex        =   16
      Top             =   4335
      Visible         =   0   'False
      Width           =   3015
   End
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
      Left            =   8595
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4305
      Visible         =   0   'False
      Width           =   1185
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
      Left            =   3150
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4305
      Visible         =   0   'False
      Width           =   1185
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   4365
      TabIndex        =   19
      Top             =   4320
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
      MICON           =   "frmRepoMonitoringMOU.frx":0000
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
      Left            =   480
      TabIndex        =   9
      Top             =   3840
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   1890
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1890
   End
   Begin VB.CommandButton tblRefresh 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Refresh"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1890
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Top             =   600
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox9 
      Height          =   270
      Left            =   7050
      TabIndex        =   10
      Top             =   600
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
      Left            =   1530
      TabIndex        =   1
      Top             =   960
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   3000
      TabIndex        =   12
      Top             =   1380
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
      Index           =   2
      Left            =   1530
      TabIndex        =   2
      Top             =   1380
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "??"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   1380
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      Top             =   1380
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "CCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   5
      Left            =   1530
      TabIndex        =   5
      Top             =   1700
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Polisi  :"
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
      Left            =   645
      TabIndex        =   13
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. MOU  :"
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
      Left            =   645
      TabIndex        =   11
      Top             =   645
      Width           =   795
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9735
      Left            =   0
      Picture         =   "frmRepoMonitoringMOU.frx":001C
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoMonitoringMOU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As Object
Private Sub hapusForm()
    For i = 0 To 5
        Text(i) = ""
    Next i
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
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
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call hapusForm
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean
Dim noPol As String
Dim pesan As String

    boleh = True
    pesan = ""
    noPol = Trim(Text(2)) & IIf(Trim(Text(2)) <> "", "-", "") & Trim(Text(3)) & IIf(Trim(Text(4)) <> "", "-", "") & Trim(Text(4))
    If Trim(Text(0)) = "" And noPol = "" Then
        boleh = False
        pesan = "Isi Salah satu filter diatas"
    End If
    
    If boleh Then
        If Trim(Text(0)) <> "" And Trim(noPol) = "" Then
            Call Cetak(Trim(noPol), 2)
        Else
            Call Cetak(Trim(noPol), 1)
        End If
    Else
        MsgBox pesan, vbOKOnly + vbDefaultButton1, "Kelengkapan Data"
    End If
End Sub

Private Sub Cetak(noPolis As String, idx As Integer)
      Dim rsRepo As ADODB.Recordset
      Dim Txt As String
1     Screen.MousePointer = vbHourglass
          
2         Set Report = New RptRepoMonitoringMOU
          
3         Select Case idx
              Case 1
                  '' BY NO POLISI
                  
4                 If Log.RPLHistori = 0 Then
5                     Set rsRepo = objO.FetchData("SELECT A.NO_POLISI, A.NO_CHASSIS, A.CUSTOMER, A.FMDL, A.NO_WO, A.DATE, A.NOINV, A.INVDATE, B.TOTAL, D.* " & _
                                              "FROM JOBDTA A WITH (NOLOCK) LEFT JOIN T_INVOICE B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NOINV = REPLACE(REPLACE(B.NOINV,'-A',''),'-B','') AND B.NOINV NOT LIKE '%-B' " & _
                                              "LEFT JOIN MOUDTL C WITH (NOLOCK) ON A.NO_CHASSIS = C.NO_RANGKA LEFT JOIN MOUHDR D WITH (NOLOCK) ON C.NO_MOU = D.NO_MOU " & _
                                              "WHERE A.NO_CHASSIS IN (SELECT NO_RANGKA FROM MOUDTL WHERE NO_MOU IN (SELECT NO_MOU FROM MOUDTL A INNER JOIN CUSTDAT B ON A.NO_RANGKA = B.NO_CHASSIS WHERE B.NO_POLISI = '" & Trim(noPolis) & "')) " & _
                                              "AND A.RPR_COND <> 'K' AND B.TPSVC <> 'W' AND B.FS_RC <> 'R001' AND B.NOINV Not like '%-B' AND A.NOINV <> '' AND A.DELIV <> 'BTL' ORDER BY A.NOINV")
6                 Else
7                     Set rsRepo = objO.FetchData("SELECT A.NO_POLISI, A.NO_CHASSIS, A.CUSTOMER, A.FMDL, A.NO_WO, A.DATE, A.NOINV, A.INVDATE, B.TOTAL, D.* " & _
                                              "FROM JOBDTA A WITH (NOLOCK) LEFT JOIN T_INVOICE B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NOINV = REPLACE(REPLACE(B.NOINV,'-A',''),'-B','') AND B.NOINV NOT LIKE '%-B' " & _
                                              "LEFT JOIN [" & dbReplikasiHO & "]..MOUDTL C WITH (NOLOCK) ON A.NO_CHASSIS = C.NO_RANGKA LEFT JOIN [" & dbReplikasiHO & "]..MOUHDR D WITH (NOLOCK) ON C.NO_MOU = D.NO_MOU " & _
                                              "WHERE A.NO_CHASSIS IN (SELECT NO_RANGKA FROM [" & dbReplikasiHO & "]..MOUDTL WHERE NO_MOU IN (SELECT NO_MOU FROM [" & dbReplikasiHO & "]..MOUDTL A INNER JOIN CUSTDAT B ON A.NO_RANGKA = B.NO_CHASSIS WHERE B.NO_POLISI = '" & Trim(noPolis) & "')) " & _
                                              "AND A.RPR_COND <> 'K' AND B.TPSVC <> 'W' AND B.FS_RC <> 'R001' AND B.NOINV Not like '%-B' AND A.NOINV <> '' AND A.DELIV <> 'BTL' ORDER BY A.NOINV")
8                 End If
9             Case 2
                  '' BY NO MOU
                  
10                If Log.RPLHistori = 0 Then
11                    Set rsRepo = objO.FetchData("SELECT DISTINCT C.NO_POLISI, C.NO_CHASSIS, C.CUSTOMER, C.FMDL, C.NO_WO, C.DATE, C.NOINV, C.INVDATE, C.TOTAL, A.* " & _
                                              "FROM MOUHDR A WITH (NOLOCK) INNER JOIN MOUDTL B WITH (NOLOCK) ON A.NO_MOU = B.NO_MOU " & _
                                              "LEFT JOIN T_INVOICE C WITH (NOLOCK) ON B.NO_RANGKA = C.NO_CHASSIS AND DELIV <> 'BTL' AND C.FS_LOG_DEL = '' AND RPR_COND <> 'K' " & _
                                              "WHERE A.NO_MOU = '" & Trim(Text(0)) & "' AND C.TPSVC <> 'W' AND C.FS_RC <> 'R001' AND C.NOINV Not like '%-B' AND ISNULL(C.NO_POLISI, '') <> '' ORDER BY NOINV")
12                Else
13                    Set rsRepo = objO.FetchData("SELECT DISTINCT C.NO_POLISI, C.NO_CHASSIS, C.CUSTOMER, C.FMDL, C.NO_WO, C.DATE, C.NOINV, C.INVDATE, C.TOTAL, A.* " & _
                                              "FROM [" & dbReplikasiHO & "]..MOUHDR A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..MOUDTL B WITH (NOLOCK) ON A.NO_MOU = B.NO_MOU " & _
                                              "LEFT JOIN T_INVOICE C WITH (NOLOCK) ON B.NO_RANGKA = C.NO_CHASSIS AND DELIV <> 'BTL' AND C.FS_LOG_DEL = '' AND RPR_COND <> 'K' " & _
                                              "WHERE A.NO_MOU = '" & Trim(Text(0)) & "' AND C.TPSVC <> 'W' AND C.FS_RC <> 'R001' AND C.NOINV Not like '%-B' AND ISNULL(C.NO_POLISI, '') <> '' ORDER BY NOINV")
14                End If
15        End Select
          
16        Report.database.SetDataSource rsRepo
17        Report.txtPerusahaan.SetText gPerush
18        Report.txtAlamat.SetText gAlamat
19        Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
          
20        With CRViewer1
21            .ReportSource = Report
22            .ViewReport
23            .Zoom 100
24            .Visible = True
25        End With
26        cmdTutup.Visible = True
27        cmdSettingPrinter.Visible = True
28        Screen.MousePointer = vbDefault
29        Call DisableExport
30        Call setLokasiTombol
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
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
Dim noPol As String

    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "MOU"
                frmCari.Show (vbModal)
            Case 2
                noPol = Trim(Text(2)) & IIf(Trim(Text(2)) <> "", "-", "") & Trim(Text(3)) & IIf(Trim(Text(4)) <> "", "-", "") & Trim(Text(4))
                gMainKeyCari = noPol
                gBantuan = "Monitoring MOU"
                gSearchOption = "CUSTOMER2"
                frmCari.Show (vbModal)
                Dim nDpn1 As String
                Dim nTgh1 As String
                Dim nBlk1 As String
                nDpn1 = ""
                nTgh1 = ""
                nBlk1 = ""
                Call pecahNoPol(gMainKeyCari, nDpn1, nTgh1, nBlk1)
                Text(2) = Trim(nDpn1)
                Text(3) = Trim(nTgh1)
                Text(4) = Trim(nBlk1)
                gBantuan = ""
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
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset
Dim noPol As String

    Text(Index) = Replace(Trim(Text(Index)), "'", "`")
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                
                If Log.RPLHistori = 0 Then
                    Set rsSearch = objO.FetchData("SELECT * FROM MOUHDR WITH (NOLOCK) WHERE NO_MOU = '" & Trim(Text(Index)) & "'")
                Else
                    Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MOUHDR WITH (NOLOCK) WHERE NO_MOU = '" & Trim(Text(Index)) & "'")
                End If
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MOU"
                    frmCari.Show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("PERUSAHAAN")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 4
            If Text(Index) <> "" Then
                Dim nDpn1 As String
                Dim nTgh1 As String
                Dim nBlk1 As String
                nDpn1 = ""
                nTgh1 = ""
                nBlk1 = ""

                noPol = Trim(Text(2)) & IIf(Trim(Text(2)) <> "", "-", "") & Trim(Text(3)) & IIf(Trim(Text(4)) <> "", "-", "") & Trim(Text(4))
                Set rsSearch = objO.FetchData("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_POLISI = '" & Trim(noPol) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = noPol
                    gBantuan = "Monitoring MOU"
                    gSearchOption = "CUSTOMER2"
                    frmCari.Show (vbModal)
                    Call pecahNoPol(gMainKeyCari, nDpn1, nTgh1, nBlk1)
                    Text(2) = Trim(nDpn1)
                    Text(3) = Trim(nTgh1)
                    Text(4) = Trim(nBlk1)
                    gBantuan = ""
                Else
                    Call pecahNoPol(noPol, nDpn1, nTgh1, nBlk1)
                    Text(2) = Trim(nDpn1)
                    Text(3) = Trim(nTgh1)
                    Text(4) = Trim(nBlk1)
                    Text(5) = rsSearch("CUSTOMER")
                End If
            Else
                Text(5) = ""
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
    
    cmdPrintOut.Top = CRViewer1.Top + 345
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




