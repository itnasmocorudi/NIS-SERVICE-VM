VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmLapHarianSOMasukPerSA 
   Caption         =   "Laporan Harian SO Masuk Per SA"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   13230
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option1 
      Caption         =   "Diurutkan Per Nomor WO"
      Height          =   240
      Left            =   855
      TabIndex        =   13
      Top             =   975
      Width           =   4260
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Diurutkan Per Nama Customer "
      Height          =   255
      Left            =   855
      TabIndex        =   12
      Top             =   1275
      Width           =   4245
   End
   Begin VB.CommandButton TblConvert 
      Caption         =   "&Simpan Text"
      Height          =   390
      Left            =   2265
      TabIndex        =   10
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Export Text"
      Height          =   390
      Left            =   4155
      TabIndex        =   9
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblLayar 
      Caption         =   "Cetak &Layar"
      Height          =   390
      Left            =   6045
      TabIndex        =   2
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      Caption         =   "Cetak &Printer"
      Height          =   390
      Left            =   7935
      TabIndex        =   3
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      TabIndex        =   4
      Top             =   7065
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   150
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
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7560
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13335
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Laporan Harian SO Masuk Per SA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
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
      Mask            =   ">CCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2175
      TabIndex        =   8
      Top             =   495
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5490
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   615
      TabIndex        =   11
      Top             =   3195
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmLapHarianSOMasukPerSA.frx":0000
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "T SA."
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   525
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2835
      TabIndex        =   6
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   180
      Width           =   960
   End
End
Attribute VB_Name = "frmLapHarianSOMasukPerSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub hapusForm()
    Text(0) = ">>->>->>>>"
End Sub

Private Sub Form_Load()
    Call hapusForm
    Text(1) = gSA
    Option1 = True
End Sub

Private Sub cetak(Layar As Byte)
    Dim rpt As New RptLapHarianSOMasukPerSA
    
    If Text(0) = "" Then
        MsgBox "Masukkan Tgl SO Masuk !"
    Else
        gTglTemp = Text(0)
        gKetTemp = "SA : " & Text(1)
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        
        If Option1.Value = True Then
            sSQL = "select CUSTOMER, CUST_ADRES, NO_POLISI ,FMDL, FRONT, [GROUP], " _
                 & " MECH, JOB, REP1, REP2, REP3, REP4, REP5, REP6, REP7, " _
                 & " REP8 ,SLIP, NO_WO, STATUS, MEKANIK1, MEKANIK2, MEKANIK3, MEKANIK4, " _
                 & " MEKANIK5, MEKANIK6, MEKANIK7, MEKANIK8, RTE1, RTE2, RTE3, RTE4, RTE5, " _
                 & " RTE6, RTE7, RTE8, KLPKERJA1, KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, " _
                 & " KLPKERJA6, KLPKERJA7, KLPKERJA8, STATUS, NOINV, TELPON from jobdta " _
                 & " where date = '" & TglYYYYMMDD(Text(0)) & "'" _
                 & " and FRONT = '" & Text(1) & "'" _
                 & " and DELIV <> 'BTL'" _
                 & " order by NO_WO, NOINV"
        Else
            sSQL = "select CUSTOMER, CUST_ADRES, NO_POLISI ,FMDL, FRONT, [GROUP], " _
                 & " MECH, JOB, REP1, REP2, REP3, REP4, REP5, REP6, REP7, " _
                 & " REP8 ,SLIP, NO_WO, STATUS, MEKANIK1, MEKANIK2, MEKANIK3, MEKANIK4, " _
                 & " MEKANIK5, MEKANIK6, MEKANIK7, MEKANIK8, RTE1, RTE2, RTE3, RTE4, RTE5, " _
                 & " RTE6, RTE7, RTE8, KLPKERJA1, KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, " _
                 & " KLPKERJA6, KLPKERJA7, KLPKERJA8, STATUS, NOINV, TELPON from jobdta " _
                 & " where date = '" & TglYYYYMMDD(Text(0)) & "'" _
                 & " and FRONT = '" & Text(1) & "'" _
                 & " and DELIV <> 'BTL'" _
                 & " order by CUSTOMER, NO_WO"
        End If
        
        rpt.DataControl1.ConnectionString = gReportConn
        rpt.DataControl1.Source = sSQL
'        rpt.Printer.PaperSize = 14
'        rpt.Printer.Orientation = ddOPortrait
        rpt.Printer.PaperSize = 39
        rpt.Printer.Orientation = ddOPortrait
        Select Case Layar
            Case 1
                rpt.Show
            Case 2
                ARViewer1.ReportSource = rpt
            Case Else
                rpt.PrintReport False
                Unload rpt
        End Select
    End If
End Sub

Private Sub TblCetak_Click()
    Call cetak(0)
End Sub

Private Sub TblConvert_Click()
    Dim txt As New ActiveReportsTextExport.ARExportText

    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.ShowSave
    
    txt.FileName = CommonDialog1.FileName
    txt.Export ARViewer1.Pages
End Sub

Private Sub Command2_Click()
    Call cetak(2)
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblLayar_Click()
    Call cetak(1)
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
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 1
                gMainKeyCari = Text(Index)
                gSearchOption = "PETUGAS"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If CekTgl(Text(Index)) <> True Then
                Text(Index) = gTanggal
            End If
        Case 1
            If Trim(Text(Index)) <> "" Then
                sSQL = "select * from t_petugas where FS_KD_PETUGAS='" & Text(Index) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "PETUGAS"
                    frmCari.Show (vbModal)
                End If
                oRs.Close
            End If
    End Select
End Sub

