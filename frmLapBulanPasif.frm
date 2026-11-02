VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLapBulanPasif 
   Caption         =   "Laporan Customer Pasif"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   8535
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&Export Text"
      Height          =   390
      Left            =   4170
      TabIndex        =   13
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblConvert 
      Caption         =   "&Simpan Text"
      Height          =   390
      Left            =   2280
      TabIndex        =   12
      Top             =   7065
      Width           =   1890
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tidak service dari :"
      Height          =   285
      Left            =   810
      TabIndex        =   9
      Top             =   255
      Width           =   1725
   End
   Begin VB.CommandButton TblLayar 
      Caption         =   "Cetak &Layar"
      Height          =   390
      Left            =   6060
      TabIndex        =   2
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      Caption         =   "Cetak &Printer"
      Height          =   390
      Left            =   7950
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
      Left            =   2760
      TabIndex        =   0
      Top             =   795
      Width           =   540
      _ExtentX        =   953
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
      Mask            =   "9999"
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
      FormatString    =   "Laporan Customer Pasif"
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
      Left            =   3780
      TabIndex        =   1
      Top             =   795
      Width           =   555
      _ExtentX        =   979
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
      Mask            =   "9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   2
      Left            =   2565
      TabIndex        =   8
      Top             =   240
      Width           =   540
      _ExtentX        =   953
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
      Mask            =   "9999"
      PromptChar      =   ">"
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Tidak service antara :"
      Height          =   285
      Left            =   810
      TabIndex        =   10
      Top             =   810
      Width           =   1920
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5505
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   630
      TabIndex        =   14
      Top             =   3180
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmLapBulanPasif.frx":0000
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Hari"
      Height          =   210
      Left            =   3225
      TabIndex        =   11
      Top             =   285
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   225
      Left            =   810
      TabIndex        =   7
      Top             =   1545
      Width           =   5085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hari"
      Height          =   210
      Left            =   4425
      TabIndex        =   6
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   825
      Width           =   285
   End
End
Attribute VB_Name = "frmLapBulanPasif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HapusForm()
    Text(0) = ""
    Text(1) = ""
End Sub

Private Sub Form_Load()
    Label2.Caption = "Dihitung mundur mulai dari tanggal : " & gTanggal
    Option1 = True
    Call HapusForm
End Sub

Private Sub Pilihan1()
    Label5.Enabled = True
    Text(2).Enabled = True
    Label3.Enabled = False
    Label1.Enabled = False
    Text(0).Enabled = False
    Text(1).Enabled = False
End Sub

Private Sub Pilihan2()
    Label5.Enabled = False
    Text(2).Enabled = False
    Label3.Enabled = True
    Label1.Enabled = True
    Text(0).Enabled = True
    Text(1).Enabled = True
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptLapBulanPasif
    
    gKetTemp = ""
    gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
    
'    sSQL = "select distinct a.DATE as DATE, a.NO_POLISI as NO_POLISI, a.CUSTOMER as CUSTOMER," _
         & " a.CUST_ADRES as CUST_ADRES, a.KOTA as KOTA, a.TELPON as TELPON, a.MDL as MDL," _
         & " a.FMDL as FMDL, a.NO_WO as NO_WO, a.FS_TH_MOBIL as FS_TH_MOBIL," _
         & " b.TGL_REP as TGL_REP from t_invoice as a, custdat as b where a.NO_POLISI = b.NO_POLISI" _
         & " and dateadd(day," & Val(Text(0)) & ",TGL_REP) <= '" & TglYYYYMMDD(gTanggal) & "' " _
         & " and dateadd(day," & Val(Text(1)) & ",TGL_REP) >= '" & TglYYYYMMDD(gTanggal) & "' " _
         & " and FS_NO_RETUR = ''" _
         & " order by NO_POLISI"
    
    If Option1.Value = True Then
        gTglTemp = Text(2)
        sSQL = "select TGL_REP as DATE, NO_POLISI, CUSTOMER," _
             & " CUST_ADRES, KOTA, TELPON, MDL," _
             & " FMDL, FS_TH_MOBIL," _
             & " TGL_REP from custdat where " _
             & " dateadd(day," & Val(Text(0)) & ",TGL_REP) <= '" & TglYYYYMMDD(gTanggal) & "' " _
             & " AND NO_POLISI NOT IN (SELECT NO_POLISI FROM JOBDTA where " _
             & " dateadd(day," & Val(Text(0)) & ",DATE) <= '" & TglYYYYMMDD(gTanggal) & "') " _
             & " order by NO_POLISI"
    Else
        gTglTemp = Text(0) & " s/d " & Text(1)
        sSQL = "select TGL_REP as DATE, NO_POLISI, CUSTOMER," _
             & " CUST_ADRES, KOTA, TELPON, MDL," _
             & " FMDL, FS_TH_MOBIL," _
             & " TGL_REP from custdat where " _
             & " dateadd(day," & Val(Text(0)) & ",TGL_REP) <= '" & TglYYYYMMDD(gTanggal) & "' " _
             & " and dateadd(day," & Val(Text(1)) & ",TGL_REP) >= '" & TglYYYYMMDD(gTanggal) & "' " _
             & " AND NO_POLISI IN (SELECT NO_POLISI FROM JOBDTA where " _
             & " dateadd(day," & Val(Text(0)) & ",DATE) <= '" & TglYYYYMMDD(gTanggal) & "' " _
             & " and dateadd(day," & Val(Text(1)) & ",DATE) >= '" & TglYYYYMMDD(gTanggal) & "') " _
             & " order by NO_POLISI"
    End If
    
    rpt.DataControl1.ConnectionString = gReportConn
    rpt.DataControl1.Source = sSQL
'    rpt.Printer.PaperSize = 14
'    rpt.Printer.Orientation = 2
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
End Sub

Private Sub Option1_Click()
    Call Pilihan1
End Sub

Private Sub Option2_Click()
    Call Pilihan2
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

