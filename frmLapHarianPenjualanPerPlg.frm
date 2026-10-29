VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLapHarianPenjualanPerPlg 
   Caption         =   "Laporan Harian Penjualan Per Pelanggan"
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
   Begin VB.CommandButton TblConvert 
      Caption         =   "&Simpan Text"
      Height          =   390
      Left            =   2265
      TabIndex        =   9
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Export Text"
      Height          =   390
      Left            =   4155
      TabIndex        =   8
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
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmLapHarianPenjualanPerPlg.frx":0000
      Left            =   1560
      List            =   "frmLapHarianPenjualanPerPlg.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Width           =   735
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
      FormatString    =   "Laporan Harian Penjualan Per Pelanggan"
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
      TabIndex        =   10
      Top             =   3180
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmLapHarianPenjualanPerPlg.frx":0018
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   2
      Left            =   3210
      TabIndex        =   11
      Top             =   870
      Width           =   3300
      _ExtentX        =   5821
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
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   57
      Left            =   1560
      TabIndex        =   12
      Top             =   870
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox9 
      Height          =   270
      Left            =   3090
      TabIndex        =   13
      Top             =   885
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Plg"
      Height          =   255
      Left            =   495
      TabIndex        =   14
      Top             =   930
      Width           =   945
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   570
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
Attribute VB_Name = "frmLapHarianPenjualanPerPlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HapusForm()
    Text(0) = ">>->>->>>>"
End Sub

Private Sub cboStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombolCombo(KeyCode)
End Sub

Private Sub Form_Load()
    Call HapusForm
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptLapHarianPenjualanPerPlg
    Dim vNamaTabel  As String
    Dim Bulan As String
    Dim Tahun As String
    
    Bulan = Mid(Text(0).Text, 4, 2)
    Tahun = Right(Text(0).Text, 4)

    If Text(0) = "" Then
        MsgBox "Masukkan Tgl SO Masuk !"
    Else
        gTglTemp = Text(0)
        gKetTemp = "(" & cboStatus.Text & ")"
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        
        vNamaTabel = "Harian_Penjualan" & DriveSerialNumber("c:")
        
        Call BuatTabelInv(vNamaTabel)
        
        sSQL = "insert into [" & vNamaTabel & "] select * from T_INVOICE " _
             & " where month([invdate])='" & Bulan & "' and year([invdate])='" & Tahun & "' " _
             & " and STATUS='" & cboStatus.Text & "'" _
             & " union all " _
             & " select " & NmFieldRetur & " from T_INVOICE where month([fd_retur])='" & Bulan & "' and year([fd_retur])='" & Tahun & "' " _
             & " and STATUS='" & cboStatus.Text & "'"
        oConn.Execute sSQL
        
        If Trim(Text(57)) = "" Then
            sSQL = "select custdat.fs_kd_plg, X.customer, FN_RETURN, FS_NO_RETUR, custdat.CUST_ADRES, custdat.NO_POLISI ,custdat.FMDL, NO_WO, RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7 as SUB_RATE, STATUS, NOINV,TOTAL-MET as TOTAL,BODY,LABOR,SBLT,MET,OIL,MTRL,PART, TOTAL as GTOTAL, TOTAL-MET-FN_PAJAK as NETTO, FN_PAJAK as FN_PAJAK," _
                 & " (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]=x.[invdate] and FS_KREDIT<>'1' and NOINV=x.NOINV ) as CASH, (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]=x.[invdate] and FS_KREDIT='1'and NOINV=x.NOINV ) as KREDIT," _
                 & " (select count(distinct NO_WO) from [" & vNamaTabel & "] where [invdate]=x.[invdate] and STATUS='" & cboStatus.Text & "') as JML_SO, (select count(distinct NO_WO) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_SO, (select sum(RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_RATE," _
                 & " (select sum(LABOR) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_JASA, (select sum(SBLT) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_SBLT," _
                 & " (select sum(MET) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_MET, (select sum(OIL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_OIL," _
                 & " (select sum(MTRL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_MTRL, (select sum(PART) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_PART," _
                 & " (select sum(TOTAL-MET) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_TOTAL, (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_GTOTAL," _
                 & " (select sum(TOTAL-MET-FN_PAJAK) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_NETTO, (select sum(FN_PAJAK) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_PAJAK," _
                 & " (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and FS_KREDIT<>'1' and STATUS='" & cboStatus.Text & "') as TOT_CASH, (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and FS_KREDIT='1'and  FS_NO_RETUR = '' and STATUS='" & cboStatus.Text & "') as TOT_KREDIT," _
                 & " (select count(NOINV) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TTL_MOBIL_INV, (select count(NOINV) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TTL_INV" _
                 & " from [" & vNamaTabel & "] x inner join custdat on x.no_polisi = custdat.no_polisi where [invdate]='" & TglYYYYMMDD(Text(0)) & "' and STATUS='" & cboStatus.Text & "' order by custdat.fs_kd_plg, FS_NO_RETUR, FN_RETURN, NOINV"
        Else
            sSQL = "select custdat.fs_kd_plg, X.customer, FN_RETURN, FS_NO_RETUR, custdat.CUST_ADRES, custdat.NO_POLISI ,custdat.FMDL, NO_WO, RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7 as SUB_RATE, STATUS, NOINV,TOTAL-MET as TOTAL,BODY,LABOR,SBLT,MET,OIL,MTRL,PART, TOTAL as GTOTAL, TOTAL-MET-FN_PAJAK as NETTO, FN_PAJAK as FN_PAJAK," _
                 & " (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]=x.[invdate] and FS_KREDIT<>'1' and NOINV=x.NOINV ) as CASH, (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]=x.[invdate] and FS_KREDIT='1'and NOINV=x.NOINV ) as KREDIT," _
                 & " (select count(distinct NO_WO) from [" & vNamaTabel & "] where [invdate]=x.[invdate] and STATUS='" & cboStatus.Text & "') as JML_SO, (select count(distinct NO_WO) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_SO, (select sum(RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_RATE," _
                 & " (select sum(LABOR) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_JASA, (select sum(SBLT) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_SBLT," _
                 & " (select sum(MET) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_MET, (select sum(OIL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_OIL," _
                 & " (select sum(MTRL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_MTRL, (select sum(PART) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_PART," _
                 & " (select sum(TOTAL-MET) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_TOTAL, (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_GTOTAL," _
                 & " (select sum(TOTAL-MET-FN_PAJAK) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_NETTO, (select sum(FN_PAJAK) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TOT_PAJAK," _
                 & " (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and FS_KREDIT<>'1' and STATUS='" & cboStatus.Text & "') as TOT_CASH, (select sum(TOTAL) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and FS_KREDIT='1'and  FS_NO_RETUR = '' and STATUS='" & cboStatus.Text & "') as TOT_KREDIT," _
                 & " (select count(NOINV) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TTL_MOBIL_INV, (select count(NOINV) from [" & vNamaTabel & "] where [invdate]<=x.[invdate] and month([invdate])=month(x.[invdate]) and year([invdate])=year(x.[invdate]) and STATUS='" & cboStatus.Text & "') as TTL_INV" _
                 & " from [" & vNamaTabel & "] x inner join custdat on x.no_polisi = custdat.no_polisi where [invdate]='" & TglYYYYMMDD(Text(0)) & "' and STATUS='" & cboStatus.Text & "' and custdat.fs_kd_plg ='" & Text(57) & "' order by custdat.fs_kd_plg, FS_NO_RETUR, FN_RETURN, NOINV"
        End If
        rpt.DataControl1.ConnectionString = gReportConn
        rpt.DataControl1.Source = sSQL
'        rpt.Printer.PaperSize = 14
'        rpt.Printer.Orientation = 2
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
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 57
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.Show (vbModal)
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
        Case 0, 1
            If CekTgl(Text(Index)) <> True Then
                Text(Index) = gTanggal
            End If
        Case 57
            If Trim(Text(Index)) <> "" And gConnecAAS = True Then
                'LOGONSTR2 = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=auser;Data Source=" & LetakAcces & "ardata" & Trim(gGenerasi) & ".MDB" & ";Persist Security Info=False;Jet OLEDB:System database=" & SysAcces
                LOGONSTR2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";Password='" & NisPass & "';Initial Catalog= " & NisDb & ";Data Source= " & NisKomp & ""
                
                oConnAcces.ConnectionString = LOGONSTR2
                oConnAcces.CommandTimeout = 120
                oConnAcces.ConnectionTimeout = 120
                oConnAcces.Open
                
                sSQL = "select KodeLgn, NamaLgn from Armst where KodeLgn = '" & Trim(Text(Index)) & "'"
                oRs.Open sSQL, oConnAcces, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "KODEPLG"
                    frmCariAcess.Show (vbModal)
                Else
                    If IsNull(oRs("NamaLgn")) Then
                        Text(2) = ""
                    Else
                        Text(2) = oRs("NamaLgn")
                    End If
                End If
                oConnAcces.Close
            End If
    End Select
End Sub

