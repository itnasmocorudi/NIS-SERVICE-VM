VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLapPenjServiceBln 
   Caption         =   "Laporan Penjualan Bulanan"
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
      TabIndex        =   7
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Export Text"
      Height          =   390
      Left            =   4155
      TabIndex        =   6
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblLayar 
      Caption         =   "Cetak &Layar"
      Height          =   390
      Left            =   6045
      TabIndex        =   1
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      Caption         =   "Cetak &Printer"
      Height          =   390
      Left            =   7935
      TabIndex        =   2
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      TabIndex        =   3
      Top             =   7065
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1860
      TabIndex        =   0
      Top             =   135
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-9999"
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
      FormatString    =   "Laporan Penjualan Bulanan"
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
      Left            =   5475
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   600
      TabIndex        =   8
      Top             =   3195
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmLapPenjServiceBln.frx":0000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(MM-YYYY)"
      Height          =   255
      Left            =   2940
      TabIndex        =   5
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan - Tahun"
      Height          =   255
      Left            =   615
      TabIndex        =   4
      Top             =   165
      Width           =   1125
   End
End
Attribute VB_Name = "frmLapPenjServiceBln"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HapusForm()
    Text(0) = ">>->>>>"
End Sub

Private Sub Form_Load()
    Call HapusForm
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptPenjServiceBln
    Dim TglAk As String
    Dim Bulan As Byte
    Dim Tahun As Integer
    Dim TglBlnThnAw As String
    Dim TglBlnThnAk As String
    Dim vNamaTabel As String
    
    If Text(0) = "" Then
        MsgBox "Masukkan Bulan-Tahun !"
    Else
        Bulan = Val(Left(Text(0), 2))
        Tahun = Val(Right(Text(0), 4))
        TglAk = Left(GetTglAkhir(Bulan, Tahun, 1), 2)
        
        gTglTemp = "01 " & GetNamaBulan(Bulan) & " " & Str(Tahun) & " s/d " & _
                   TglAk & " " & GetNamaBulan(Bulan) & " " & Str(Tahun)
        
        TglBlnThnAw = "01-" & Format(Bulan, "00") & " " & Str(Tahun)
        TglBlnThnAk = TglAk & "-" & Format(Bulan, "00") & " " & Str(Tahun)
        
        vNamaTabel = "PenjServiceBln" & DriveSerialNumber("c:")
        Call BuatTabelInv(vNamaTabel)
        
        sSQL = "insert into [" & vNamaTabel & "] select * from T_INVOICE " _
            & " where INVDATE >= '" & TglYYYYMMDD(TglBlnThnAw) & "' AND INVDATE <= '" & TglYYYYMMDD(TglBlnThnAk) & "' " _
            & " union all " _
            & " select " & NmFieldRetur & " from T_INVOICE where " _
            & " convert(char(10),FD_RETUR,120) >='" & TglYYYYMMDD(TglBlnThnAw) & "' " _
            & " AND convert(char(10),FD_RETUR,120) <='" & TglYYYYMMDD(TglBlnThnAk) & "' "
        oConn.Execute sSQL
        
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        sSQL = "select invdate, COUNT(NOINV) AS T_INV, SUM(PART) AS T_PART, SUM(MTRL) AS T_MTRL, " _
             & " SUM(OIL) AS T_OIL, SUM(MET) AS T_MET, SUM(SBLT) AS T_SBLT, SUM(LABOR) AS T_LABOR, " _
             & " SUM(TOTAL-MET) AS T_SUBTTL, SUM(TOTAL) - SUM(FN_PAJAK+MET) AS T_NETTO, " _
             & " SUM(TOTAL) AS T_TOTAL, SUM(FN_PAJAK) AS T_FN_PAJAK, " _
             & " (select SUM(total) from [" & vNamaTabel & "] where INVDATE=X.INVDATE and fs_kredit='' AND STATUS='EXT' ) as T_KONTAN, " _
             & " (select SUM(total) from [" & vNamaTabel & "] where INVDATE=X.INVDATE and fs_kredit='1' AND STATUS='EXT' ) as T_KREDIT, " _
             & " (select SUM(RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7+RTE8) from [" & vNamaTabel & "] where INVDATE=X.INVDATE) as T_RTE " _
             & " from [" & vNamaTabel & "] X where INVDATE >= '" & TglYYYYMMDD(TglBlnThnAw) & "' AND INVDATE <= '" & TglYYYYMMDD(TglBlnThnAk) & "'  " _
             & " GROUP BY INVDATE order by INVDATE"

        rpt.DataControl1.ConnectionString = gReportConn
        rpt.DataControl1.Source = sSQL
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
'            If Val(Left(Text(Index), 2)) > 12 Or Text(Index) = ">>->>>>" Then
'                Text(Index) = Format(gTanggal, "MM-YYYY")
'            End If
            If CekTgl("01-" & Text(Index)) = False Then
                Text(Index) = Format(gTanggal, "MM-YYYY")
            End If
    End Select
End Sub

