VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLapBulanPenjualan 
   Caption         =   "Laporan Bulanan Penjualan"
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
      Left            =   4155
      TabIndex        =   12
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblConvert 
      Caption         =   "&Simpan Text"
      Height          =   390
      Left            =   2265
      TabIndex        =   11
      Top             =   7065
      Width           =   1890
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmLapBulanPenjualan.frx":0000
      Left            =   1560
      List            =   "frmLapBulanPenjualan.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   810
      Width           =   735
   End
   Begin VB.CommandButton TblLayar 
      Caption         =   "Cetak &Layar"
      Height          =   390
      Left            =   6045
      TabIndex        =   3
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      Caption         =   "Cetak &Printer"
      Height          =   390
      Left            =   7935
      TabIndex        =   4
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      TabIndex        =   5
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
      FormatString    =   "Laporan Bulanan Penjualan"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5505
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   630
      TabIndex        =   13
      Top             =   3165
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmLapBulanPenjualan.frx":0020
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   885
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   510
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2835
      TabIndex        =   8
      Top             =   510
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2835
      TabIndex        =   7
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   180
      Width           =   960
   End
End
Attribute VB_Name = "frmLapBulanPenjualan"
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
    Dim rpt As New RptLapBulanPenjualan
    Dim vNamaTabel As String
    
    If Not (CekTgl(Text(0)) And CekTgl(Text(1))) And cboStatus.Text = "" Then
        MsgBox "Masukkan Status, Tgl Awal dan Akhir !"
    Else
        gTglTemp = Text(0) & " s/d " & Text(1)
        gKetTemp = "(" + cboStatus.Text + ")"
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        'gReportCon = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";Password='" & NisPass & "';Initial Catalog= " & NisDb & ";Data Source= " & NisKomp & ""
                
        vNamaTabel = "Bulan_Penjualan" & DriveSerialNumber("c:")
        
        Call BuatTabelInv(vNamaTabel)
        
        sSQL = "insert into [" & vNamaTabel & "] select * from T_INVOICE " _
             & " where invdate>='" & TglYYYYMMDD(Text(0)) & "' and invdate<='" & TglYYYYMMDD(Text(1)) & "'" _
             & " union all " _
             & " select " & NmFieldRetur & " from T_INVOICE where convert(char(10),FD_RETUR,120) >='" & TglYYYYMMDD(Text(0)) & "'" _
             & " and convert(char(10),FD_RETUR,120) <='" & TglYYYYMMDD(Text(1)) & "'"
        oConn.Execute sSQL
        
        Select Case cboStatus.ListIndex
            Case 0
                sSQL = "select INVDATE, (select count(NOINV) from [" & vNamaTabel & "] where invdate=x.invdate ) " _
                     & " as JML_INV, sum(RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7) as RATE," _
                     & " sum(LABOR) as LABOR, sum(SBLT) as SBLT, sum(PART) as PART, sum(OIL) as OIL, sum(MTRL) as MTRL," _
                     & " sum(TOTAL-MET) as SUB_TTL, sum(MET) as MET, sum(TOTAL) as GTOTAL, sum(TOTAL-MET-FN_PAJAK) as NETTO, sum(FN_PAJAK) as PAJAK," _
                     & " (select sum(TOTAL) from [" & vNamaTabel & "] where invdate=x.invdate and FS_KREDIT<>'1' ) as CASH," _
                     & " (select sum(TOTAL) from [" & vNamaTabel & "] where invdate=x.invdate and FS_KREDIT='1') as KREDIT" _
                     & " from [" & vNamaTabel & "] x where invdate>='" & TglYYYYMMDD(Text(0)) & "' and invdate<='" & TglYYYYMMDD(Text(1)) & "'" _
                     & " group by invdate order by invdate"
            Case Else
                sSQL = "select INVDATE, (select count(NOINV) from [" & vNamaTabel & "] where invdate=x.invdate  and STATUS='" & cboStatus.Text & "') " _
                     & " as JML_INV, sum(RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7) as RATE," _
                     & " sum(LABOR) as LABOR, sum(SBLT) as SBLT, sum(PART) as PART, sum(OIL) as OIL, sum(MTRL) as MTRL," _
                     & " sum(TOTAL-MET) as SUB_TTL, sum(MET) as MET, sum(TOTAL) as GTOTAL, sum(TOTAL-MET-FN_PAJAK) as NETTO, sum(FN_PAJAK) as PAJAK," _
                     & " (select sum(TOTAL) from [" & vNamaTabel & "] where invdate=x.invdate and FS_KREDIT<>'1'  and STATUS='" & cboStatus.Text & "') as CASH," _
                     & " (select sum(TOTAL) from [" & vNamaTabel & "] where invdate=x.invdate and FS_KREDIT='1' and STATUS='" & cboStatus.Text & "') as KREDIT" _
                     & " from [" & vNamaTabel & "] x where invdate>='" & TglYYYYMMDD(Text(0)) & "' and invdate<='" & TglYYYYMMDD(Text(1)) & "'" _
                     & " and STATUS='" & cboStatus.Text & "' " _
                     & " group by invdate order by invdate"
        End Select
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
        Case 0, 1
            If CekTgl(Text(Index)) <> True Then
                Text(Index) = gTanggal
            End If
            If TglYYYYMMDD(Text(0)) > TglYYYYMMDD(Text(1)) Then
                MsgBox "Tanggal awal lebih besar dari tanggal akhir !"
            End If
    End Select
End Sub

