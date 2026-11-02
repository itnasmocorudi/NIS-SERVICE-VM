VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLapBulanPasifIRC5000 
   Caption         =   "Laporan Customer Pasif IRC 5000"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   8535
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&Export Text"
      Height          =   390
      Left            =   4155
      TabIndex        =   10
      Top             =   7080
      Width           =   1890
   End
   Begin VB.CommandButton TblConvert 
      Caption         =   "&Simpan Text"
      Height          =   390
      Left            =   2265
      TabIndex        =   9
      Top             =   7080
      Width           =   1890
   End
   Begin VB.CommandButton TblLayar 
      Caption         =   "Cetak &Layar"
      Height          =   390
      Left            =   6045
      TabIndex        =   2
      Top             =   7080
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      Caption         =   "Cetak &Printer"
      Height          =   390
      Left            =   7935
      TabIndex        =   3
      Top             =   7080
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      TabIndex        =   4
      Top             =   7080
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   2205
      TabIndex        =   0
      Top             =   195
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
      FormatString    =   "Laporan Customer Pasif IRC 5000"
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
      Left            =   3195
      TabIndex        =   1
      Top             =   180
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5475
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   600
      TabIndex        =   11
      Top             =   3180
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmLapBulanPasifIRC5000.frx":0000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   240
      Left            =   615
      TabIndex        =   8
      Top             =   690
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hari"
      Height          =   210
      Left            =   3885
      TabIndex        =   7
      Top             =   240
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   255
      Left            =   2790
      TabIndex        =   6
      Top             =   225
      Width           =   285
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tidak Service antara :"
      Height          =   255
      Left            =   -30
      TabIndex        =   5
      Top             =   225
      Width           =   2190
   End
End
Attribute VB_Name = "frmLapBulanPasifIRC5000"
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
    Call HapusForm
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptLapBulanPasifIRC5000
    Dim Kondisi As String
    Dim vNamaTabel As String
    
    gTglTemp = Text(0) & " s/d " & Text(1)
    gKetTemp = ""
    gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
    
    Kondisi = " and dateadd(day," & Val(Text(0)) & ",INVDATE) <= '" & TglYYYYMMDD(gTanggal) & "'  and dateadd(day," & Val(Text(1)) & ",INVDATE) >= '" & TglYYYYMMDD(gTanggal) & "' " _
            & " and KLPKERJA1 +KLPKERJA2 +KLPKERJA3 +KLPKERJA4 +KLPKERJA5 +KLPKERJA6 +KLPKERJA7 like '%IRC%' " _
            & " and FS_NO_RETUR = '' "
            
    vNamaTabel = "Bulan_IRC5000Pasif" & DriveSerialNumber("c:")
    
    Call BuatTabelInv(vNamaTabel)
    sSQL = "insert into [" & vNamaTabel & "] select * from T_INVOICE " _
         & " where dateadd(day," & Val(Text(0)) & ",INVDATE) <= '" & TglYYYYMMDD(gTanggal) & "' and dateadd(day," & Val(Text(1)) & ",INVDATE) >= '" & TglYYYYMMDD(gTanggal) & "' " _
         & " and FS_NO_RETUR = ''"
    oConn.Execute sSQL
    
    sSQL = "select distinct a.NO_POLISI as NO_POLISI, a.CUSTOMER as CUSTOMER, a.CUST_ADRES as CUST_ADRES, a.KOTA as KOTA, a.TELPON as TELPON, a.MDL as MDL, " _
         & " a.FMDL as FMDL, a.DATE as DATE, a.INVDATE as INVDATE, a.NO_WO as NO_WO, a.NOINV as NOINV, a.FS_TH_MOBIL as FS_TH_MOBIL," _
         & " (select KLPKERJA1 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KLP1," _
         & " (select KLPKERJA2 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KLP2," _
         & " (select KLPKERJA3 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KLP3," _
         & " (select KLPKERJA4 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KLP4," _
         & " (select KLPKERJA5 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KLP5," _
         & " (select KLPKERJA6 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KLP6," _
         & " (select KLPKERJA7 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KLP7," _
         & " (select REP1 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KET1," _
         & " (select REP2 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KET2," _
         & " (select REP3 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KET3," _
         & " (select REP4 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KET4," _
         & " (select REP5 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KET5," _
         & " (select REP6 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KET6," _
         & " (select REP7 from [" & vNamaTabel & "] where FS_RC='R005' and NO_WO=a.NO_WO " & Kondisi & " ) as KET7" _
         & " from [" & vNamaTabel & "] as a, CUSTDAT as b where FS_RC='R005' " & Kondisi & " order by NO_POLISI"
    
    
'         & " and dateadd(day," & Val(Text(0)) & ",DATE) <= '" & TglYYYYMMDD(gTanggal) & "'  and dateadd(day," & Val(Text(1)) & ",DATE) >= '" & TglYYYYMMDD(gTanggal) & "' " _
         & " and dateadd(day," & Val(Text(0)) & ",RCDATE) <= '" & TglYYYYMMDD(gTanggal) & "'  and dateadd(day," & Val(Text(1)) & ",RCDATE) >= '" & TglYYYYMMDD(gTanggal) & "' " _
         & " order by NO_POLISI"
    
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

