VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfoDayOnJob 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Informasi Day On Job Progress"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton TblConvert 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Simpan Text"
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
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Export Text"
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
      Left            =   4155
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblLayar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Cetak &Layar"
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
      Left            =   6045
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Cetak &Printer"
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
      Left            =   7935
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7065
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
      Left            =   9825
      Style           =   1  'Graphical
      TabIndex        =   3
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5505
      Top             =   2670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   630
      TabIndex        =   8
      Top             =   3225
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmInfoDayOnJob.frx":0000
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7575
      Left            =   0
      Picture         =   "frmInfoDayOnJob.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2835
      TabIndex        =   5
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   180
      Width           =   960
   End
End
Attribute VB_Name = "frmInfoDayOnJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub hapusForm()
    Text(0) = ">>->>->>>>"
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptInfoDayOnJob
    
    If Text(0) = "" Then
        MsgBox "Masukkan Tgl Repair !"
    Else
        gTglTemp = Text(0)
        gKetTemp = ""
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        sSQL = "select MECH, FRONT, NO_WO, NO_POLISI, RPR_COND,DELIV,REP1,MECH as MEKANIK1,REP2,MECH as MEKANIK2,REP3,MECH as MEKANIK3,REP4,MECH as MEKANIK4, " _
             & " REP5,MECH as MEKANIK5,REP6,MECH as MEKANIK6,REP7,MECH as MEKANIK7, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' and DELIV<>'BTL') as WO_KELUAR, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' and DELIV='BTL') as WO_BATAL, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "') as TERIMA, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' and RPR_COND='' and DELIV<>'BTL') as WO_BELUM, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' and RPR_COND<>'' and DELIV<>'BTL') as WO_PROSES " _
             & " from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' order by date,no_polisi"
        rpt.DataControl1.ConnectionString = gReportConn
        rpt.DataControl1.Source = sSQL
        rpt.Printer.PaperSize = 14
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

Private Sub CetakDblKwarto(Layar As Byte)
    Dim rpt2 As New RptInfoDayOnJobDblKwarto
    
    If Text(0) = "" Then
        MsgBox "Masukkan Tgl Repair !"
    Else
        gTglTemp = Text(0)
        gKetTemp = ""
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        sSQL = "select MECH, FRONT, NO_WO, NO_POLISI, RPR_COND,DELIV,REP1,MECH as MEKANIK1,REP2,MECH as MEKANIK2,REP3,MECH as MEKANIK3,REP4,MECH as MEKANIK4, " _
             & " REP5,MECH as MEKANIK5,REP6,MECH as MEKANIK6,REP7,MECH as MEKANIK7, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' and DELIV<>'BTL') as WO_KELUAR, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' and DELIV='BTL') as WO_BATAL, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "') as TERIMA, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' and RPR_COND='' and DELIV<>'BTL') as WO_BELUM, " _
             & " (select count(NO_POLISI) from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' and RPR_COND<>'' and DELIV<>'BTL') as WO_PROSES " _
             & " from jobdta WITH(NOLOCK) where date = '" & TglYYYYMMDD(Text(0)) & "' order by date,no_polisi"
        rpt2.DataControl1.ConnectionString = gReportConn
        rpt2.DataControl1.Source = sSQL
        rpt2.Printer.PaperSize = 39
        rpt2.Printer.Orientation = ddOPortrait
        Select Case Layar
            Case 1
                rpt2.Show
            Case 2
                ARViewer1.ReportSource = rpt2
            Case Else
                rpt2.PrintReport False
                Unload rpt2
        End Select
    End If
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub TblCetak_Click()
    If MsgBox("Menggunakan kertas folio/Double Kwarto ? (Yes/No)", vbYesNo) = vbYes Then
        Call Cetak(0)
    Else
        Call CetakDblKwarto(0)
    End If
End Sub

Private Sub TblConvert_Click()
    Dim Txt As New ActiveReportsTextExport.ARExportText

    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.ShowSave
    
    Txt.FileName = CommonDialog1.FileName
    Txt.Export ARViewer1.Pages
End Sub

Private Sub Command2_Click()
    Call Cetak(2)
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblLayar_Click()
    If MsgBox("Menggunakan kertas folio/Double Kwarto ? (Yes/No)", vbYesNo) = vbYes Then
        Call Cetak(1)
    Else
        Call CetakDblKwarto(1)
    End If
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
    Select Case Index
        Case 0, 1
            If CekTgl(Text(Index)) <> True Then
                Text(Index) = gTanggal
            End If
    End Select
End Sub

