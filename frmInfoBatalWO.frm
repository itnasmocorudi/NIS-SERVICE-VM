VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfoBatalWO 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Informasi Batal Order"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmInfoBatalWO.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option1 
      BackColor       =   &H00BFE8EA&
      Caption         =   "Berdasar Tanggal Order"
      Height          =   285
      Left            =   585
      TabIndex        =   13
      Top             =   240
      Width           =   2610
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00BFE8EA&
      Caption         =   "Berdasar Tanggal Pembatalan"
      Height          =   255
      Left            =   585
      TabIndex        =   12
      Top             =   555
      Width           =   2595
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
      TabIndex        =   10
      Top             =   7065
      Width           =   1890
   End
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
      TabIndex        =   9
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
      TabIndex        =   2
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
      TabIndex        =   3
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
      TabIndex        =   4
      Top             =   7065
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1965
      TabIndex        =   0
      Top             =   855
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
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   1965
      TabIndex        =   1
      Top             =   1230
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
      Left            =   5010
      Top             =   795
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   480
      TabIndex        =   11
      Top             =   1785
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmInfoBatalWO.frx":0342
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7575
      Left            =   0
      Picture         =   "frmInfoBatalWO.frx":037A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Akhir"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1260
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1260
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   885
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Awal"
      Height          =   255
      Left            =   495
      TabIndex        =   5
      Top             =   885
      Width           =   1350
   End
End
Attribute VB_Name = "frmInfoBatalWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub hapusForm()
    text(0) = ">>->>->>>>"
End Sub

Private Sub Command2_Click()
    Call Cetak(2)
End Sub

Private Sub Form_Load()
    Call hapusForm
    Option1.value = True
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptInfoBatalWO
    
    If text(0) = "" Then
        MsgBox "Masukkan Tgl WO / Tgl Batal !"
    Else
        gTglTemp = text(0) & " s/d " & text(1)
        gKetTemp = ""
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        
        If Option1.value = True Then
            sSQL = "select DATE, NO_WO, NO_POLISI, CUSTOMER, FS_LOG_DEL, FS_NM_PETUGAS, FD_TGL_DEL, ALASAN, DELIV " _
                 & " from jobdta as a WITH(NOLOCK), t_petugas as b WITH(NOLOCK) where b.FS_KD_PETUGAS = a.FS_LOG_DEL " _
                 & " and (a.date >= '" & TglYYYYMMDD(text(0)) & "'" _
                 & " and a.date <= '" & TglYYYYMMDD(text(1)) & "'" _
                 & " ) and DELIV = 'BTL' " _
                 & " order by a.date, a.no_WO"
        Else
            sSQL = "select DATE, NO_WO, NO_POLISI, CUSTOMER, FS_LOG_DEL, FS_NM_PETUGAS, FD_TGL_DEL, ALASAN, DELIV " _
                 & " from jobdta as a WITH(NOLOCK), t_petugas as b WITH(NOLOCK) where b.FS_KD_PETUGAS = a.FS_LOG_DEL " _
                 & " and (a.FD_TGL_DEL >= '" & TglYYYYMMDD(text(0)) & "'" _
                 & " and a.FD_TGL_DEL <= '" & TglYYYYMMDD(text(1)) & "'" _
                 & " ) and DELIV = 'BTL' " _
                 & " order by a.FD_TGL_DEL, a.no_WO"
        End If
        
        rpt.DataControl1.ConnectionString = gReportConn
        rpt.DataControl1.Source = sSQL
        rpt.Printer.Orientation = ddOPortrait
        rpt.Printer.PaperSize = 14
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

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub TblCetak_Click()
    Call Cetak(0)
End Sub

Private Sub TblConvert_Click()
    Dim Txt As New ActiveReportsTextExport.ARExportText

    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.ShowSave
    
    Txt.FileName = CommonDialog1.FileName
    Txt.Export ARViewer1.Pages
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
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
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
            If CekTgl(text(Index)) <> True Then
                text(Index) = gTanggal
            End If
    End Select
End Sub

