VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfoBatalInv 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Informasi Retur Invoice"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      BackColor       =   &H00BFE8EA&
      Caption         =   "Berdasar Tanggal Retur"
      Height          =   255
      Left            =   870
      TabIndex        =   13
      Top             =   480
      Width           =   2595
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00BFE8EA&
      Caption         =   "Berdasar Tanggal Invoice"
      Height          =   285
      Left            =   870
      TabIndex        =   12
      Top             =   180
      Width           =   2610
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
      Left            =   4170
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
      Left            =   2280
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
      Left            =   6060
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
      Left            =   7950
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
      Left            =   2220
      TabIndex        =   0
      Top             =   780
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
      Left            =   2220
      TabIndex        =   1
      Top             =   1125
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
      Left            =   5640
      Top             =   855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   765
      TabIndex        =   11
      Top             =   1860
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmInfoBatalInv.frx":0000
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7455
      Left            =   0
      Picture         =   "frmInfoBatalInv.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Akhir"
      Height          =   255
      Left            =   915
      TabIndex        =   8
      Top             =   1155
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3495
      TabIndex        =   7
      Top             =   1155
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3495
      TabIndex        =   6
      Top             =   795
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Awal"
      Height          =   255
      Left            =   900
      TabIndex        =   5
      Top             =   795
      Width           =   1200
   End
End
Attribute VB_Name = "frmInfoBatalInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub hapusForm()
    Text(0) = ">>->>->>>>"
End Sub

Private Sub Command2_Click()
    Call Cetak(2)
End Sub

Private Sub Form_Load()
    Call hapusForm
    Option2.value = True
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptInfoBatalInv
    
    If Text(0) = ">>->>->>>>" And Text(1) = ">>->>->>>>" Then
        MsgBox "Masukkan Tgl Invoice / Tanggal Batal !"
    Else
        gTglTemp = Text(0) & " s/d " & Text(1)
        gKetTemp = ""
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        
        If Option1.value = True Then
            sSQL = "select INVDATE, NOINV, NO_WO, NO_POLISI, CUSTOMER, FS_LOG_DEL, FS_NM_PETUGAS, FD_RETUR, ALASAN_INV, FS_NO_RETUR, FS_KD_RETUR " _
                 & " from T_INVOICE as a WITH(NOLOCK), t_petugas as b WITH(NOLOCK) where b.FS_KD_PETUGAS = a.FS_LOG_DEL " _
                 & " and (a.invdate >= '" & TglYYYYMMDD(Text(0)) & "'" _
                 & " and a.invdate <= '" & TglYYYYMMDD(Text(1)) & "'" _
                 & " ) and FS_NO_RETUR <> '' " _
                 & " order by a.invdate, a.no_WO"
        Else
            sSQL = "select INVDATE, NOINV, NO_WO, NO_POLISI, CUSTOMER, FS_LOG_DEL, FS_NM_PETUGAS, FD_RETUR, ALASAN_INV, FS_NO_RETUR, FS_KD_RETUR " _
                 & " from T_INVOICE as a WITH(NOLOCK), t_petugas as b WITH(NOLOCK) where b.FS_KD_PETUGAS = a.FS_LOG_DEL " _
                 & " and (convert(char(10),FD_RETUR,120) >= '" & TglYYYYMMDD(Text(0)) & "'" _
                 & " and convert(char(10),FD_RETUR,120) <= '" & TglYYYYMMDD(Text(1)) & "'" _
                 & " ) and FS_NO_RETUR <> '' " _
                 & " order by a.FD_RETUR, a.no_WO"
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

