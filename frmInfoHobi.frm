VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfoHobi 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Informasi Hobi Customer"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12525
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   12525
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton TblConvert 
      Caption         =   "&Simpan Text"
      Height          =   390
      Left            =   2280
      TabIndex        =   14
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Export Text"
      Height          =   390
      Left            =   4170
      TabIndex        =   13
      Top             =   7065
      Width           =   1890
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00BFE8EA&
      Caption         =   "Tanggal service akhir :"
      Height          =   255
      Left            =   780
      TabIndex        =   6
      Top             =   1095
      Width           =   2865
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00BFE8EA&
      Caption         =   "Hobi :"
      Height          =   255
      Left            =   780
      TabIndex        =   5
      Top             =   315
      Width           =   2880
   End
   Begin VB.CommandButton TblLayar 
      Caption         =   "Cetak &Layar"
      Height          =   390
      Left            =   6060
      TabIndex        =   1
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      Caption         =   "Cetak &Printer"
      Height          =   390
      Left            =   7950
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
      Left            =   1020
      TabIndex        =   0
      Top             =   630
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   3300
      TabIndex        =   4
      Top             =   645
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
      Left            =   2460
      TabIndex        =   7
      Top             =   1380
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
      Index           =   2
      Left            =   2460
      TabIndex        =   8
      Top             =   1695
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
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   630
      TabIndex        =   15
      Top             =   3180
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmInfoHobi.frx":0000
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7695
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Awal"
      Height          =   255
      Left            =   660
      TabIndex        =   12
      Top             =   1410
      Width           =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3735
      TabIndex        =   11
      Top             =   1410
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3735
      TabIndex        =   10
      Top             =   1725
      Width           =   1425
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d Tanggal Akhir"
      Height          =   255
      Left            =   660
      TabIndex        =   9
      Top             =   1725
      Width           =   1680
   End
End
Attribute VB_Name = "frmInfoHobi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub hapusForm()
    Text(0) = ""
End Sub

Private Sub Form_Load()
    Call hapusForm
    Option1 = True
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptInfoHobi
    
    If (Text(1) = ">>->>->>>>" Or Text(2) = ">>->>->>>>") And Option2.value = True Then
        MsgBox "Masukkan tgl service akhir !"
    Else
        gTglTemp = Text(0)
        gKetTemp = ""
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        
        If Option1.value = True Then
            sSQL = "select CUSTOMER, CUST_ADRES, KOTA, TELPON, MDL, MODEL, FMDL, TGL_REP, FS_NM_HOBI from custdat WITH(NOLOCK) where FS_NM_HOBI='" & Trim(Text(0)) & "' ORDER BY CUSTOMER"
    
            gKetTemp = "HOBI : " & Text(0)
        Else
            sSQL = "select CUSTOMER, CUST_ADRES, KOTA, TELPON, MDL, MODEL, FMDL, TGL_REP, FS_NM_HOBI from custdat WITH(NOLOCK) where TGL_REP >= '" & TglYYYYMMDD(Text(1)) & "' and TGL_REP <= '" & TglYYYYMMDD(Text(2)) & "' ORDER BY CUSTOMER"
            gTglTemp = "TANGGAL : " & TglDDMMYYYY(Text(1)) & " s/d " & TglDDMMYYYY(Text(2))
        End If
        rpt.DataControl1.ConnectionString = gReportConn
        rpt.DataControl1.Source = sSQL
        rpt.Printer.PaperSize = 14
        rpt.Printer.Orientation = ddOLandscape
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

Private Sub Option1_Click()
    Call Pilihan1
End Sub

Private Sub Pilihan2()
    Text(0).Enabled = False
    Text(1).Enabled = True
    Text(2).Enabled = True
    Label16.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
End Sub

Private Sub Pilihan1()
    Text(0).Enabled = True
    Text(1).Enabled = False
    Text(2).Enabled = False
    Label16.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
End Sub

Private Sub Option2_Click()
    Call Pilihan2
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
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "HOBI"
                frmCari.Show (vbModal)
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
End Sub


Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                sSQL = "Select * from T_HOBI WITH(NOLOCK) where FS_NM_HOBI ='" & Text(0) & "'"
                Set rsSearch = objO.FetchData(sSQL)
                If rsSearch.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "HOBI"
                    frmCari.Show (vbModal)
                End If
            End If
        Case 1, 2
            If CekTgl(Text(Index)) <> True Then
                Text(Index) = gTanggal
            End If
    End Select
End Sub
