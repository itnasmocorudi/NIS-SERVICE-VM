VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLapPenjService 
   Caption         =   "Laporan Penjualan Harian Service"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   10515
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
   Begin VB.ComboBox CboStatus 
      Height          =   315
      ItemData        =   "frmLapPenjService.frx":0000
      Left            =   1545
      List            =   "frmLapPenjService.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   450
      Width           =   825
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
      Left            =   1545
      TabIndex        =   0
      Top             =   120
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
      FormatString    =   "Laporan Penjualan Harian Service"
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
      Left            =   5490
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   615
      TabIndex        =   10
      Top             =   3195
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmLapPenjService.frx":0018
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2820
      TabIndex        =   7
      Top             =   150
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   150
      Width           =   945
   End
End
Attribute VB_Name = "frmLapPenjService"
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
    Dim rpt As New RptPenjService
    Dim vNamaTabel As String
    
    If cboStatus.Text = "" Or Text(0) = "" Then
        MsgBox "Masukkan Tgl Repair dan Status !"
    Else
        gTglTemp = Text(0)
        Select Case cboStatus.ListIndex
            Case 0, 1
                vNamaTabel = "PenjService" & DriveSerialNumber("c:")
                Call BuatTabelInv(vNamaTabel)
                
                sSQL = "insert into [" & vNamaTabel & "] select * from T_INVOICE " _
                    & " where STATUS='" & Trim(cboStatus.Text) & "' " _
                    & " and INVDATE='" & TglYYYYMMDD(Text(0)) & "' " _
                    & " union all " _
                    & " select " & NmFieldRetur & " from T_INVOICE where convert(char(10),FD_RETUR,120) ='" & TglYYYYMMDD(Text(0)) & "' " _
                    & " and STATUS='" & cboStatus.Text & "'"
                
                oConn.Execute sSQL
                
                gKetTemp = cboStatus.Text
                gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
                sSQL = "select NOINV, NO_POLISI, PART, MTRL, OIL, MET, SBLT, LABOR, BODY, FN_PAJAK, " _
                     & " (select total from [" & vNamaTabel & "] where noinv=x.noinv and fs_kredit='') as KONTAN, " _
                     & " (select total from [" & vNamaTabel & "] where noinv=x.noinv and fs_kredit='1') as KREDIT, " _
                     & " (select count(NOINV) from [" & vNamaTabel & "] where len(NOINV) = 9 and " _
                     & " STATUS='" & Trim(cboStatus.Text) & "' and INVDATE='" & TglYYYYMMDD(Text(0)) & "') as TOT_RET, " _
                     & " (select count(NOINV) from [" & vNamaTabel & "] where len(NOINV) = 8 and " _
                     & " STATUS='" & Trim(cboStatus.Text) & "' and INVDATE='" & TglYYYYMMDD(Text(0)) & "') as TOT_INV, " _
                     & " (select sum(TOTAL) from [" & vNamaTabel & "] where " _
                     & " STATUS='" & Trim(cboStatus.Text) & "' and INVDATE='" & TglYYYYMMDD(Text(0)) & "') as TOTAL " _
                     & " from [" & vNamaTabel & "] x where STATUS='" & Trim(cboStatus.Text) & "' " _
                     & " and INVDATE='" & TglYYYYMMDD(Text(0)) & "' " _
                     & "" _
                     & " order by FS_NO_RETUR,NOINV"
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
            If CekTgl(Text(Index)) <> True Then
                Text(Index) = gTanggal
            End If
    End Select
End Sub

