VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form FrmLap_SiCetakfakturJual 
   Caption         =   "Cetak Faktur"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton cmdCetak 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1185
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   270
      Left            =   4560
      TabIndex        =   6
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   476
      BTYPE           =   8
      TX              =   "Printer :"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmLap_SiCetakfakturJual.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1185
   End
   Begin VB.CommandButton CmdSetPrinter 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1395
   End
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9810
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1185
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   5805
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   6240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmLap_SiCetakfakturJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As Report
Dim Report As Object

Private Sub cmdCetak_Click()
Dim Psize As Integer
Dim POr As Integer

Psize = Report.PaperSize
POr = Report.PaperOrientation

For Each pr In Printers
    If UCase(pr.DeviceName) = UCase(cboPrinter) Then
        Report.SelectPrinter pr.DriverName, pr.DeviceName, pr.Port
        GoTo Cetak
    End If
Next pr
MsgBox "Printer tidak ditemukan Pilih Printer dahulu", vbInformation, "Proses Cetak Gagal"
cboPrinter.SetFocus
Exit Sub

Cetak:
'Report.PrinterSetup Me.hwnd
'Set Printer = Printers(cboPrinter.ListIndex)
'Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
Report.PaperOrientation = POr
Report.PaperSize = Psize
Report.PrintOut
End Sub

Private Sub CmdSetPrinter_Click()
  Report.PrinterSetup Me.hWnd
  
  CRViewer1.Refresh
  
  
  SaveSetting App.Title, "Startup", "PaperSize", Report.PaperSize
  
End Sub

Private Sub CmdKeluar_Click()
   Unload Me
End Sub

Private Sub Command3_Click()
      'J.Rev20160310
1     On Error GoTo Salah
      Dim nmFile As String
      Dim NamaInfo As String
2     NamaInfo = Me.Caption
3     CDialog.filter = "*.doc"
4     CDialog.FileName = Trim(NamaInfo)
5     CDialog.CancelError = False
6     CDialog.ShowSave
7     nmFile = CDialog.FileName & ".doc"
8     Report.ExportOptions.DestinationType = crEDTDiskFile
9     Report.ExportOptions.DiskFileName = nmFile
10    Report.ExportOptions.RTFExportAllPages = True
11    Report.ExportOptions.FormatType = crEFTExactRichText
12    Report.Export False
13    Exit Sub
Salah:
14    MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".export "), vbExclamation, "Err Number : " & Erl
15    MsgBox "Data batal export", vbInformation
End Sub

Private Sub CRViewer1_DownloadFinished(ByVal loadingType As CRVIEWERLibCtl.CRLoadingType)
    CmdSetPrinter.SetFocus
End Sub

Private Sub Form_Load()
Dim oRs As Object
Dim SubReport As Report
On Error GoTo Err_Form_Load
          'J.rev20160317
    Dim objPrinter As Printer
    Set objPrinter = Printer

    cboPrinter.Clear

'2         If InStr(UCase(objPrinter.DeviceName), "PDF") <> 0 Then
'3             cboPrinter = ""
'4         Else
'5             cboPrinter = objPrinter.DeviceName
'6         End If
    
    'J.rev0160308
    If Profile.PdfOnly <> 0 Then
        For Each pr In Printers
            'cboPrinter.AddItem pr.DeviceName
            If Win32_PnPEntity(pr.DeviceName) = False Then
                If InStr(UCase(pr.DeviceName), "PDF") = 0 Then
                    cboPrinter.AddItem pr.DeviceName
                End If
            End If
        Next pr
    End If


    If Profile.PdfOnly = 0 Then
        CRViewer1.EnableExportButton = True
        Command3.Visible = True
        cboPrinter.Visible = False
        cmdCetak.Visible = False
        CRViewer1.EnablePrintButton = True
        XPButton1.Visible = False
    ElseIf Profile.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        Command3.Visible = True
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        CRViewer1.EnablePrintButton = False
        XPButton1.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        Command3.Visible = False
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        CRViewer1.EnablePrintButton = False
        XPButton1.Visible = True
    End If


Dim UkuranKertas As Double
FrmMain.Timer3.Enabled = False
   SQL = "SELECT DISTINCT Sihdr.NoFaktur, Sihdr.NoFjDulu, Sihdr.TypeTrn, Armst.NamaLgn, Armst.Alamat1, Armst.Alamat2, Armst.Alamat3, Sihdr.NoPo, Sihdr.NoSo, Sihdr.TglJthtmp, Sihdr.TglFaktur, Sihdr.Kodelgn, Sihdr.NamaKirim, Sihdr.AlamatKirim1, Sihdr.AlamatKirim2, Sihdr.AlamatKirim3, Sihdr.Subtotal, [Subtotal]*[ExtDisPersen] AS ExtDisJlh, Sihdr.Meterai, Sihdr.TypePpn, Sihdr.PPN, Sihdr.OngkosAngkut, Sihdr.GrandTotal, Sihdr.UangMuka, Sihdr.SisaTotal, Left(cast([Sihdr].[Keterangan] as varchar),300) AS Ketr,"
   With frmSI_KriteriaCetakfaktur
        If .Prt1170 = True Then
            Set Report = New Report_SiCetakFakturJual1
        Else
            Set Report = New Report_SiCetakFakturJual
        End If
        If .CboTypeFaktur.value = "A" Then
           SQL = SQL + " case when Sihdr.NoSO='' then ' P A R T S  INVOICE' ELSE ' S U P P L Y   SLIP ' END AS KetFaktur,"
        Else
           SQL = SQL + " ' P A R T S  RETUR ' AS KetFaktur,"
        End If
        SQL = SQL + " Sisal.Nomor, Sisal.KodeItem, Inmst.NoRack, Sisal.Keterangan, Sisal.Qty, Sisal.SatuanNs, Sisal.HargaJual, Sisal.ItemDisPsn, "
        If .ChkDisc.value = Unchecked Then
'Edit Tgl 3 Nov 2008 di Majapahit
'           Sql = Sql + "Sisal.Qty*[HargaJual] AS Jumlah"
'        Else
'           Sql = Sql + "abs(Sisal.Qty*[HargaJual])-(ItemDisc) AS Jumlah"
'        End If
'--------------------------------
           SQL = SQL + "(abs(Sisal.Qty*[HargaJual]))*(Sisal.Qty/abs(Sisal.Qty)) AS Jumlah"
        Else
           SQL = SQL + "(abs(Sisal.Qty*[HargaJual])-(ItemDisc))*(Sisal.Qty/abs(Sisal.Qty)) AS Jumlah"
        End If
        
        SQL = SQL + " FROM ((Sihdr INNER JOIN Armst ON Sihdr.KodeLgn = Armst.KodeLgn) INNER JOIN Sisal ON Sihdr.NoFaktur = Sisal.NoFaktur) INNER JOIN Inmst ON Sisal.KodeItem = Inmst.KodeItem"
        SQL = SQL + " WHERE (Sihdr.NoFaktur Between '" & .CboFakturAwal & "' And '" & .cboFakturAkhir & "')"
        If .CboTypeFaktur.value = "A" Then
            SQL = SQL + " AND (Sihdr.TypeTrn In ('A','B'))"
        Else
            SQL = SQL + " AND (Sihdr.TypeTrn In ('K','L'))"
        End If
        SQL = SQL + " ORDER BY Sihdr.NoFaktur, Sisal.Nomor;"
        'Set oRS = DbCon.Execute(Sql)
        Report.database.SetDataSource DbCon.Execute(SQL)
'        Sql = "SELECT DISTINCT Sihdr.NoFaktur,Sihdr.Grandtotal,Inmst.NoRack, Sisal.KodeItem, Sisal.Keterangan,Sisal.Qty, Sisal.HargaJual,"
'        If .ChkDisc.Value = Unchecked Then
'           Sql = Sql + "[Qty]*[HargaJual] AS Jumlah,"
'        Else
'        MsgBox "g"
'           Sql = Sql + "abs(Qty)*HargaJuals))-(ItemDisc) AS Jumlah,"
'        End If
'
'        Sql = Sql & " Sisal.ItemDisPsn"
'        Sql = Sql & " FROM ((Sihdr INNER JOIN Armst ON Sihdr.KodeLgn = Armst.KodeLgn) "
'        Sql = Sql & " INNER JOIN Sisal ON Sihdr.NoFaktur = Sisal.NoFaktur) INNER JOIN Inmst"
'        Sql = Sql & " ON Sisal.KodeItem = Inmst.KodeItem WHERE (Sihdr.NoFaktur Between '" & .CboFakturAwal
'        Sql = Sql & "' And '" & .CboFakturAkhir & "') AND (Sihdr.TypeTrn In ('A','B')) ORDER BY Sisal.KodeItem"
'        MsgBox Sql
        
        'J.Rev20180417
        If Left(.CboFakturAwal, 1) = "D" Then
            Report.txtKodeSOP.SetText "FM-NRM-SVC-01.03"
        ElseIf Left(.CboFakturAwal, 1) = "F" Or Left(.CboFakturAwal, 1) = "H" Then
            Report.txtKodeSOP.SetText "FM-NRM-SVC-02.03"
        Else
            Report.txtKodeSOP.SetText ""
        End If

        If .ChkNmPerusahaan.value = Checked Then
            Report.TxtNamaPT.SetText Profile.gNamaProfil
            Report.TxtAlamatPT.SetText Profile.AlamatProfil1
            Report.TxtKotaPT.SetText Profile.KotaProfil
            Report.TxtTelpPT.SetText ": " & Profile.TeleponProfil
            Report.TxtFax.SetText ": " & Profile.FaxProfil
            Report.TxtNPWP.SetText ": " & Profile.npwp
            Report.TxtPKP.SetText ": " & Profile.SKPengukuhan
            Report.TxtTglPengukuhan.SetText ": " & Profile.TglPengukuhan
        Else
            Report.Text4.SetText "": Report.Text5.SetText ""
            Report.Text6.SetText "": Report.Text7.SetText "": Report.Text8.SetText ""
        End If
        Report.FormulaFields(2).text = "" & IIf(.ChkDisc.value = Checked, 0, 1) & ""
End With
UkuranKertas = Val(GetSetting(App.Title, "Startup", "PaperSize", 1))
If UkuranKertas = 0 Then UkuranKertas = 1
Report.PaperSize = UkuranKertas
Screen.MousePointer = vbHourglass
With CRViewer1
     .ReportSource = Report
     .ViewReport
     .Zoom 100
End With
Screen.MousePointer = vbDefault

Set oRs = Nothing
Exit Sub
Err_Form_Load:
MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".Err_Form_Load"), vbInformation
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = -70
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

