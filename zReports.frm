VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form zReports 
   BorderStyle     =   0  'None
   Caption         =   ".:: Reports ::."
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11565
   Icon            =   "zReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicPrint 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   4080
      TabIndex        =   3
      Top             =   255
      Width           =   4080
      Begin VB.ComboBox CboPrint 
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
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   825
      End
   End
   Begin VB.PictureBox PicR 
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   8655
      TabIndex        =   1
      Top             =   120
      Width           =   8655
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   -1380
      TabIndex        =   2
      Top             =   405
      Width           =   20735
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   3555
      Left            =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1530
      Width           =   8715
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
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
   Begin ucXPButton.XPButton CmdSeting 
      Height          =   495
      Left            =   1380
      TabIndex        =   6
      Top             =   5625
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&s"
      ENAB            =   -1  'True
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "zReports.frx":29C12
      PICN            =   "zReports.frx":29C2E
      PICH            =   "zReports.frx":2A3BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton CmdCetak1 
      Height          =   495
      Left            =   3585
      TabIndex        =   7
      Top             =   5640
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&p"
      ENAB            =   -1  'True
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "zReports.frx":2ACF5
      PICN            =   "zReports.frx":2AD11
      PICH            =   "zReports.frx":2B469
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton CmdExport 
      Height          =   495
      Left            =   5775
      TabIndex        =   8
      Top             =   5655
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&e"
      ENAB            =   -1  'True
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "zReports.frx":2BD5D
      PICN            =   "zReports.frx":2BD79
      PICH            =   "zReports.frx":2C4D8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton x 
      Height          =   495
      Left            =   7995
      TabIndex        =   9
      Top             =   5625
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&c"
      ENAB            =   -1  'True
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "zReports.frx":2CD87
      PICN            =   "zReports.frx":2CDA3
      PICH            =   "zReports.frx":2D5CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "zReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Report      As Object
Private isLoad      As String
Private nNoBok      As String
Private nNoWo       As String
Private nNoEst      As String
Private nCust       As Boolean
Private nInv        As String
Private nMeter      As Integer
Private nCroz(295)  As Integer
Private nTipe       As String
Private nSup        As String
Private nGroup      As String
Private nGroupKet   As String
Private nJob        As String
Private nSA         As String
Private nRetur      As Boolean
Private nLapBul     As Boolean
Private nRekap      As Boolean
Private nTgl0       As Date
Private nTgl1       As Date
Private nNoRangka   As String
Private oRS         As ADODB.Recordset
Private nRS         As ADODB.Recordset
Private data(10)    As String
Private nDt         As Integer
Private i           As Integer
Private rsGoSublet  As New ADODB.Recordset
Private rsGrid      As New ADODB.Recordset
Private txt5        As String
Private txt6        As String
Private nCab        As String
Private Total       As Double
Private Jabatan     As String
Private Pejabat     As String
Private TipeFP      As Boolean
Private noFP        As String
Private nAsr        As Integer
Private nInc        As Boolean
Private rsDc        As ADODB.Recordset
Private nFilter     As String
Private npersn      As Double
'----------------------------------------
Private WmiObjSet   As Variant
Private Wmiobj      As Variant
Private IndexPrint  As Integer
Private Psize       As Integer
Private POr         As Integer
'--------------------------------------------------------------
'--------------------------------------------------------------
Public Property Let Load_Report(value As String)
    isLoad = value
End Property
Public Property Let Load_NoWo(value As String)
    nNoWo = value
End Property
Public Property Let Load_NoInvoice(value As String)
    nInv = value
End Property
Public Property Let Load_NoEst(value As String)
    nNoEst = value
End Property
Public Property Let Load_IncWOBatal(value As Boolean)
    nInc = value
End Property
Public Property Let Load_NoBok(value As String)
    nNoBok = value
End Property
Public Property Let Load_Cust(value As Boolean)
    nCust = value
End Property
Public Property Let Load_Retur(value As Boolean)
    nRetur = value
End Property
Public Property Let Load_LapBul(value As Boolean)
    nLapBul = value
End Property
Public Property Let Load_Rekap(value As Boolean)
    nRekap = value
End Property
Public Property Let Load_SA(value As String)
    nSA = value
End Property
Public Property Let Load_Meter(value As Integer)
    nMeter = value
    i = 0
End Property
Public Property Let Load_Croz(value As Integer)
    nCroz(i) = value
    i = i + 1
End Property
Public Property Let Load_Tipe(value As String)
    nTipe = value
End Property
Public Property Let Load_Cabang(value As String)
    nCab = value
End Property
Public Property Let Load_Suplier(value As String)
    nSup = value
End Property
Public Property Let Load_Group(value As String)
    nGroup = value
End Property
Public Property Let Load_Job(value As String)
    nJob = value
End Property
Public Property Let Load_GroupKet(value As String)
    nGroupKet = value
End Property
Public Property Let Load_Tgl0(value As Date)
    nTgl0 = value
End Property
Public Property Let Load_Tgl1(value As Date)
    nTgl1 = value
End Property
Public Property Let Load_NoRangka(value As String)
    nNoRangka = value
End Property
Public Property Let Load_Recordset(value As ADODB.Recordset)
    Set nRS = value
End Property
Public Property Let Load_Id_Data(value As Integer)
    nDt = value
End Property
Public Property Let Load_Data(value As String)
    data(nDt) = value
End Property
Public Property Let Val5(value As String)
    txt5 = value
End Property
Public Property Let Val6(value As String)
    txt6 = value
End Property
Public Property Let Load_Total(value As Double)
    Total = value
End Property
Public Property Let NamaJabatan(value As String)
    Jabatan = value
End Property
Public Property Let NamaPejabat(value As String)
    Pejabat = value
End Property
Public Property Let TipeFakturPajak(value As Boolean)
    TipeFP = value
End Property
Public Property Let NomorFP(value As String)
    noFP = value
End Property
Public Property Let Load_Asuransi(value As Integer)
    nAsr = value
End Property
Public Property Let Load_Filter(value As String)
    nFilter = value
End Property
Public Property Let Load_Persn(value As Double)
    npersn = value
End Property

Private Sub CboPrint_Change()
    Dim q As Integer
    Psize = 0
    POr = 0
    IndexPrint = 0
    
    Psize = Report.PaperSize
    POr = Report.PaperOrientation

    Set Printer = Printers(CboPrint.ListIndex)
    If UCase(Printer.DeviceName) <> UCase(CboPrint.Text) Then
        GoSub SettingPrinter
    Else
        Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    End If
    Exit Sub
SettingPrinter:
    For q = 0 To 50
        Set Printer = Printers(q)
        If UCase(Printer.DeviceName) = UCase(CboPrint.Text) Then
            IndexPrint = q
            Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            On Error GoTo PsizeErr
            Report.PaperSize = Psize
Jx:
            Report.PaperOrientation = POr
            CRViewer.Refresh
            Return
        End If
    Next q
    Return
    
PsizeErr:
    Report.PaperSize = 0
    GoTo Jx
End Sub

Private Sub CboPrint_Click()
    Call CboPrint_Change
End Sub

'--------------------------------------------------------------
'--------------------------------------------------------------
Private Sub CmdKeluar_Click()
    Unload Me
End Sub
Private Sub CmdSeting_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer.Refresh
End Sub
Private Sub Form_Resize()
    CRViewer.Top = 0
    CRViewer.Left = 0
    CRViewer.Width = Me.Width - 200
    CRViewer.Height = Me.Height - 100
    PicR.Left = Me.Width - 8855
    PicPrint.Left = IIf(PicR.Left - 4100 < 15, CRViewer.Left + 900, PicR.Left - 4100)
    PicPrint.Top = CRViewer.Top + 300
    If PicPrint.Left + PicPrint.Width > PicR.Left Then
        PicPrint.Left = CRViewer + 200
        PicPrint.Top = CRViewer.Top + 300
    End If
End Sub
Private Sub CmdCetak_Click()
    Label19.ForeColor = vbBlack
    If onExpo = 1 Or onExpo = 2 Then 'Just Print Physical & Export *.Doc
        If CboPrint.Text = "" Or (UCase(Trim(CboPrint.Text)) <> UCase(Trim(Printer.DeviceName))) Then
            With msg
                .PassingMsg = "Printer belum dipilih!"
                .passingTipe = 3
                .Show vbModal
            End With
            Label19.ForeColor = vbRed
            CboPrint.SetFocus
            Exit Sub
        Else
            Set Printer = Printers(IndexPrint)
            Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            Report.PaperSize = Psize
            Report.PaperOrientation = POr
            CRViewer.Refresh
        End If
    End If
    Report.PrintOut
End Sub
Private Sub cmdExport_Click()
    On Error GoTo Salah
    Dim NamaFile As String
    NamaFile = ""
    If onExpo = 1 Then
        ft.ShowForm = "ExportRestrict"
        ft.Show vbModal
        If ft.onVal = "" Then Exit Sub
        
        If UCase(Right(Trim(ft.onVal), 5)) = ".XLSX" Then NamaFile = Mid(Trim(ft.onVal), 1, Len(Trim(ft.onVal)) - 5)
        If UCase(Right(Trim(ft.onVal), 4)) = ".XLS" Then NamaFile = Mid(Trim(ft.onVal), 1, Len(Trim(ft.onVal)) - 4)
        
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = NamaFile & ".Doc"
        Report.ExportOptions.RTFExportAllPages = True
        Report.ExportOptions.FormatType = crEFTExactRichText
        Report.Export False
        
        With msg
            .PassingMsg = NamaFile & ".Doc" & "  ... Data sudah di Export !"
            .passingTipe = 2
            .Show vbModal
        End With
        Exit Sub
    ElseIf onExpo = 0 Then
        Report.Export
    End If
    Exit Sub
Salah:
    If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
    MsgBox "Err.Line : " & Erl & "Err.Numb : " & err.Number & "Err.Desc. : " & err.Description & vbCr & "Procedure : " & Me.Name & " export"
Keluar:
End Sub
'--------------------------------------------------------------
'--------------------------------------------------------------

Private Sub Form_Load()
    On Error GoTo Vb_Error
    Dim rsVr As ADODB.Recordset
    Set rsDc = objA.Get_Counter
    Move 0, objA.MyTop + 90
    Me.Width = objA.MyLeft + 3240
    Me.Height = fHeight + 40
    Me.Caption = TitleBP & "          .:: " & isLoad & " ::."
    Set rsVr = objA.Get_Versi
    onExpo = CInt(rsVr("OnExport"))
    Select Case onExpo
        Case 0      'Print & Export All (Normal)
            CRViewer.EnablePrintButton = True
            CRViewer.EnableExportButton = True
            cmdExport.Enabled = True
            PicPrint.Visible = False
            CboPrint.Visible = False
            cmdCetak.Visible = True
            CmdSeting.Visible = True
        Case 1      'Print Pyshical & Export *.Doc
            CRViewer.EnablePrintButton = False
            CRViewer.EnableExportButton = False
            cmdExport.Enabled = True
            PicPrint.Visible = True
            CboPrint.Visible = True
            Call GetPrinter
            cmdCetak.Visible = True
            CmdSeting.Visible = True
        Case 2      'Just Print Physical
            CRViewer.EnablePrintButton = False
            CRViewer.EnableExportButton = False
            cmdExport.Enabled = False
            PicPrint.Visible = True
            CboPrint.Visible = True
            Call GetPrinter
            cmdCetak.Visible = True
            CmdSeting.Visible = True
    End Select
    Select Case isLoad
        Case "Gatepass"                             '01
            Call Go_Gatepass
        Case "Gatepass Sementara"                   '02
            Call Go_Gatepass_Sementara
        Case "Gatepass Batal"                       '03
            Call Go_Gatepass_Batal
        Case "Estimasi"                             '04
            Call Go_Estimasi
        Case "Booking"                              '05
            Call Go_Booking
        Case "OwnRisk"                              '06
            Call Go_OwnRisk
        Case "Sublet"                               '07
            Call Go_Sublet
        Case "Order"                                '08
            Call Go_Order
        Case "Order_AddJob"                         '09
            Call Go_Order_addJob
        Case "MaterialGroup"                        '10
            Call Go_MaterialGroup
        Case "MaterialWo"                           '11
            Call Go_MaterialWo
        Case "Material"                             '12
            Call Go_Material
        Case "Progress"                             '13
            Call Go_Progress
        Case "PoSup"                                '14
            Call Go_PoSuplier
        Case "Order10"                              '15
            Call Go_Order_10
        Case "Sublet_Inv"                           '16
            Call Go_Sublet_Invoice
        Case "eRepHist"                             '17
            Call Go_Repair_History
        Case "eJobLight"                            '18
            Call Go_Job_Light
        Case "eJobMediumHeavy"                      '19
            Call Go_Job_MediumHeavy
        Case "eOrder"                               '20
            Call Go_RepOrder
        Case "SetTagih"                             '21
            Call Go_SetTagih
        Case "SetTagih_ulang"                       '22
            Call Go_SetTagih_Ulang
        Case "SetTagihSublet"                       '23
            Call Go_SetTagihSublet
        Case "SetTagihSublet_ulang"                 '24
            Call Go_SetTagihSublet_Ulang
        Case "eEstimasi"                            '25
            Call Go_Info_Estimasi
        Case "eOutstandingSPK"                      '26
            Call Go_OutstandingSPK
        Case "eValidExt"                            '27
            Call Go_ValidasiEXT
        Case "eValidExt_SUM"                        '28
            Call Go_ValidasiEXT_SUM
        Case "Invoice"                              '29
            Call Go_Invoice
        Case "Faktur"                               '30
            Call Go_Faktur
        Case "FakturLamp"                           '31
            Call Go_FakturLamp
        Case "KuitansiOR"                           '32
            Call Go_KuitansiOR
        Case "eBooking"                             '33
            Call Go_RepBooking
        Case "eBatalOrder"                          '34
            Call Go_RepBatalOrder
        Case "eBilling"                             '35
            Call Go_Info_Billing
        Case "eBilling2"                            '36
            Call Go_Info_Billing2
        Case "eBilling3"                            '37
            Call Go_Info_Billing3
        Case "eBilling4"                            '38
            Call Go_Info_Billing4
        Case "PartOrd"                              '39
            Call Go_PartOrder
        Case "eDataAR"                              '40
            Call Go_Info_DataAR
        Case "eOwnRisk"                             '41
            Call Go_Info_OR
        Case "Registrasi"                           '42
            Call Go_Registrasi
        Case "eReEstimasi"                          '43
            Call Go_ReEstimasi
        Case "SetTagih_Group"
            Call Go_SetTagih_Group
    End Select
    Exit Sub
Vb_Error:
  err.Raise err.Number, err.Source, err.Description
End Sub

'====================================================================================
'====================================================================================
'====================================================================================
Private Sub Go_Gatepass()                           '01
    Set Report = New Rpt_sGatepass
    Set oRS = objR.Gatepass(nNoWo)
    Report.database.SetDataSource oRS
    Call SetReport
    Call Set_On_Gatepass
End Sub

Private Sub Go_PartOrder()                           '02
    Set Report = New Rpt_aOrderPart
    Set oRS = objR.OwnRisk_Main(nNoWo)
    Report.database.SetDataSource oRS
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    '-- List Part ------------------------------------------
        Dim crxPr As New CRAXDRT.Report
        Dim prRS As ADODB.Recordset
        Set prRS = objR.PartOrder_List(nNoWo)
        Dim SumPr As Double
        If prRS.RecordCount > 0 Then
            prRS.MoveFirst
            SumPr = 0
            Do Until prRS.EOF
                SumPr = SumPr + Val(prRS.Fields(4).value)
                prRS.MoveNext
            Loop
        End If
        Set crxPr = Report.SubOR.OpenSubreport
        crxPr.database.SetDataSource prRS
    Call SetReport
End Sub

Private Sub Go_Gatepass_Sementara()                 '03
    Set Report = New Rpt_sGatepassSementara
    Set oRS = objR.Gatepass(nNoWo)
    oRS.MoveLast
    Report.database.SetDataSource oRS
    Call SetReport
    Call Set_On_Gatepass
End Sub

Private Sub Go_Gatepass_Batal()                     '04
    Set Report = New Rpt_sGatepassBatal
    Set oRS = objR.Gatepass(nNoWo)
    Report.database.SetDataSource oRS
    Call Set_On_Gatepass
    Call SetReport
End Sub

    Private Sub Set_On_Gatepass()
        Report.txtNas.SetText Profil.nNamaDealer
        Report.txtAL.SetText Profil.nAlamatProfil
        Report.txtKabeng.SetText Profil.nNama_Kabeng
    End Sub

Private Sub Go_Estimasi()                           '05
    If nCust = True Then
        Set Report = New Rpt_cEstimasi_Customer
    Else
        Set Report = New Rpt_cEstimasi_Asuransi
    End If
    Set oRS = objR.Estimasi_Main(nNoEst)
    Report.database.SetDataSource oRS
    Report.txtGrand.SetText Format(oRS.Fields(35).value, "###,###,###.00")
    Report.txtBilang.SetText Angka2Huruf(Val(oRS.Fields(35).value), 1)
    Report.txtSignFax.SetText "*) Please sign and re-fax to " & Profil.nFaxProfil
    Report.txtKota.SetText (Profil.nKotaProfil) & ","
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    Report.txtReg.SetText nJob
    '-- EST_JOB --------------------------------------------
        Dim crxJob As New CRAXDRT.Report
        Dim jobRS As ADODB.Recordset
        Set jobRS = objR.Estimasi_Job(nNoEst)
        Set crxJob = Report.subJob.OpenSubreport
        crxJob.database.SetDataSource jobRS
        crxJob.TxtExJasa.SetText Trim(oRS.Fields(27).value) & " %"
        crxJob.txtTotJasa.SetText Format(oRS.Fields(29).value, "###,###,###.00")
        
    '-- EST_PART -------------------------------------------
        Dim crxPart As New CRAXDRT.Report
        Dim PartRS As ADODB.Recordset
        Set PartRS = objR.Estimasi_Part(nNoEst)
        Set crxPart = Report.subPart.OpenSubreport
        crxPart.database.SetDataSource PartRS
        crxPart.TxtExPart.SetText Trim(oRS.Fields(31).value) & " %"
        crxPart.txtTotPart.SetText Format(oRS.Fields(33).value, "###,###,###.00")
        
    '-- EST_SUBLET -----------------------------------------
        Dim crxSblt As New CRAXDRT.Report
        Dim SbltRS As ADODB.Recordset
        Set SbltRS = objR.Estimasi_Sublet(nNoEst)
        Set crxSblt = Report.subSublet.OpenSubreport
        crxSblt.database.SetDataSource SbltRS
        
    '-------------------------------------------------------
    If rsDc.Fields(24).value = 0 Then
        crxJob.txt1.TextColor = &HFFC0C0
        crxJob.txt2.TextColor = &HFFC0C0
        crxJob.Field1.TextColor = &HFFFFFF
        crxJob.Field13.TextColor = &HFFFFFF
        
        crxJob.Line7.LineColor = &HFFFFFF
        crxJob.Text6.TextColor = &HFFFFFF
        crxJob.TxtExJasa.TextColor = &HFFFFFF
        crxJob.Line3.LineColor = &HFFFFFF
        crxJob.Field2.TextColor = &HFFFFFF
        crxJob.Text7.TextColor = &HFFFFFF
        crxJob.Text9.Top = 20
        crxJob.txtTotJasa.Top = 20
        
        crxPart.Line4.LineColor = &HFFFFFF
        crxPart.Text6.TextColor = &HFFFFFF
        crxPart.TxtExPart.TextColor = &HFFFFFF
        crxPart.Line6.LineColor = &HFFFFFF
        crxPart.Field7.TextColor = &HFFFFFF
        crxPart.Text7.TextColor = &HFFFFFF
        crxPart.Text9.Top = 20
        crxPart.txtTotPart.Top = 20
    End If
    Call SetReport
End Sub

Private Sub Go_Booking()                           '06
    Set Report = New Rpt_cBooking
    Set oRS = objR.Booking_Main(nNoBok)
    Report.database.SetDataSource oRS
    Report.txtBilang.SetText Angka2Huruf(Val(oRS.Fields(46).value), 1)
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    '-- BOOK_JOB --------------------------------------------
        Dim crxJob As New CRAXDRT.Report
        Dim jobRS As ADODB.Recordset
        Set jobRS = objR.Booking_Job(nNoBok)
        Set crxJob = Report.subJob.OpenSubreport
        crxJob.database.SetDataSource jobRS
        crxJob.TxtExJasa.SetText Trim(oRS.Fields(38).value) & " %"
        crxJob.txtTotJasa.SetText Format(oRS.Fields(40).value, "###,###,###.00")
    '-- BOOK_PART -------------------------------------------
        Dim crxPart As New CRAXDRT.Report
        Dim PartRS As ADODB.Recordset
        Set PartRS = objR.Booking_Part(nNoBok)
        Set crxPart = Report.subPart.OpenSubreport
        crxPart.database.SetDataSource PartRS
        crxPart.TxtExPart.SetText Trim(oRS.Fields(42).value) & " %"
        crxPart.txtTotPart.SetText Format(oRS.Fields(44).value, "###,###,###.00")
    '-- BOOK_SUBLET -----------------------------------------
        Dim crxSblt As New CRAXDRT.Report
        Dim SbltRS As ADODB.Recordset
        Set SbltRS = objR.Booking_Sublet(nNoBok)
        Set crxSblt = Report.subSublet.OpenSubreport
        crxSblt.database.SetDataSource SbltRS
    '-------------------------------------------------------
    If rsDc.Fields(24).value = 0 Then
        crxJob.txt1.TextColor = &HFFFFFF
        crxJob.txt2.TextColor = &HFFFFFF
        crxJob.Field4.TextColor = &HFFFFFF
        crxJob.Field5.TextColor = &HFFFFFF
        crxPart.Field6.TextColor = &HFFFFFF
        crxPart.Text6.TextColor = &HFFFFFF
        crxPart.TxtExPart.TextColor = &HFFFFFF
        crxPart.Text16.TextColor = &HFFFFFF
        crxPart.Line4.LineColor = &HFFFFFF
        crxPart.Text17.Top = 15
        crxPart.txtTotPart.Top = 15
    End If
    Call SetReport
End Sub

Private Sub Go_OwnRisk()                           '06
    Set Report = New Rpt_aOwnRisk
    Set oRS = objR.OwnRisk_Main(nNoWo)
    Report.database.SetDataSource oRS
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    Report.txtJabatan.SetText "Section Head"
    '-- List OR ------------------------------------------
        Dim crxOR As New CRAXDRT.Report
        Dim orRS As ADODB.Recordset
        Set orRS = objR.OwnRisk_List(nNoWo)
        Dim SumOr As Double
        If orRS.RecordCount > 0 Then
            orRS.MoveFirst
            SumOr = 0
            Do Until orRS.EOF
                SumOr = SumOr + Val(orRS.Fields(2).value)
                orRS.MoveNext
            Loop
        End If
        Set crxOR = Report.SubOR.OpenSubreport
        crxOR.database.SetDataSource orRS
        crxOR.txtBilang.SetText Angka2Huruf(SumOr, 1)
    Call SetReport
End Sub

Private Sub Go_Sublet()                          '07
    createTbl_Go_Sublet_ALFA True
    Dim rsAlfa As ADODB.Recordset
    Set rsAlfa = objR.Sublet_ALFA(nNoWo)
    If rsAlfa.RecordCount > 0 Then
        Dim nAlfa As String
        C = 0
        rsAlfa.MoveFirst
        Do Until rsAlfa.EOF
            If rsAlfa.Fields(9).value <> "" Then
                If C = 0 Then
                    nAlfa = "'" & rsAlfa.Fields(9).value & "'"
                Else
                    nAlfa = nAlfa & "," & "'" & rsAlfa.Fields(9).value & "'"
                End If
                C = C + 1
            End If
            rsAlfa.MoveNext
        Loop
        nAlfa = "(" & nAlfa & ")"
        Dim rsBeta As ADODB.Recordset
        If nAlfa = "()" Then nAlfa = "('x')"
        Set rsBeta = objR.Sublet_BETA(nAlfa)

        rsAlfa.MoveFirst
        Do Until rsAlfa.EOF
                rsGoSublet.AddNew
                rsGoSublet!NoPO = rsAlfa!NoPO
                rsGoSublet!noWO = rsAlfa!noWO
                rsGoSublet!noRangka = rsAlfa!noRangka
                rsGoSublet!noPol = rsAlfa!noPol
                rsGoSublet!Pekerjaan = rsAlfa!Pekerjaan
                rsGoSublet!qty = Val(rsAlfa!qty)
                rsGoSublet!KodeSatuan = rsAlfa!KodeSatuan
                rsGoSublet!keterangan = rsAlfa!keterangan
                rsGoSublet!HargaJual = Val(rsAlfa!HargaJual)
                If rsAlfa.Fields(9).value = "" Then
                    rsGoSublet!Supplier = ""
                Else
                    Dim nSup As String
                    rsBeta.Find "KodeLgn = '" & rsAlfa.Fields(9).value & "'", , adSearchForward, 1
                    If Not rsBeta.EOF Then
                        rsGoSublet!Supplier = rsBeta.Fields(1).value
                    Else
                        rsGoSublet!Supplier = ""
                    End If
                End If
                isEstm = Profil.nUserAcc & " " & Profil.nPassAcc
                rsGoSublet.Update
            rsAlfa.MoveNext
        Loop
    Else
        Exit Sub
        Unload Me
    End If

    Set Report = New Rpt_aPO
    Set oRS = rsGoSublet
    Report.database.SetDataSource oRS
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    Report.txtAdh.SetText Profil.nNama_ADH
    Call SetReport
End Sub

Private Sub createTbl_Go_Sublet_ALFA(Optional Edit As Boolean = False)
    If rsGoSublet.State Then rsGoSublet.Close
    With rsGoSublet
        .Fields.Append "NoPo", adVarChar, 20
        .Fields.Append "NoWo", adVarChar, 15
        .Fields.Append "NoRangka", adVarChar, 30
        .Fields.Append "NoPol", adVarChar, 15
        .Fields.Append "Pekerjaan", adVarChar, 100
        .Fields.Append "Qty", adDouble
        .Fields.Append "KodeSatuan", adVarChar, 15
        .Fields.Append "Keterangan", adVarChar, 100
        .Fields.Append "HargaJual", adDouble
        .Fields.Append "Supplier", adVarChar, 100
        .Open
    End With
End Sub

Private Sub Go_Order()                           '08
    Set Report = New Rpt_aWorkOrder
    Set oRS = objR.Order_Main(nNoWo)
    Report.database.SetDataSource oRS
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.Meter0.Left = 15000
    Report.Meter1.Left = 15000
    Report.meter2.Left = 15000
    Report.meter3.Left = 15000
    Report.meter4.Left = 15000
    Select Case nMeter
        Case 0
            Report.Meter0.Left = 480
        Case 1
            Report.Meter1.Left = 480
        Case 2
            Report.meter2.Left = 480
        Case 3
            Report.meter3.Left = 480
        Case 4
            Report.meter4.Left = 480
    End Select
        
    If nTipe = "A" Then
        Report.PicA.Left = 120
        Report.PicB.Left = 15000
        Report.PicC.Left = 15000
        Report.PicD.Left = 15000
        Report.PicE.Left = 15000
        Report.PicG.Left = 15000
        Report.PicT.Left = 15000
    ElseIf nTipe = "B" Then
        Report.PicB.Left = 120
        Report.PicA.Left = 15000
        Report.PicC.Left = 15000
        Report.PicD.Left = 15000
        Report.PicE.Left = 15000
        Report.PicG.Left = 15000
        Report.PicT.Left = 15000
    ElseIf nTipe = "C" Then
        Report.PicC.Left = 120
        Report.PicB.Left = 15000
        Report.PicA.Left = 15000
        Report.PicD.Left = 15000
        Report.PicE.Left = 15000
        Report.PicG.Left = 15000
        Report.PicT.Left = 15000
    ElseIf nTipe = "D" Then
        Report.PicD.Left = 120
        Report.PicB.Left = 15000
        Report.PicC.Left = 15000
        Report.PicA.Left = 15000
        Report.PicE.Left = 15000
        Report.PicG.Left = 15000
        Report.PicT.Left = 15000
    ElseIf nTipe = "E" Then
        Report.PicE.Left = 120
        Report.PicB.Left = 15000
        Report.PicC.Left = 15000
        Report.PicD.Left = 15000
        Report.PicA.Left = 15000
        Report.PicG.Left = 15000
        Report.PicT.Left = 15000
    ElseIf nTipe = "F" Then
        Report.PicG.Left = 120
        Report.PicB.Left = 15000
        Report.PicC.Left = 15000
        Report.PicD.Left = 15000
        Report.PicE.Left = 15000
        Report.PicA.Left = 15000
        Report.PicT.Left = 15000
    ElseIf nTipe = "G" Then
        Report.PicT.Left = 120
        Report.PicB.Left = 15000
        Report.PicC.Left = 15000
        Report.PicD.Left = 15000
        Report.PicE.Left = 15000
        Report.PicG.Left = 15000
        Report.PicA.Left = 15000
    End If
    For E = 0 To 295
        If nCroz(E) = 300 Then
            Exit For
        Else
            Call ViewCroz(nCroz(E))
        End If
    Next E
    Dim nOR As String
    nOR = "(Nilai OR :  Rp " & Format(objR.Order_JobOR(nNoWo), "###,###,###.00") & ")"
    If nOR = "(Nilai OR :  Rp )" Then nOR = "(Nilai OR :  Rp 0 )"
    Report.txtOR.SetText (nOR)
    Report.txtSPK.SetText objWO.FetchDataSPK(nNoWo)
    Report.txtEst.SetText "No Est : " & oRS.Fields(46).value
    '-- Order_JOB ------------------------------------------
        Dim jobRS As ADODB.Recordset
        Set jobRS = objR.Order_Job(nNoWo, 0)
        If jobRS.RecordCount > 10 Then
            Set jobRS = objR.Order_Job("Antztein", 0)
            LampWO = True
        Else
            Report.terlampir.Left = 15000
            LampWO = False
        End If
        Dim crxJob As New CRAXDRT.Report
        Set crxJob = Report.subJob.OpenSubreport
        crxJob.database.SetDataSource jobRS
        If rsDc.Fields(24).value = 0 Then
            crxJob.txt1.TextColor = &HFFFFFF
            crxJob.txt2.TextColor = &HFFFFFF
            crxJob.Len1.LineColor = &HFFFFFF
            crxJob.Len2.LineColor = &HFFFFFF
            crxJob.Field3.TextColor = &HFFFFFF
            crxJob.Field4.TextColor = &HFFFFFF
            Report.Len3.LineColor = &HFFFFFF
            Report.Len4.LineColor = &HFFFFFF
            
            Report.Text13.SetText "Total Part"
            Report.Field17.TextColor = &HFFFFFF
            Report.Text22.TextColor = &HFFFFFF
            Report.Field18.TextColor = &HFFFFFF
            Report.Text22.TextColor = &HFFFFFF
        End If
    '-------------------------------------------------------
    Call SetReport
End Sub

Private Sub Go_Order_addJob()                           '09
    Set Report = New Rpt_aAddJob
    Set oRS = objR.Order_Main_addJob(nNoWo)
    Report.database.SetDataSource oRS
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    '-- Order_JOB ------------------------------------------
        Dim crxJob As New CRAXDRT.Report
        Dim jobRS As ADODB.Recordset
        Set jobRS = objR.Order_Job(nNoWo, 1)
        Set crxJob = Report.subJob.OpenSubreport
        crxJob.database.SetDataSource jobRS
    '-------------------------------------------------------
        If rsDc.Fields(24).value = 0 Then
            crxJob.txt1.TextColor = &HFFFFFF
            crxJob.txt2.TextColor = &HFFFFFF
            crxJob.Len1.LineColor = &HFFFFFF
            crxJob.Len2.LineColor = &HFFFFFF
            crxJob.Field3.TextColor = &HFFFFFF
            crxJob.Field4.TextColor = &HFFFFFF
        End If
    Call SetReport
End Sub

Private Sub Go_MaterialGroup()                          '10
    Set Report = New Rpt_rMaterialByGroup
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtGroup.SetText nGroup
    Report.txtGroupKet.SetText nGroupKet
    Report.txtRange.SetText Trim(Format(nTgl0, "DD/MM/YYYY") & " - " & Format(nTgl1, "DD/MM/YYYY"))
    '-- JOB --------------------------------------------
        Dim crxJob As New CRAXDRT.Report
        Dim jobRS As ADODB.Recordset
        Set jobRS = objWO.FetchWO_Order_Finish(nTgl0, nTgl1, nGroup)
        Set crxJob = Report.subJob.OpenSubreport
        crxJob.database.SetDataSource jobRS
        If jobRS.RecordCount > 0 Then
            Dim xJob As Double
            jobRS.MoveFirst
            Do Until jobRS.EOF
                xJob = xJob + jobRS.Fields(5).value
                jobRS.MoveNext
            Loop
        End If
        crxJob.txt5.SetText txt5
        crxJob.txt6.SetText ("Rp   " & txt6)
    '-- Material -------------------------------------------
        Dim xBhn As Double
        Dim crxMat As New CRAXDRT.Report
        Dim MatRS As ADODB.Recordset
        Dim nAlfa As String
        Dim rsAlfa As ADODB.Recordset
        Set rsAlfa = objR.MaterialGroup_MATERIAL_BAHAN_ALFA(nGroup, False, nTgl0, nTgl1)
        If rsAlfa.RecordCount = 0 Then
            nAlfa = "('_')"
        Else
            C = 0
            rsAlfa.MoveFirst
            Do Until rsAlfa.EOF
                If rsAlfa.Fields(0).value <> "" Then
                    If C = 0 Then
                        nAlfa = "'" & rsAlfa.Fields(0).value & "'"
                    Else
                        nAlfa = nAlfa & "," & "'" & rsAlfa.Fields(0).value & "'"
                    End If
                    C = C + 1
                End If
                rsAlfa.MoveNext
            Loop
            nAlfa = "(" & nAlfa & ")"
        End If
        Set MatRS = objR.MaterialGroup_MATERIAL_BAHAN_BETA(nGroup, False, nTgl0, nTgl1, nAlfa)
        Set crxMat = Report.SubMat.OpenSubreport
        crxMat.database.SetDataSource MatRS
        If MatRS.RecordCount > 0 Then
            MatRS.MoveFirst
            Do Until MatRS.EOF
                xBhn = xBhn + MatRS.Fields(5).value
                MatRS.MoveNext
            Loop
        End If
    '-- Cat -----------------------------------------
        Dim crxCat As New CRAXDRT.Report
        Dim CatRS As ADODB.Recordset
        Set rsAlfa = objR.MaterialGroup_MATERIAL_BAHAN_ALFA(nGroup, True, nTgl0, nTgl1)
        If rsAlfa.RecordCount = 0 Then
            nAlfa = "('_')"
        Else
            C = 0
            rsAlfa.MoveFirst
            Do Until rsAlfa.EOF
                If rsAlfa.Fields(0).value <> "" Then
                    If C = 0 Then
                        nAlfa = "'" & rsAlfa.Fields(0).value & "'"
                    Else
                        nAlfa = nAlfa & "," & "'" & rsAlfa.Fields(0).value & "'"
                    End If
                    C = C + 1
                End If
                rsAlfa.MoveNext
            Loop
            nAlfa = "(" & nAlfa & ")"
        End If
        Set CatRS = objR.MaterialGroup_MATERIAL_BAHAN_BETA(nGroup, True, nTgl0, nTgl1, nAlfa)
        Set crxCat = Report.SubCat.OpenSubreport
        crxCat.database.SetDataSource CatRS
        If CatRS.RecordCount > 0 Then
            CatRS.MoveFirst
            Do Until CatRS.EOF
                xBhn = xBhn + CatRS.Fields(5).value
                CatRS.MoveNext
            Loop
        End If
    '-------------------------------------------------------
    Report.txtJob.SetText (Format(Val(xJob) - Val(Replace(txt5, ".", "")), "###,###,###.00"))
    Report.txtBahan.SetText (Format(xBhn, "###,###,###.00"))
    Report.txtPersen.SetText Round(((xBhn / (Val(xJob) - Val(Replace(txt5, ".", "")))) * 100), 2) & " %"
    Report.txtW.SetText (nTipe & " %")
    Call SetReport
End Sub

Private Sub Go_MaterialWo()                             '11
    Set Report = New Rpt_rMaterialByWo
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtNoWO.SetText nNoWo
    Dim rsWo As ADODB.Recordset
    Set rsWo = objR.MaterialWO(nNoWo)
    Report.txtKodelgn.SetText Trim(rsWo.Fields(0).value)
    Report.txtNama.SetText Trim(rsWo.Fields(1).value)
    Report.txtAlamat.SetText Trim(rsWo.Fields(2).value)
    Report.txtRangka.SetText Trim(rsWo.Fields(4).value)
    Report.txtModel.SetText Trim(rsWo.Fields(5).value)
    Report.txtModelKet.SetText Trim(rsWo.Fields(6).value)
    Report.txtNoPol.SetText Trim(rsWo.Fields(7).value)
    Report.txtGroup.SetText Trim("Group : " & rsWo.Fields(8).value)
    '-- JOB --------------------------------------------
        Dim crxJob As New CRAXDRT.Report
        Dim jobRS As ADODB.Recordset
        Set jobRS = objR.MaterialWO_JOB(nNoWo)
        Set crxJob = Report.subJob.OpenSubreport
        crxJob.database.SetDataSource jobRS
        If jobRS.RecordCount > 0 Then
            Dim xJob As Double
            jobRS.MoveFirst
            Do Until jobRS.EOF
                xJob = xJob + jobRS.Fields(5).value
                jobRS.MoveNext
            Loop
        End If
        crxJob.txt5.SetText txt5
        crxJob.txt6.SetText ("Rp   " & txt6)
    '-- Material ---------------------------------------
        Dim xBhn As Double
        Dim crxMat As New CRAXDRT.Report
        Dim MatRS As ADODB.Recordset
        Set MatRS = objR.MaterialWO_MATERIAL_BAHAN(nNoWo, False, nInv)
        Set crxMat = Report.SubMat.OpenSubreport
        crxMat.database.SetDataSource MatRS
        If MatRS.RecordCount > 0 Then
            MatRS.MoveFirst
            Do Until MatRS.EOF
                xBhn = xBhn + MatRS.Fields(5).value
                MatRS.MoveNext
            Loop
        End If
    '-- Cat -----------------------------------------
        Dim crxCat As New CRAXDRT.Report
        Dim CatRS As ADODB.Recordset
        Set CatRS = objR.MaterialWO_MATERIAL_BAHAN(nNoWo, True, nInv)
        Set crxCat = Report.SubCat.OpenSubreport
        crxCat.database.SetDataSource CatRS
        If CatRS.RecordCount > 0 Then
            CatRS.MoveFirst
            Do Until CatRS.EOF
                xBhn = xBhn + CatRS.Fields(5).value
                CatRS.MoveNext
            Loop
        End If
    '-------------------------------------------------------
    Report.txtJob.SetText (Format(Val(xJob) - Val(Replace(txt5, ".", "")), "###,###,###.00"))
    Report.txtBahan.SetText (Format(xBhn, "###,###,###.00"))
    Report.txtPersen.SetText Round(((xBhn / (Val(xJob) - Val(Replace(txt5, ".", "")))) * 100), 2) & " %"
    Report.txtW.SetText (nTipe & " %")
    Call SetReport
End Sub

Private Sub Go_Material()                             '12
    Set Report = New Rpt_rMaterial
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText Trim(Format(nTgl0, "DD/MM/YYYY") & " - " & Format(nTgl1, "DD/MM/YYYY"))
    '-- JOB --------------------------------------------
        Dim crxJob As New CRAXDRT.Report
        Dim jobRS As ADODB.Recordset
        Set jobRS = objR.Material_JOB(nTgl0, nTgl1)
        Set crxJob = Report.subJob.OpenSubreport
        crxJob.database.SetDataSource jobRS
        If jobRS.RecordCount > 0 Then
            Dim xJob As Double
            jobRS.MoveFirst
            Do Until jobRS.EOF
                xJob = xJob + jobRS.Fields(5).value
                jobRS.MoveNext
            Loop
        End If
    '-- Material -------------------------------------------
        Dim xBhn As Double
        Dim crxMat As New CRAXDRT.Report
        Dim MatRS As ADODB.Recordset
        Dim nAlfa As String
        Dim rsAlfa As ADODB.Recordset
        Set rsAlfa = objR.Material_MATERIAL_BAHAN_ALFA(False, nTgl0, nTgl1)
        If rsAlfa.RecordCount = 0 Then
            nAlfa = "('_')"
        Else
            C = 0
            rsAlfa.MoveFirst
            Do Until rsAlfa.EOF
                If rsAlfa.Fields(0).value <> "" Then
                    If C = 0 Then
                        nAlfa = "'" & rsAlfa.Fields(0).value & "'"
                    Else
                        nAlfa = nAlfa & "," & "'" & rsAlfa.Fields(0).value & "'"
                    End If
                    C = C + 1
                End If
                rsAlfa.MoveNext
            Loop
            nAlfa = "(" & nAlfa & ")"
        End If
        Set MatRS = objR.Material_MATERIAL_BAHAN_BETA(False, nTgl0, nTgl1, nAlfa)
        Set crxMat = Report.SubMat.OpenSubreport
        crxMat.database.SetDataSource MatRS
        If MatRS.RecordCount > 0 Then
            MatRS.MoveFirst
            Do Until MatRS.EOF
                xBhn = xBhn + MatRS.Fields(5).value
                MatRS.MoveNext
            Loop
        End If
    '-- Cat -----------------------------------------
        Dim crxCat As New CRAXDRT.Report
        Dim CatRS As ADODB.Recordset
        Set rsAlfa = objR.Material_MATERIAL_BAHAN_ALFA(True, nTgl0, nTgl1)
        If rsAlfa.RecordCount = 0 Then
            nAlfa = "('_')"
        Else
            C = 0
            rsAlfa.MoveFirst
            Do Until rsAlfa.EOF
                If rsAlfa.Fields(0).value <> "" Then
                    If C = 0 Then
                        nAlfa = "'" & rsAlfa.Fields(0).value & "'"
                    Else
                        nAlfa = nAlfa & "," & "'" & rsAlfa.Fields(0).value & "'"
                    End If
                    C = C + 1
                End If
                rsAlfa.MoveNext
            Loop
            nAlfa = "(" & nAlfa & ")"
        End If
        Set CatRS = objR.Material_MATERIAL_BAHAN_BETA(True, nTgl0, nTgl1, nAlfa)
        Set crxCat = Report.SubCat.OpenSubreport
        crxCat.database.SetDataSource CatRS
        If CatRS.RecordCount > 0 Then
            CatRS.MoveFirst
            Do Until CatRS.EOF
                xBhn = xBhn + CatRS.Fields(5).value
                CatRS.MoveNext
            Loop
        End If
    '-------------------------------------------------------
    On Error Resume Next
    Report.txtJob.SetText (Format(xJob, "###,###,###.00"))
    Report.txtBahan.SetText (Format(xBhn, "###,###,###.00"))
    Report.txtPersen.SetText Round(((xBhn / xJob) * 100), 2) & " %"
    Call SetReport
End Sub

Private Sub Go_Progress()                           '13
    Set Report = New Rpt_aProgress
    Set oRS = objR.Order_Main(nNoWo)
    Report.database.SetDataSource oRS
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtW.SetText data(0)
    Report.txtS.SetText data(1)
    Report.txtP.SetText data(2)
    If nTipe = "A" Then
        Report.PicA.Left = 120
        Report.PicB.Left = 15000
        Report.PicC.Left = 15000
    ElseIf nTipe = "B" Then
        Report.PicA.Left = 15000
        Report.PicB.Left = 120
        Report.PicC.Left = 15000
    Else
        Report.PicA.Left = 15000
        Report.PicB.Left = 15000
        Report.PicC.Left = 120
    End If
    For E = 0 To 295
        If nCroz(E) = 300 Then
            Exit For
        Else
            Call ViewCroz(nCroz(E))
        End If
    Next E
    '-- Order_JOB ------------------------------------------
        Dim crxJob As New CRAXDRT.Report
        Dim jobRS As ADODB.Recordset
        Set jobRS = objR.Order_Job(nNoWo, 0)
        Set crxJob = Report.subJob.OpenSubreport
        crxJob.database.SetDataSource jobRS
    '-- Progress -------------------------------------------
        Dim crxPro As New CRAXDRT.Report
        Set crxPro = Report.SubPro.OpenSubreport
        crxPro.database.SetDataSource nRS
    '-------------------------------------------------------
    Call SetReport
End Sub

Private Sub Go_PoSuplier()                              '14
    Dim rsSp As ADODB.Recordset
    Set rsSp = objR.Sublet_PoSup(nNoWo, nSup)
    
    Set Report = New Rpt_aPoSup

    Report.database.SetDataSource rsSp
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    Report.txtJabatan.SetText "Section Head"
    Report.txtAdh.SetText Profil.nNama_ADH
    Dim rsSup As ADOR.Recordset
    Set rsSup = objR.Sublet_NamaSup(nSup)
    Report.txtSup.SetText (rsSup.Fields(0).value & "     ( KODE AP : " & rsSup.Fields(2).value & " )")
    Report.txtAla.SetText rsSup.Fields(1).value
    Call SetReport
End Sub

Private Sub Go_Order_10()                           '15
    Set Report = New Rpt_aWorkOrder_10
    Set oRS = objR.Order_Main(nNoWo)
    Report.database.SetDataSource oRS
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    '-- Order_JOB ------------------------------------------
        Dim jobRS As ADODB.Recordset
        Set jobRS = objR.Order_Job(nNoWo, 0)
        Dim crxJob As New CRAXDRT.Report
        Set crxJob = Report.subJob.OpenSubreport
        crxJob.database.SetDataSource jobRS
    '-------------------------------------------------------
        If rsDc.Fields(24).value = 0 Then
            crxJob.txt1.TextColor = &HFFFFFF
            crxJob.txt2.TextColor = &HFFFFFF
            crxJob.Len1.LineColor = &HFFFFFF
            crxJob.Len2.LineColor = &HFFFFFF
            crxJob.Field3.TextColor = &HFFFFFF
            crxJob.Field4.TextColor = &HFFFFFF
            Report.len5.LineColor = &HFFFFFF
            Report.len6.LineColor = &HFFFFFF
        End If
        
    Call SetReport
End Sub

Private Sub Go_Sublet_Invoice()                     '16
    Dim rsSi As ADODB.Recordset
    Set rsSi = objR.Sublet_Invoice(nNoWo)
    
    Set Report = New Rpt_sInvSublet

    Report.database.SetDataSource rsSi
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    Report.txtPrint.SetText Profil.nKotaProfil & ", " & Format(Now, "DD/MM/YYYY")
    Call SetReport
End Sub

Private Sub Go_Repair_History()                     '17
    Dim rsRh As ADODB.Recordset
    Set rsRh = objR.Repair_Hist(nNoRangka, nDt)
    
    Set Report = New Rpt_eRepHistory

    Report.database.SetDataSource rsRh
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Call SetReport
End Sub

Private Sub Go_Job_Light()                          '18
    Dim rsJL As ADODB.Recordset
    Set rsJL = objR.Data_Job_Light(nTgl0, nTgl1, nDt, nTipe)
    
    Set Report = New Rpt_eJobLight

    Report.database.SetDataSource rsJL
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")) & (IIf(nDt = 0, ", Sudah Invoice", IIf(nDt = 1, ", Belum Invoice", ", Semua Data Order")) & nTipe))
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_Job_MediumHeavy()                    '19
    Dim rsJL As ADODB.Recordset
    Set rsJL = objR.Data_Job_MediumHeavy(nTgl0, nTgl1, nDt, nTipe)
    
    Set Report = New Rpt_eJobMediumHeavy

    Report.database.SetDataSource rsJL
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")) & (IIf(nDt = 0, ", Sudah Invoice", IIf(nDt = 1, ", Belum Invoice", ", Semua Data Order")) & nTipe))
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_RepOrder()                           '20
    Dim rsRO As ADODB.Recordset
    Set rsRO = objR.Data_Job_RepOrder(nTgl0, nTgl1, nDt, nTipe, nGroup, nGroupKet, nJob, nInc, CInt(nSup), nAsr, nCab)
    If nCust = False Then
        Set Report = New Rpt_eOrder
    Else
        Set Report = New Rpt_eOrderCust
    End If
    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")) & (IIf(nDt = 0, ", Sudah Invoice", IIf(nDt = 1, ", Belum Invoice", ", Semua Data Order")) & nTipe))
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_SetTagih()                           '21
    Dim rsSt As ADODB.Recordset
    Set rsSt = objIn.FetchDataTagihanBorong_SetDitagihkan(nNoWo, 1)
    
    rsSt.MoveFirst
    Total = 0
    Do Until rsSt.EOF
        Total = Total + rsSt.Fields(10).value
        rsSt.MoveNext
    Loop
    rsSt.MoveFirst

    Set Report = New Rpt_sSetTagihan

    Report.database.SetDataSource rsSt
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    Report.txtKacab.SetText Profil.nNama_KaCab
    Report.txtAdh.SetText Profil.nNama_ADH
    
    Dim rsx As ADODB.Recordset
    Set rsx = objMk.FetchDataGroup(rsSt.Fields(4).value)
    Dim nPPh As Double
    If rsx.RecordCount > 0 Then
        Report.txtBorong.SetText rsx.Fields(3).value
        nPPh = (Val(rsx.Fields(10).value) / 100) * Total
        Report.txtPPh21persen.SetText "Potongan PPH (" & rsx.Fields(10).value & "%)"
        Report.txtNpwp.SetText rsx.Fields(4).value
        
        Report.txtNama.SetText rsx.Fields(6).value
        Report.txtPIC.SetText rsx.Fields(3).value
        Report.txtAla1.SetText rsx.Fields(7).value
        Report.txtAla2.SetText rsx.Fields(8).value
        Report.txtAla3.SetText ""
        
        Report.txtAP.SetText rsx.Fields(9).value
    End If
    Report.txtPPh21.SetText Format(nPPh, "###,###,###.00")
    Report.txtTagihan.SetText Format((Total - nPPh), "###,###,###.00")
    Report.txtPrint.SetText Profil.nKotaProfil & ", " & Format(Now, "DD/MM/YYYY")
    Report.txtBil.SetText Angka2Huruf((Total - nPPh), 1)
    If nCust = False Then Report.txtKet.SetText "Keterangan :  - CETAK ULANG -"
    Call SetReport
End Sub

Private Sub Go_SetTagih_Ulang()                     '22
    Dim rsSt As ADODB.Recordset
    Set rsSt = objIn.FetchDataTagihanBorong_SetDitagihkan(nTipe, 2)
    rsSt.MoveFirst
    Total = 0
    Do Until rsSt.EOF
        Total = Total + rsSt.Fields(10).value
        rsSt.MoveNext
    Loop
    rsSt.MoveFirst
    Set Report = New Rpt_sSetTagihan

    Report.database.SetDataSource rsSt
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    Report.txtKacab.SetText Profil.nNama_KaCab
    Report.txtAdh.SetText Profil.nNama_ADH
    Report.txtAP.SetText Profil.nApBorong
    
    Dim rsx As ADODB.Recordset
    Set rsx = objMk.FetchDataGroup(rsSt.Fields(4).value)
    Dim nPPh As Double
    If rsx.RecordCount > 0 Then
        Report.txtBorong.SetText rsx.Fields(3).value
        nPPh = (Val(rsx.Fields(10).value) / 100) * Total
        Report.txtPPh21persen.SetText "Potongan PPH (" & rsx.Fields(10).value & "%)"
        Report.txtNpwp.SetText rsx.Fields(4).value
        
        Report.txtNama.SetText rsx.Fields(6).value
        Report.txtPIC.SetText rsx.Fields(3).value
        Report.txtAla1.SetText rsx.Fields(7).value
        Report.txtAla2.SetText rsx.Fields(8).value
        Report.txtAla3.SetText rsx.Fields(9).value
    End If
    Report.txtPPh21.SetText Format(nPPh, "###,###,###.00")
    Report.txtTagihan.SetText Format((Total - nPPh), "###,###,###.00")
    Report.txtPrint.SetText Profil.nKotaProfil & ", " & Format(Now, "DD/MM/YYYY")
    Report.txtBil.SetText Angka2Huruf((Total - nPPh), 1)
    If nCust = False Then Report.txtKet.SetText "Keterangan :  - CETAK ULANG / COPY -"
    Call SetReport


End Sub

Private Sub Go_SetTagihSublet()                     '23
    Dim rsSs As ADODB.Recordset
    Set rsSs = objIn.FetchDataSublet_SetDitagihkan(nNoWo, 3)
    Call createTbl_Tagih(True)
    If rsSs.RecordCount > 0 Then
        rsSs.MoveFirst
        nTot = 0
        Do Until rsSs.EOF
                rsGrid.AddNew
                rsGrid!NoPO = Nul(rsSs.Fields(0).value)
                rsGrid!noWO = Nul(rsSs.Fields(1).value)
                rsGrid!noinv = Nul(rsSs.Fields(2).value)
                rsGrid!noRangka = Nul(rsSs.Fields(3).value)
                rsGrid!NoPolisi = Nul(rsSs.Fields(4).value)
                rsGrid!Pekerjaan = Nul(rsSs.Fields(5).value)
                rsGrid!qty = Val(rsSs.Fields(6).value)
                rsGrid!KodeSatuan = Nul(rsSs.Fields(7).value)
                rsGrid!NoNota = Nul(rsSs.Fields(8).value)
                rsGrid!HargaBeli = Val(rsSs.Fields(9).value)
                rsGrid!Supplier = objSb.FetchNamaSupplier(rsSs.Fields(10).value)
                rsGrid!KodeTagih = Nul(rsSs.Fields(11).value)
                rsGrid!TglTagih = Format(TglNul(rsSs.Fields(12).value), "MM/DD/YYYY")
                rsGrid.Update
            rsSs.MoveNext
        Loop
    End If

    Set Report = New Rpt_sTagihPO

    Report.database.SetDataSource rsGrid
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    Report.txtAdh.SetText Profil.nNama_ADH
    Report.txtPrint.SetText Profil.nKotaProfil & ", " & Format(isnow, "DD/MM/YYYY")
    Report.txtBil.SetText Angka2Huruf(Total, 1)
    If nCust = False Then Report.txtKet.SetText "Keterangan :  - CETAK ULANG / COPY -"
    Call SetReport
End Sub

Private Sub Go_SetTagihSublet_Ulang()                     '24
    Dim rsSs As ADODB.Recordset
    Set rsSs = objIn.FetchDataSublet_SetDitagihkan(nTipe, 2)
    Call createTbl_Tagih(True)
    If rsSs.RecordCount > 0 Then
        rsSs.MoveFirst
        Total = 0
        Do Until rsSs.EOF
                rsGrid.AddNew
                rsGrid!NoPO = Nul(rsSs.Fields(0).value)
                rsGrid!noWO = Nul(rsSs.Fields(1).value)
                rsGrid!noinv = Nul(rsSs.Fields(2).value)
                rsGrid!noRangka = Nul(rsSs.Fields(3).value)
                rsGrid!NoPolisi = Nul(rsSs.Fields(4).value)
                rsGrid!Pekerjaan = Nul(rsSs.Fields(5).value)
                rsGrid!qty = Val(rsSs.Fields(6).value)
                rsGrid!KodeSatuan = Nul(rsSs.Fields(7).value)
                rsGrid!NoNota = Nul(rsSs.Fields(8).value)
                rsGrid!HargaBeli = Val(rsSs.Fields(9).value)
                rsGrid!Supplier = objSb.FetchNamaSupplier(rsSs.Fields(10).value)
                rsGrid!KodeTagih = Nul(rsSs.Fields(11).value)
                rsGrid!TglTagih = Format(TglNul(rsSs.Fields(12).value), "MM/DD/YYYY")
                Total = Total + Val(rsSs.Fields(9).value)
                rsGrid.Update
            rsSs.MoveNext
        Loop
    End If

    Set Report = New Rpt_sTagihPO

    Report.database.SetDataSource rsGrid
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    Report.txtAdh.SetText Profil.nNama_ADH
    Report.txtPrint.SetText Profil.nKotaProfil & ", " & Format(isnow, "DD/MM/YYYY")
    Report.txtBil.SetText Angka2Huruf(Total, 1)
    Report.txtKet.SetText "Keterangan : Cetaka Ulang / Copy !"
    If nCust = False Then Report.txtKet.SetText "Keterangan :  - CETAK ULANG / COPY -"
    Call SetReport
End Sub

Private Sub createTbl_Tagih(Optional Edit As Boolean = False)
    If rsGrid.State Then rsGrid.Close
    With rsGrid
        .Fields.Append "NoPO", adVarChar, 15
        .Fields.Append "NoWo", adVarChar, 15
        .Fields.Append "NoInv", adVarChar, 15
        .Fields.Append "NoRangka", adVarChar, 25
        .Fields.Append "NoPolisi", adVarChar, 15
        .Fields.Append "Pekerjaan", adVarChar, 100
        .Fields.Append "Qty", adDouble
        .Fields.Append "KodeSatuan", adVarChar, 100
        .Fields.Append "NoNota", adVarChar, 25
        .Fields.Append "HargaBeli", adDouble
        .Fields.Append "Supplier", adVarChar, 50
        .Fields.Append "KodeTagih", adVarChar, 50
        .Fields.Append "TglTagih", adDate
        .Open
    End With
End Sub

Private Sub Go_Info_Estimasi()                     '25
    Dim rsRO As ADODB.Recordset
    Set rsRO = objE.FetchDataInfoEst(nTgl0, nTgl1, nGroup, nGroupKet, nAsr)
    
    Set Report = New Rpt_eEstimasi

    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")) & IIf(nGroup <> "", ", SA : " & nGroup, "") & IIf(nGroupKet <> "", " (" & nGroupKet & ")", ""))
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_OutstandingSPK()                     '26
    Dim rsRO As ADODB.Recordset
    Set rsRO = objAs.Fetch_Outstanding_SPK(nTgl0, nTgl1, IIf(nTipe = ", Sudah WO", 1, IIf(nTipe = ", Belum WO", 2, 3)))
    
    Set Report = New Rpt_eOutstandingSPK

    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")) & nTipe)
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_ValidasiEXT()                        '27
    Dim rsRO As ADODB.Recordset
    Set rsRO = objIn.Fetch_DataValid_Ext(nTgl0, nTgl1, nDt, nTipe, nGroup)
    
    Set Report = New Rpt_eValidExt

    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")) & nDt & nTipe & nGroup)
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_ValidasiEXT_SUM()                        '28
    Dim rsRO As ADODB.Recordset
    Set rsRO = objIn.Fetch_DataValid_Ext(nTgl0, nTgl1, nDt, nTipe, nGroup)
    
    Set Report = New Rpt_eValidExt_SUM
    
    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")) & nDt & nTipe & nGroup)
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_Invoice()                                '29
    Dim rsIn As ADODB.Recordset
    Set rsIn = objIn.FetchInvoiceReports(nInv)
    
    If rsIn.RecordCount > 0 Then
    Set Report = New Rpt_sInvoice

    Report.database.SetDataSource rsIn
    
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtKode.SetText "Kode : " & rsIn.Fields(0).value & IIf(Nul(rsIn.Fields(3).value) <> "", " (" & rsIn.Fields(4).value & ")", "")
    Report.txtTelp.SetText Profil.nTeleponProfil
    Report.txtFax.SetText Profil.nFaxProfil
    Report.txtPrint.SetText Profil.nKotaProfil & ", " & rsIn.Fields(28).value
    Report.txtSlip.SetText rsIn.Fields(29).value
    Report.txtBilang.SetText Angka2Huruf(rsIn.Fields(26).value, 1)
    Report.txtSPK.SetText objWO.FetchTglSPK(rsIn.Fields(30).value)
    Report.txtCatat.SetText txt5
    Report.txtORm.SetText objWO.FetchOrMinus(rsIn.Fields(30).value)
    If Right(nInv, 1) <> "R" Then
        Report.txtR.SetText ""
        Report.txtR2.SetText ""
    End If
    Dim nOR As Double
    nOR = objWO.FetchNilaiOR(rsIn.Fields(30).value)
    If nOR = 0 Then
        Report.txtOR.SetText "Rp. 0,00"
    Else
        Report.txtOR.SetText "Rp. " & Format(nOR, "###,###,###.00")
    End If
    Report.txtAc.SetText Profil.nNoAC
    Report.txtKabeng.SetText Profil.nNama_Kabeng
    
    '-- Order_JOB ------------------------------------------
            
        Dim jobRS As ADODB.Recordset
        If rsIn.Fields(31).value <> "" Then
            Set jobRS = objR.Order_Job_Inv(rsIn.Fields(30).value)
        Else
            Set jobRS = objR.Order_Job_Inv(objWO.FetchItemWO_InvRetur(nInv))
        End If
        Dim crxJob As New CRAXDRT.Report
        Set crxJob = Report.JobOrderInv.OpenSubreport
        crxJob.database.SetDataSource jobRS
    '-------------------------------------------------------
        If rsDc.Fields(24).value = 0 Then
            crxJob.txt1.TextColor = &HFFFFFF
            crxJob.txt2.TextColor = &HFFFFFF
            crxJob.Len1.LineColor = &HFFFFFF
            crxJob.Len2.LineColor = &HFFFFFF
            crxJob.Field3.TextColor = &HFFFFFF
            crxJob.Field4.TextColor = &HFFFFFF
            Report.Len3.LineColor = &HFFFFFF
            Report.Len4.LineColor = &HFFFFFF
        End If
    Call SetReport
    End If
End Sub


Private Sub Go_Faktur()                                '30
'    If TipeFP = True Then
        Dim rsF As ADODB.Recordset
        Set rsF = objFP.GetFP_Tunggal(noFP, nRS.Fields(6).value)
        Select Case CInt(nRS.Fields(2).value)
            Case 0
                Set Report = New ReportFakturPajak4LQ2180
            Case 1
                Set Report = New ReportFakturPajak4LQ300
            Case 2
                Set Report = New ReportFakturPajak4LX300
            Case 3
                Set Report = New ReportFakturPajak4LQ2190
        End Select
    
        Report.database.SetDataSource rsF
        If CInt(nRS.Fields(3).value) = 0 Then
            Report.Text1.SetText ""
        End If
        If CInt(nRS.Fields(4).value) = 0 Then
            Report.Text2.SetText ""
        End If
        If CInt(nRS.Fields(5).value) = 0 Then
            Report.Text10.SetText ""
        End If
        Report.txtKotaTgl.SetText Trim(Profil.nKotaProfil) & ", " & Format(rsF.Fields(2).value, "DD MMMM YYYY")
        Report.TxtNamaTtd.SetText Pejabat
        Report.TxtBag.SetText Jabatan
        
        Report.txtNamaPT.SetText Profil.nDealerAcc
        Report.txtAlamatPT.SetText Profil.nAlamatD
        Report.txtNPWPPT.SetText Profil.nNpwp
        
        Call SetReport
End Sub

Private Sub Go_FakturLamp()                                '31
    Dim rsL As ADODB.Recordset
    Set rsL = objFP.FetchLampiranGabungan(noFP)
    Set Report = New ReportFakturPajakLampOK
    Report.database.SetDataSource rsL
    Call SetReport
End Sub

Private Sub Go_KuitansiOR()                                '32
    Dim rsOr As ADODB.Recordset
    Set rsOr = objIn.FetchNilaiOR(data(2))

    Set Report = New Rpt_sKuitansiOR
    Report.database.SetDataSource rsOr
    
    Report.txtNama.SetText data(0)
    Report.txtNoPol.SetText data(1)
    Report.txtNoWO.SetText data(2)
    Report.txtAsuransi.SetText data(3)
    Report.txtNilai.SetText Angka2Huruf(Total, 1)
    Report.txtTotal.SetText "Rp. " & Format(Total, "###,###,###.00")
    
    Report.txtTgl.SetText (Profil.nKotaProfil & ", " & Format(isnow, "DD MMMM YYYY"))
    Report.txtNas.SetText UCase(Profil.nNamaDealer)
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtTelp.SetText Profil.nTeleponProfil
    Report.txtFax.SetText Profil.nFaxProfil
    Report.txtNo.SetText "No." & rsOr.Fields(1).value
    Report.txtKet.SetText rsOr.Fields(7).value
    
    Call SetReport
End Sub


Private Sub Go_RepBooking()                                   '33
'    Dim rsRO As ADODB.Recordset
'    Set rsRO = objR.Data_Job_RepOrder(nTgl0, nTgl1, nDt, nTipe)
'
'    Set Report = New Rpt_eOrder
'
'    Report.Database.SetDataSource rsRO
'    Report.txtNas.SetText Profil.nNamaDealer
'    Report.txtAl.SetText Profil.nAlamatProfil
'    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")) & (IIf(nDt = 0, ", Sudah Invoice", IIf(nDt = 1, ", Belum Invoice", ", Semua Data Order")) & nTipe))
'    Report.txtDate.SetText isNow
'    Call SetReport
End Sub

Private Sub Go_RepBatalOrder()                                 '34
    Dim rsRO As ADODB.Recordset
    Set rsRO = objR.Data_Job_RepBatalOrder(nTgl0, nTgl1, nGroup, nGroupKet, nJob)
    
    Set Report = New Rpt_eBatalOrder

    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")))
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_Info_Billing()
    If nRekap = True Then
        Call Load_RekapReport
        Exit Sub
    End If
    Dim rsBL As ADODB.Recordset
    Set rsBL = objR.Data_Invoice(nTgl0, nTgl1, nGroup, nGroupKet, nJob, nSup, nMeter, nAsr, nNoBok, nRetur, nLapBul)
    
    Set Report = New Rpt_eBilling

    Report.database.SetDataSource rsBL
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    If nRetur = False Then
        Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")))
    Else
        Report.txtRange.SetText ("Range Retur : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")))
        Report.Text3.SetText "Tanggal Inv Retur :"
    End If
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_Info_Billing2()
    If nRekap = True Then
        Call Load_RekapReport
        Exit Sub
    End If
    Dim rsBL As ADODB.Recordset
    Set rsBL = objR.Data_Invoice(nTgl0, nTgl1, nGroup, nGroupKet, nJob, nSup, nMeter, nAsr, nNoBok, nRetur, nLapBul)
    
    Set Report = New Rpt_eBilling2

    Report.database.SetDataSource rsBL
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    If nRetur = False Then
        Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")))
    Else
        Report.txtRange.SetText ("Range Retur : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")))
        Report.Text3.SetText "Tanggal Inv Retur :"
    End If
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_Info_Billing3()
    If nRekap = True Then
        Call Load_RekapReport
        Exit Sub
    End If
    Dim rsBL As ADODB.Recordset
    Set rsBL = objR.Data_Invoice(nTgl0, nTgl1, nGroup, nGroupKet, nJob, nSup, nMeter, nAsr, nNoBok, nRetur, nLapBul)
    
    Set Report = New Rpt_eBilling3

    Report.database.SetDataSource rsBL
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    If nRetur = False Then
        Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")))
    Else
        Report.txtRange.SetText ("Range Retur : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")))
        Report.Text3.SetText "Tanggal Inv Retur :"
    End If
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_Info_Billing4()
    If nRekap = True Then
        Call Load_RekapReport
        Exit Sub
    End If
    Dim rsBL As ADODB.Recordset
    Set rsBL = objR.Data_Invoice(nTgl0, nTgl1, nGroup, nGroupKet, nJob, nSup, nMeter, nAsr, nNoBok, nRetur, nLapBul)
    
    Set Report = New Rpt_eBilling4

    Report.database.SetDataSource rsBL
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    If nRetur = False Then
        Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")))
    Else
        Report.txtRange.SetText ("Range Retur : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")))
        Report.Text3.SetText "Tanggal Inv Retur :"
    End If
    Report.txtDate.SetText isnow
    
    Dim NilaiI As Double
    Dim NilaiIr As Double
    
    NilaiI = objR.Data_Invoice_Count(nTgl0, nTgl1, False)
    NilaiR = objR.Data_Invoice_Count(nTgl0, nTgl1, True)
    
    Report.txtI.SetText NilaiI
    Report.txtIr.SetText NilaiR
    Report.txtITot.SetText NilaiI - NilaiR
    Call SetReport
End Sub


Private Sub Go_Info_DataAR()                       '38
    Dim rsRO As ADODB.Recordset
    Set rsRO = objIn.FetchDataInfoAR(nTgl0, nTgl1, nGroup, nGroupKet, nCust, nRetur)
    
    Set Report = New Rpt_eDataAR
    
    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Call SetReport
End Sub

Private Sub Go_Info_OR()                           '39
    Dim rsRO As ADODB.Recordset
    Set rsRO = objR.Data_OwnRisk(nTgl0, nTgl1, nDt, nTipe, nGroup, nGroupKet, nJob)
    Set Report = New Rpt_eOwnRisk
    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText ("Range : " & Trim(Format(nTgl0, "DD/MM/YYYY")) & " s/d " & Trim(Format(nTgl1, "DD/MM/YYYY")) & (IIf(nDt = 0, ", Sudah Invoice", IIf(nDt = 1, ", Belum Invoice", ", Semua Data Order")) & nTipe))
    Report.txtDate.SetText isnow
    Call SetReport
End Sub

Private Sub Go_Registrasi()                           '40
    Dim rsRO As ADODB.Recordset
    Set rsRO = objReg.FetchItemRegist_Print(nNoEst)
    Set Report = New Rpt_cRegistrasi
    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtNas2.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtAl2.SetText Profil.nAlamatProfil
    Report.txtData.SetText nJob
    Report.txtData2.SetText nJob
    Report.txtDate.SetText isnow
    Report.txtDate2.SetText isnow
    Call SetReport
End Sub

Private Sub Go_ReEstimasi()                           '41
    Dim rsRO As ADODB.Recordset
    Set rsRO = objE.FetchDataReEst(nTgl0, nTgl1, nSA, nGroup, nJob)
    Set Report = New Rpt_eReEstimasi
    Report.database.SetDataSource rsRO
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Call SetReport
End Sub

Private Sub Go_SetTagih_Group()                           '42
    Dim rsStG As ADODB.Recordset
    Set rsStG = objIn.FetchDataTagihanBorongGroup(nTgl0, nTgl1, nFilter, npersn)
    
'    rsStG.MoveFirst
'    Total = 0
'    Do Until rsSt.EOF
'        Total = Total + rsSt.Fields(10).value
'        rsSt.MoveNext
'    Loop
'    rsSt.MoveFirst

    Set Report = New Rpt_sSetTagihanGroup

    Report.database.SetDataSource rsStG
    Report.txtNas.SetText Profil.nNamaDealer
    Report.txtAL.SetText Profil.nAlamatProfil
    Report.txtRange.SetText nTgl0 & " s/d " & nTgl1
    Report.txtDate.SetText Format(isnow, "DD MMMM YYYY")
'    Report.txtAdh.SetText Profil.nNama_ADH
    
'    Dim rsx As ADODB.Recordset
'    Set rsx = objMk.FetchDataGroup(rsSt.Fields(4).value)
'    Dim nPPh As Double
'    If rsx.RecordCount > 0 Then
'        Report.txtBorong.SetText rsx.Fields(3).value
'        nPPh = (Val(rsx.Fields(10).value) / 100) * Total
'        Report.txtPPh21persen.SetText "Potongan PPH (" & rsx.Fields(10).value & "%)"
'        Report.txtNPWP.SetText rsx.Fields(4).value
'
'        Report.txtNama.SetText rsx.Fields(6).value
'        Report.txtPIC.SetText rsx.Fields(3).value
'        Report.txtAla1.SetText rsx.Fields(7).value
'        Report.txtAla2.SetText rsx.Fields(8).value
'        Report.txtAla3.SetText ""
'
'        Report.txtAP.SetText rsx.Fields(9).value
'    End If
'    Report.txtPPh21.SetText Format(nPPh, "###,###,###.00")
'    Report.txtTagihan.SetText Format((Total - nPPh), "###,###,###.00")
'    Report.txtPrint.SetText Profil.nKotaProfil & ", " & Format(Now, "DD/MM/YYYY")
'    Report.txtBil.SetText Angka2Huruf((Total - nPPh), 1)
'    If nCust = False Then Report.txtKet.SetText "Keterangan :  - CETAK ULANG -"
    Call SetReport
End Sub

Private Sub Load_RekapReport()
    Dim data(180) As String
    Set Report = New Rpt_sRekap
    Report.txtNas.SetText Profil.nNamaDealer
    Report.TextBln.SetText namaBulan(Month(nTgl0)) & " " & year(nTgl0)
    
    For E = 45 To 68
        If E = 48 Then
            data(E) = Val(data(45)) + Val(data(46)) + Val(data(47))
        ElseIf E = 51 Then
            data(E) = Val(data(49)) + Val(data(50))
        ElseIf E = 52 Then
            data(E) = Val(data(48)) + Val(data(51))
        ElseIf E = 56 Then
            data(E) = Val(data(53)) + Val(data(54)) + Val(data(55))
        ElseIf E = 62 Then
            data(E) = Val(data(58)) + Val(data(59)) + Val(data(60)) + Val(data(61))
        ElseIf E = 67 Then
            data(E) = Val(data(63)) + Val(data(64)) + Val(data(65)) + Val(data(66))
        ElseIf E = 68 Then
            data(E) = Val(data(62)) + Val(data(67))
        Else
            data(E) = objR.Data_Rekap(Month(nTgl0), year(nTgl0), E)
        End If
    Next E
    
    For E = 125 To 169
        If E = 130 Then
            data(E) = Val(data(125)) + Val(data(126)) + Val(data(127)) + Val(data(128)) + Val(data(129))
        ElseIf E = 141 Then
            data(E) = Val(data(131)) + Val(data(132)) + Val(data(133)) + Val(data(134)) + Val(data(135)) + Val(data(136)) + Val(data(137)) + Val(data(138)) + Val(data(139)) + Val(data(140))
        ElseIf E = 152 Then
            data(E) = Val(data(142)) + Val(data(143)) + Val(data(144)) + Val(data(145)) + Val(data(146)) + Val(data(147)) + Val(data(148)) + Val(data(149)) + Val(data(150)) + Val(data(151))
        ElseIf E = 153 Then
            data(E) = Val(data(141)) + Val(data(152))
        ElseIf E = 156 Then
            data(E) = Val(data(154)) + Val(data(155))
        ElseIf E = 161 Then
            data(E) = Val(data(157)) + Val(data(158)) + Val(data(159)) + Val(data(160))
        Else
            data(E) = objR.Data_Rekap(Month(nTgl0), year(nTgl0), E)
        End If
    Next E
    
    Dim rsOTD As ADODB.Recordset
    Dim rsNOTD As ADODB.Recordset
    Set rsOTD = objKp.FetchOTD(Month(nTgl0), year(nTgl0), "", "")
    Set rsNOTD = objKp.FetchNOTD(Month(nTgl0), year(nTgl0), "", "")
    If rsOTD.RecordCount > 0 Then
        data(167) = Trim(Round(((rsOTD.RecordCount * 100) / (rsOTD.RecordCount + rsNOTD.RecordCount)), 2)) & " %"
    Else
        data(167) = "0 %"
    End If
    data(17) = rsOTD.RecordCount & " dari " & (rsOTD.RecordCount + rsNOTD.RecordCount) & " Unit ="
    
    Dim rsRe As Double
    Dim rsNRe As Double
    rsRe = objR.Data_Rekap(Month(nTgl0), year(nTgl0), 166)
    rsNRe = objR.Data_Rekap(Month(nTgl0), year(nTgl0), 167)
    If rsRe > 0 Then
        data(166) = Trim(Round(((rsRe * 100) / (rsRe + rsNRe)), 2)) & " %"
    Else
        data(166) = "0 %"
    End If
    data(16) = rsRe & " dari " & rsNRe & " Unit ="
    
    '--------------------------------------------------------------------------------------------------
    '  ORI 14 JUNI 2016   |   Rev. Penambahan Menu KPI LeadTimeINV & Rev. INF Billing Rekap    |   ARIKLD
    '--------------------------------------------------------------------------------------------------
'    Dim rsSum As ADODB.Recordset, rsSumR As ADODB.Recordset
'    Set rsSum = objKp.FetchSumLeadTimeInv(Month(nTgl0), Year(nTgl0), "LIGHT", "")
'    data(168) = SetBreakdownTime(NulNum(rsSum.Fields(2).value))
'    data(175) = SetBreakdownTime(NulNum(Val(rsSum.Fields(2).value) / Val(objR.Data_Rekap(Month(nTgl0), Year(nTgl0), 53))))
'    Set rsSum = objKp.FetchSumLeadTimeInv(Month(nTgl0), Year(nTgl0), "MEDIUM", "")
'    data(169) = SetBreakdownTime(NulNum(rsSum.Fields(2).value))
'    data(176) = SetBreakdownTime(NulNum(Val(rsSum.Fields(2).value) / Val(objR.Data_Rekap(Month(nTgl0), Year(nTgl0), 54))))
    
    '--------------------------------------------------------------------------------------------------
    '  28 JUNI 2016   |   Rev. Penambahan Menu KPI LeadTimeINV & Rev. INF Billing Rekap    |   ARIKLD2
    '--------------------------------------------------------------------------------------------------
    Dim rsSum As ADODB.Recordset, rsSumR As ADODB.Recordset
    Set rsSum = objKp.FetchSumLeadTimeInv(Month(nTgl0), year(nTgl0), "LIGHT", "")
    Set rsSumR = objKp.FetchSumLeadTimeInvR(Month(nTgl0), year(nTgl0), "LIGHT", "")
    data(168) = SetBreakdownTime(NulNum(rsSum.Fields(2).value) - NulNum(rsSumR.Fields(2).value))
    data(175) = SetBreakdownTime(NulNum(Val(NulNum(rsSum.Fields(2).value) - NulNum(rsSumR.Fields(2).value)) / Val(objR.Data_Rekap(Month(nTgl0), year(nTgl0), 53))))
    Set rsSum = objKp.FetchSumLeadTimeInv(Month(nTgl0), year(nTgl0), "MEDIUM", "")
    Set rsSumR = objKp.FetchSumLeadTimeInvR(Month(nTgl0), year(nTgl0), "MEDIUM", "")
    data(169) = SetBreakdownTime(NulNum(rsSum.Fields(2).value) - NulNum(rsSumR.Fields(2).value))
    data(176) = SetBreakdownTime(NulNum(Val(NulNum(rsSum.Fields(2).value) - NulNum(rsSumR.Fields(2).value)) / Val(objR.Data_Rekap(Month(nTgl0), year(nTgl0), 54))))

    Report.Text45.SetText Val(data(45)) + Val(data(46))
    Report.Text46.SetText "0"
    Report.Text47.SetText data(47)
    Report.Text48.SetText data(48)
    Report.Text49.SetText data(49)
    Report.Text50.SetText data(50)
    Report.Text51.SetText data(51)
    Report.Text52.SetText data(52)
    Report.Text53.SetText data(53)
    Report.Text54.SetText data(54)
    Report.Text55.SetText data(55)
    Report.Text56.SetText data(56)
    Report.Text57.SetText data(57)
    Report.Text58.SetText Format(Round(data(58), 0), "#,###") & ",00"
    Report.Text59.SetText Format(Round(data(59), 0), "#,###") & ",00"
    Report.Text60.SetText Format(Round(data(60), 0), "#,###") & ",00"
    Report.Text61.SetText Format(Round(data(61), 0), "#,###") & ",00"
    Report.Text62.SetText Format(Round(data(62), 0), "#,###") & ",00"
    Report.Text63.SetText Format(Round(data(63), 0), "#,###") & ",00"
    Report.Text64.SetText Format(Round(data(64), 0), "#,###") & ",00"
    Report.Text65.SetText Format(Round(data(65), 0), "#,###") & ",00"
    Report.Text66.SetText Format(Round(data(66), 0), "#,###") & ",00"
    Report.Text67.SetText Format(Round(data(67), 0), "#,###") & ",00"
    Report.Text68.SetText Format(Round(data(68), 0), "#,###") & ",00"
    
    Report.Text125.SetText data(125)
    Report.Text126.SetText data(126)
    Report.Text127.SetText data(127)
    Report.Text128.SetText data(128)
    Report.Text129.SetText data(129)
    Report.Text130.SetText data(130)
    Report.Text131.SetText data(131)
    Report.Text132.SetText data(132)
    Report.Text133.SetText data(133)
    Report.Text134.SetText data(134)
    Report.Text135.SetText data(135)
    Report.Text136.SetText data(136)
    Report.Text137.SetText data(137)
    Report.Text138.SetText data(138)
    Report.Text139.SetText data(139)
    Report.Text140.SetText data(140)
    Report.Text141.SetText data(141)
    Report.Text142.SetText data(142)
    Report.Text143.SetText data(143)
    Report.Text144.SetText data(144)
    Report.Text145.SetText data(145)
    Report.Text146.SetText data(146)
    Report.Text147.SetText data(147)
    Report.Text148.SetText data(148)
    Report.Text149.SetText data(149)
    Report.Text150.SetText data(150)
    Report.Text151.SetText data(151)
    Report.Text152.SetText data(152)
    Report.Text153.SetText data(153)
    Report.Text154.SetText data(154)
    Report.Text155.SetText data(155)
    Report.Text156.SetText data(156)
    Report.Text157.SetText data(157)
    Report.Text158.SetText data(158)
    Report.Text159.SetText data(159)
    Report.Text160.SetText data(160)
    Report.Text161.SetText data(161)
    Report.Text162.SetText data(162)
    Report.Text163.SetText data(163)
    Report.Text164.SetText data(164)
    Report.Text165.SetText data(165)
    Report.Text166.SetText data(166)
    Report.Text166i.SetText data(16)
    Report.Text167.SetText data(167)
    Report.Text167i.SetText data(17)
    Report.Text168.SetText data(168)
    Report.Text168i.SetText objR.Data_Rekap(Month(nTgl0), year(nTgl0), 53) & " Unit, Total ="
    Report.Text169.SetText data(169)
    Report.Text169i.SetText objR.Data_Rekap(Month(nTgl0), year(nTgl0), 54) & " Unit, Total ="
    Report.Text175.SetText data(175)
    Report.Text176.SetText data(176)
    
    Call SetReport
End Sub
'====================================================================================
'====================================================================================
'====================================================================================
Private Sub SetReport()
    CRViewer.ReportSource = Report
    CRViewer.ViewReport
    CRViewer.Zoom 100
    Set oRS = Nothing
End Sub

Private Function SetBreakdownTime(nWaktu As Double) As String
    If nWaktu = "0" Then
        SetBreakdownTime = "00 : 00"
        Exit Function
    End If
    Dim xSq As Double
    Dim xMq As Double
    Dim xHq As Double
    Dim xDq As Double
    
    xSq = nWaktu Mod 60
    nWaktu = (nWaktu - xSq) / 60
    xMq = nWaktu Mod 60
    nWaktu = (nWaktu - xMq) / 60
    xHq = nWaktu Mod 24
    nWaktu = (nWaktu - xHq) / 24
    xDq = nWaktu

    SetBreakdownTime = Format(Trim((xDq * 24) + xHq), "###,#") & " : " & IIf(Len(Trim(xMq)) = 2, xMq, "0" + Trim(xMq)) '& " : " & xSq

End Function

Private Sub ViewCroz(n As Integer)
    Select Case n
    Case 0
        Report.X0.TextColor = &HC0&
        Report.X0.Font.Size = 11
    Case 1
        Report.X1.TextColor = &HC0&
        Report.X1.Font.Size = 11
    Case 2
        Report.X2.TextColor = &HC0&
        Report.X2.Font.Size = 11
    Case 3
        Report.X3.TextColor = &HC0&
        Report.X3.Font.Size = 11
    Case 4
        Report.X4.TextColor = &HC0&
        Report.X4.Font.Size = 11
    Case 5
        Report.X5.TextColor = &HC0&
        Report.X5.Font.Size = 11
    Case 6
        Report.X6.TextColor = &HC0&
        Report.X6.Font.Size = 11
    Case 7
        Report.X7.TextColor = &HC0&
        Report.X7.Font.Size = 11
    Case 8
        Report.X8.TextColor = &HC0&
        Report.X8.Font.Size = 11
    Case 9
        Report.X9.TextColor = &HC0&
        Report.X9.Font.Size = 11
    Case 10
        Report.X10.TextColor = &HC0&
        Report.X10.Font.Size = 11
    Case 11
        Report.X11.TextColor = &HC0&
        Report.X11.Font.Size = 11
    Case 12
        Report.X12.TextColor = &HC0&
        Report.X12.Font.Size = 11
    Case 13
        Report.X13.TextColor = &HC0&
        Report.X13.Font.Size = 11
    Case 14
        Report.X14.TextColor = &HC0&
        Report.X14.Font.Size = 11
    Case 15
        Report.X15.TextColor = &HC0&
        Report.X15.Font.Size = 11
    Case 16
        Report.X16.TextColor = &HC0&
        Report.X16.Font.Size = 11
    Case 17
        Report.X17.TextColor = &HC0&
        Report.X17.Font.Size = 11
    Case 18
        Report.X18.TextColor = &HC0&
        Report.X18.Font.Size = 11
    Case 19
        Report.X19.TextColor = &HC0&
        Report.X19.Font.Size = 11
    Case 20
        Report.X20.TextColor = &HC0&
        Report.X20.Font.Size = 11
    Case 21
        Report.X21.TextColor = &HC0&
        Report.X21.Font.Size = 11
    Case 22
        Report.X22.TextColor = &HC0&
        Report.X22.Font.Size = 11
    Case 23
        Report.X23.TextColor = &HC0&
        Report.X23.Font.Size = 11
    Case 24
        Report.X24.TextColor = &HC0&
        Report.X24.Font.Size = 11
    Case 25
        Report.X25.TextColor = &HC0&
        Report.X25.Font.Size = 11
    Case 26
        Report.X26.TextColor = &HC0&
        Report.X26.Font.Size = 11
    Case 27
        Report.X27.TextColor = &HC0&
        Report.X27.Font.Size = 11
    Case 28
        Report.X28.TextColor = &HC0&
        Report.X28.Font.Size = 11
    Case 29
        Report.X29.TextColor = &HC0&
        Report.X29.Font.Size = 11
    Case 30
        Report.X30.TextColor = &HC0&
        Report.X30.Font.Size = 11
    Case 31
        Report.X31.TextColor = &HC0&
        Report.X31.Font.Size = 11
    Case 32
        Report.X32.TextColor = &HC0&
        Report.X32.Font.Size = 11
    Case 33
        Report.X33.TextColor = &HC0&
        Report.X33.Font.Size = 11
    Case 34
        Report.X34.TextColor = &HC0&
        Report.X34.Font.Size = 11
    Case 35
        Report.X35.TextColor = &HC0&
        Report.X35.Font.Size = 11
    Case 36
        Report.X36.TextColor = &HC0&
        Report.X36.Font.Size = 11
    Case 37
        Report.X37.TextColor = &HC0&
        Report.X37.Font.Size = 11
    Case 38
        Report.X38.TextColor = &HC0&
        Report.X38.Font.Size = 11
    Case 39
        Report.X39.TextColor = &HC0&
        Report.X39.Font.Size = 11
    Case 40
        Report.X40.TextColor = &HC0&
        Report.X40.Font.Size = 11
    Case 41
        Report.X41.TextColor = &HC0&
        Report.X41.Font.Size = 11
    Case 42
        Report.X42.TextColor = &HC0&
        Report.X42.Font.Size = 11
    Case 43
        Report.X43.TextColor = &HC0&
        Report.X43.Font.Size = 11
    Case 44
        Report.X44.TextColor = &HC0&
        Report.X44.Font.Size = 11
    Case 45
        Report.X45.TextColor = &HC0&
        Report.X45.Font.Size = 11
    Case 46
        Report.X46.TextColor = &HC0&
        Report.X46.Font.Size = 11
    Case 47
        Report.X47.TextColor = &HC0&
        Report.X47.Font.Size = 11
    Case 48
        Report.X48.TextColor = &HC0&
        Report.X48.Font.Size = 11
    Case 49
        Report.X49.TextColor = &HC0&
        Report.X49.Font.Size = 11
    Case 50
        Report.X50.TextColor = &HC0&
        Report.X50.Font.Size = 11
    Case 51
        Report.X51.TextColor = &HC0&
        Report.X51.Font.Size = 11
    Case 52
        Report.X52.TextColor = &HC0&
        Report.X52.Font.Size = 11
    Case 53
        Report.X53.TextColor = &HC0&
        Report.X53.Font.Size = 11
    Case 54
        Report.X54.TextColor = &HC0&
        Report.X54.Font.Size = 11
    Case 55
        Report.X55.TextColor = &HC0&
        Report.X55.Font.Size = 11
    Case 56
        Report.X56.TextColor = &HC0&
        Report.X56.Font.Size = 11
    Case 57
        Report.X57.TextColor = &HC0&
        Report.X57.Font.Size = 11
    Case 58
        Report.X58.TextColor = &HC0&
        Report.X58.Font.Size = 11
    Case 59
        Report.X59.TextColor = &HC0&
        Report.X59.Font.Size = 11
    Case 60
        Report.X60.TextColor = &HC0&
        Report.X60.Font.Size = 11
    Case 61
        Report.X61.TextColor = &HC0&
        Report.X61.Font.Size = 11
    Case 62
        Report.X62.TextColor = &HC0&
        Report.X62.Font.Size = 11
    Case 63
        Report.X63.TextColor = &HC0&
        Report.X63.Font.Size = 11
    Case 64
        Report.X64.TextColor = &HC0&
        Report.X64.Font.Size = 11
    Case 65
        Report.X65.TextColor = &HC0&
        Report.X65.Font.Size = 11
    Case 66
        Report.X66.TextColor = &HC0&
        Report.X66.Font.Size = 11
    Case 67
        Report.X67.TextColor = &HC0&
        Report.X67.Font.Size = 11
    Case 68
        Report.X68.TextColor = &HC0&
        Report.X68.Font.Size = 11
    Case 69
        Report.X69.TextColor = &HC0&
        Report.X69.Font.Size = 11
    Case 70
        Report.X70.TextColor = &HC0&
        Report.X70.Font.Size = 11
    Case 71
        Report.X71.TextColor = &HC0&
        Report.X71.Font.Size = 11
    Case 72
        Report.X72.TextColor = &HC0&
        Report.X72.Font.Size = 11
    Case 73
        Report.X73.TextColor = &HC0&
        Report.X73.Font.Size = 11
    Case 74
        Report.X74.TextColor = &HC0&
        Report.X74.Font.Size = 11
    Case 75
        Report.X75.TextColor = &HC0&
        Report.X75.Font.Size = 11
    Case 76
        Report.X76.TextColor = &HC0&
        Report.X76.Font.Size = 11
    Case 77
        Report.X77.TextColor = &HC0&
        Report.X77.Font.Size = 11
    Case 78
        Report.X78.TextColor = &HC0&
        Report.X78.Font.Size = 11
    Case 79
        Report.X79.TextColor = &HC0&
        Report.X79.Font.Size = 11
    Case 80
        Report.X80.TextColor = &HC0&
        Report.X80.Font.Size = 11
    Case 81
        Report.X81.TextColor = &HC0&
        Report.X81.Font.Size = 11
    Case 82
        Report.X82.TextColor = &HC0&
        Report.X82.Font.Size = 11
    Case 83
        Report.X83.TextColor = &HC0&
        Report.X83.Font.Size = 11
    Case 84
        Report.X84.TextColor = &HC0&
        Report.X84.Font.Size = 11
    Case 85
        Report.X85.TextColor = &HC0&
        Report.X85.Font.Size = 11
    Case 86
        Report.X86.TextColor = &HC0&
        Report.X86.Font.Size = 11
    Case 87
        Report.X87.TextColor = &HC0&
        Report.X87.Font.Size = 11
    Case 88
        Report.X88.TextColor = &HC0&
        Report.X88.Font.Size = 11
    Case 89
        Report.X89.TextColor = &HC0&
        Report.X89.Font.Size = 11
    Case 90
        Report.X90.TextColor = &HC0&
        Report.X90.Font.Size = 11
    Case 91
        Report.X91.TextColor = &HC0&
        Report.X91.Font.Size = 11
    Case 92
        Report.X92.TextColor = &HC0&
        Report.X92.Font.Size = 11
    Case 93
        Report.X93.TextColor = &HC0&
        Report.X93.Font.Size = 11
    Case 94
        Report.X94.TextColor = &HC0&
        Report.X94.Font.Size = 11
    Case 95
        Report.X95.TextColor = &HC0&
        Report.X95.Font.Size = 11
    Case 96
        Report.X96.TextColor = &HC0&
        Report.X96.Font.Size = 11
    Case 97
        Report.X97.TextColor = &HC0&
        Report.X97.Font.Size = 11
    Case 98
        Report.X98.TextColor = &HC0&
        Report.X98.Font.Size = 11
    Case 99
        Report.X99.TextColor = &HC0&
        Report.X99.Font.Size = 11
    Case 100
        Report.X100.TextColor = &HC0&
        Report.X100.Font.Size = 11
    Case 101
        Report.X101.TextColor = &HC0&
        Report.X101.Font.Size = 11
    Case 102
        Report.X102.TextColor = &HC0&
        Report.X102.Font.Size = 11
    Case 103
        Report.X103.TextColor = &HC0&
        Report.X103.Font.Size = 11
    Case 104
        Report.X104.TextColor = &HC0&
        Report.X104.Font.Size = 11
    Case 105
        Report.X105.TextColor = &HC0&
        Report.X105.Font.Size = 11
    Case 106
        Report.X106.TextColor = &HC0&
        Report.X106.Font.Size = 11
    Case 107
        Report.X107.TextColor = &HC0&
        Report.X107.Font.Size = 11
    Case 108
        Report.X108.TextColor = &HC0&
        Report.X108.Font.Size = 11
    Case 109
        Report.X109.TextColor = &HC0&
        Report.X109.Font.Size = 11
    Case 110
        Report.X110.TextColor = &HC0&
        Report.X110.Font.Size = 11
    Case 111
        Report.X111.TextColor = &HC0&
        Report.X111.Font.Size = 11
    Case 112
        Report.X112.TextColor = &HC0&
        Report.X112.Font.Size = 11
    Case 113
        Report.X113.TextColor = &HC0&
        Report.X113.Font.Size = 11
    Case 114
        Report.X114.TextColor = &HC0&
        Report.X114.Font.Size = 11
    Case 115
        Report.X115.TextColor = &HC0&
        Report.X115.Font.Size = 11
    Case 116
        Report.X116.TextColor = &HC0&
        Report.X116.Font.Size = 11
    Case 117
        Report.X117.TextColor = &HC0&
        Report.X117.Font.Size = 11
    Case 118
        Report.X118.TextColor = &HC0&
        Report.X118.Font.Size = 11
    Case 119
        Report.X119.TextColor = &HC0&
        Report.X119.Font.Size = 11
    Case 120
        Report.X120.TextColor = &HC0&
        Report.X120.Font.Size = 11
    Case 121
        Report.X121.TextColor = &HC0&
        Report.X121.Font.Size = 11
    Case 122
        Report.X122.TextColor = &HC0&
        Report.X122.Font.Size = 11
    Case 123
        Report.X123.TextColor = &HC0&
        Report.X123.Font.Size = 11
    Case 124
        Report.X124.TextColor = &HC0&
        Report.X124.Font.Size = 11
    Case 125
        Report.X125.TextColor = &HC0&
        Report.X125.Font.Size = 11
    Case 126
        Report.X126.TextColor = &HC0&
        Report.X126.Font.Size = 11
    Case 127
        Report.X127.TextColor = &HC0&
        Report.X127.Font.Size = 11
    Case 128
        Report.X128.TextColor = &HC0&
        Report.X128.Font.Size = 11
    Case 129
        Report.X129.TextColor = &HC0&
        Report.X129.Font.Size = 11
    Case 130
        Report.X130.TextColor = &HC0&
        Report.X130.Font.Size = 11
    Case 131
        Report.X131.TextColor = &HC0&
        Report.X131.Font.Size = 11
    Case 132
        Report.X132.TextColor = &HC0&
        Report.X132.Font.Size = 11
    Case 133
        Report.X133.TextColor = &HC0&
        Report.X133.Font.Size = 11
    Case 134
        Report.X134.TextColor = &HC0&
        Report.X134.Font.Size = 11
    Case 135
        Report.X135.TextColor = &HC0&
        Report.X135.Font.Size = 11
    Case 136
        Report.X136.TextColor = &HC0&
        Report.X136.Font.Size = 11
    Case 137
        Report.X137.TextColor = &HC0&
        Report.X137.Font.Size = 11
    Case 138
        Report.X138.TextColor = &HC0&
        Report.X138.Font.Size = 11
    Case 139
        Report.X139.TextColor = &HC0&
        Report.X139.Font.Size = 11
    Case 140
        Report.X140.TextColor = &HC0&
        Report.X140.Font.Size = 11
    Case 141
        Report.X141.TextColor = &HC0&
        Report.X141.Font.Size = 11
    Case 142
        Report.X142.TextColor = &HC0&
        Report.X142.Font.Size = 11
    Case 143
        Report.X143.TextColor = &HC0&
        Report.X143.Font.Size = 11
    Case 144
        Report.X144.TextColor = &HC0&
        Report.X144.Font.Size = 11
    Case 145
        Report.X145.TextColor = &HC0&
        Report.X145.Font.Size = 11
    Case 146
        Report.X146.TextColor = &HC0&
        Report.X146.Font.Size = 11
    Case 147
        Report.X147.TextColor = &HC0&
        Report.X147.Font.Size = 11
    Case 148
        Report.X148.TextColor = &HC0&
        Report.X148.Font.Size = 11
    Case 149
        Report.X149.TextColor = &HC0&
        Report.X149.Font.Size = 11
    Case 150
        Report.X150.TextColor = &HC0&
        Report.X150.Font.Size = 11
    Case 151
        Report.X151.TextColor = &HC0&
        Report.X151.Font.Size = 11
    Case 152
        Report.X152.TextColor = &HC0&
        Report.X152.Font.Size = 11
    Case 153
        Report.X153.TextColor = &HC0&
        Report.X153.Font.Size = 11
    Case 154
        Report.X154.TextColor = &HC0&
        Report.X154.Font.Size = 11
    Case 155
        Report.X155.TextColor = &HC0&
        Report.X155.Font.Size = 11
    Case 156
        Report.X156.TextColor = &HC0&
        Report.X156.Font.Size = 11
    Case 157
        Report.X157.TextColor = &HC0&
        Report.X157.Font.Size = 11
    Case 158
        Report.X158.TextColor = &HC0&
        Report.X158.Font.Size = 11
    Case 159
        Report.X159.TextColor = &HC0&
        Report.X159.Font.Size = 11
    Case 160
        Report.X160.TextColor = &HC0&
        Report.X160.Font.Size = 11
    Case 161
        Report.X161.TextColor = &HC0&
        Report.X161.Font.Size = 11
    Case 162
        Report.X162.TextColor = &HC0&
        Report.X162.Font.Size = 11
    Case 163
        Report.X163.TextColor = &HC0&
        Report.X163.Font.Size = 11
    Case 164
        Report.X164.TextColor = &HC0&
        Report.X164.Font.Size = 11
    Case 165
        Report.X165.TextColor = &HC0&
        Report.X165.Font.Size = 11
    Case 166
        Report.X166.TextColor = &HC0&
        Report.X166.Font.Size = 11
    Case 167
        Report.X167.TextColor = &HC0&
        Report.X167.Font.Size = 11
    Case 168
        Report.X168.TextColor = &HC0&
        Report.X168.Font.Size = 11
    Case 169
        Report.X169.TextColor = &HC0&
        Report.X169.Font.Size = 11
    Case 170
        Report.X170.TextColor = &HC0&
        Report.X170.Font.Size = 11
    Case 171
        Report.X171.TextColor = &HC0&
        Report.X171.Font.Size = 11
    Case 172
        Report.X172.TextColor = &HC0&
        Report.X172.Font.Size = 11
    Case 173
        Report.X173.TextColor = &HC0&
        Report.X173.Font.Size = 11
    Case 174
        Report.X174.TextColor = &HC0&
        Report.X174.Font.Size = 11
    Case 175
        Report.X175.TextColor = &HC0&
        Report.X175.Font.Size = 11
    Case 176
        Report.X176.TextColor = &HC0&
        Report.X176.Font.Size = 11
    Case 177
        Report.X177.TextColor = &HC0&
        Report.X177.Font.Size = 11
    Case 178
        Report.X178.TextColor = &HC0&
        Report.X178.Font.Size = 11
    Case 179
        Report.X179.TextColor = &HC0&
        Report.X179.Font.Size = 11
    Case 180
        Report.X180.TextColor = &HC0&
        Report.X180.Font.Size = 11
    Case 181
        Report.X181.TextColor = &HC0&
        Report.X181.Font.Size = 11
    Case 182
        Report.X182.TextColor = &HC0&
        Report.X182.Font.Size = 11
    Case 183
        Report.X183.TextColor = &HC0&
        Report.X183.Font.Size = 11
    Case 184
        Report.X184.TextColor = &HC0&
        Report.X184.Font.Size = 11
    Case 185
        Report.X185.TextColor = &HC0&
        Report.X185.Font.Size = 11
    Case 186
        Report.X186.TextColor = &HC0&
        Report.X186.Font.Size = 11
    Case 187
        Report.X187.TextColor = &HC0&
        Report.X187.Font.Size = 11
    Case 188
        Report.X188.TextColor = &HC0&
        Report.X188.Font.Size = 11
    Case 189
        Report.X189.TextColor = &HC0&
        Report.X189.Font.Size = 11
    Case 190
        Report.X190.TextColor = &HC0&
        Report.X190.Font.Size = 11
    Case 191
        Report.X191.TextColor = &HC0&
        Report.X191.Font.Size = 11
    Case 192
        Report.X192.TextColor = &HC0&
        Report.X192.Font.Size = 11
    Case 193
        Report.X193.TextColor = &HC0&
        Report.X193.Font.Size = 11
    Case 194
        Report.X194.TextColor = &HC0&
        Report.X194.Font.Size = 11
    Case 195
        Report.X195.TextColor = &HC0&
        Report.X195.Font.Size = 11
    Case 196
        Report.X196.TextColor = &HC0&
        Report.X196.Font.Size = 11
    Case 197
        Report.X197.TextColor = &HC0&
        Report.X197.Font.Size = 11
    Case 198
        Report.X198.TextColor = &HC0&
        Report.X198.Font.Size = 11
    Case 199
        Report.X199.TextColor = &HC0&
        Report.X199.Font.Size = 11
    Case 200
        Report.X200.TextColor = &HC0&
        Report.X200.Font.Size = 11
    Case 201
        Report.X201.TextColor = &HC0&
        Report.X201.Font.Size = 11
    Case 202
        Report.X202.TextColor = &HC0&
        Report.X202.Font.Size = 11
    Case 203
        Report.X203.TextColor = &HC0&
        Report.X203.Font.Size = 11
    Case 204
        Report.X204.TextColor = &HC0&
        Report.X204.Font.Size = 11
    Case 205
        Report.X205.TextColor = &HC0&
        Report.X205.Font.Size = 11
    Case 206
        Report.X206.TextColor = &HC0&
        Report.X206.Font.Size = 11
    Case 207
        Report.X207.TextColor = &HC0&
        Report.X207.Font.Size = 11
    Case 208
        Report.X208.TextColor = &HC0&
        Report.X208.Font.Size = 11
    Case 209
        Report.X209.TextColor = &HC0&
        Report.X209.Font.Size = 11
    Case 210
        Report.X210.TextColor = &HC0&
        Report.X210.Font.Size = 11
    Case 211
        Report.X211.TextColor = &HC0&
        Report.X211.Font.Size = 11
    Case 212
        Report.X212.TextColor = &HC0&
        Report.X212.Font.Size = 11
    Case 213
        Report.X213.TextColor = &HC0&
        Report.X213.Font.Size = 11
    Case 214
        Report.X214.TextColor = &HC0&
        Report.X214.Font.Size = 11
    Case 215
        Report.X215.TextColor = &HC0&
        Report.X215.Font.Size = 11
    Case 216
        Report.X216.TextColor = &HC0&
        Report.X216.Font.Size = 11
    Case 217
        Report.X217.TextColor = &HC0&
        Report.X217.Font.Size = 11
    Case 218
        Report.X218.TextColor = &HC0&
        Report.X218.Font.Size = 11
    Case 219
        Report.X219.TextColor = &HC0&
        Report.X219.Font.Size = 11
    Case 220
        Report.X220.TextColor = &HC0&
        Report.X220.Font.Size = 11
    Case 221
        Report.X221.TextColor = &HC0&
        Report.X221.Font.Size = 11
    Case 222
        Report.X222.TextColor = &HC0&
        Report.X222.Font.Size = 11
    Case 223
        Report.X223.TextColor = &HC0&
        Report.X223.Font.Size = 11
    Case 224
        Report.X224.TextColor = &HC0&
        Report.X224.Font.Size = 11
    Case 225
        Report.X225.TextColor = &HC0&
        Report.X225.Font.Size = 11
    Case 226
        Report.X226.TextColor = &HC0&
        Report.X226.Font.Size = 11
    Case 227
        Report.X227.TextColor = &HC0&
        Report.X227.Font.Size = 11
    Case 228
        Report.X228.TextColor = &HC0&
        Report.X228.Font.Size = 11
    Case 229
        Report.X229.TextColor = &HC0&
        Report.X229.Font.Size = 11
    Case 230
        Report.X230.TextColor = &HC0&
        Report.X230.Font.Size = 11
    Case 231
        Report.X231.TextColor = &HC0&
        Report.X231.Font.Size = 11
    Case 232
        Report.X232.TextColor = &HC0&
        Report.X232.Font.Size = 11
    Case 233
        Report.X233.TextColor = &HC0&
        Report.X233.Font.Size = 11
    Case 234
        Report.X234.TextColor = &HC0&
        Report.X234.Font.Size = 11
    Case 235
        Report.X235.TextColor = &HC0&
        Report.X235.Font.Size = 11
    Case 236
        Report.X236.TextColor = &HC0&
        Report.X236.Font.Size = 11
    Case 237
        Report.X237.TextColor = &HC0&
        Report.X237.Font.Size = 11
    Case 238
        Report.X238.TextColor = &HC0&
        Report.X238.Font.Size = 11
    Case 239
        Report.X239.TextColor = &HC0&
        Report.X239.Font.Size = 11
    Case 240
        Report.X240.TextColor = &HC0&
        Report.X240.Font.Size = 11
    Case 241
        Report.X241.TextColor = &HC0&
        Report.X241.Font.Size = 11
    Case 242
        Report.X242.TextColor = &HC0&
        Report.X242.Font.Size = 11
    Case 243
        Report.X243.TextColor = &HC0&
        Report.X243.Font.Size = 11
    Case 244
        Report.X244.TextColor = &HC0&
        Report.X244.Font.Size = 11
    Case 245
        Report.X245.TextColor = &HC0&
        Report.X245.Font.Size = 11
    Case 246
        Report.X246.TextColor = &HC0&
        Report.X246.Font.Size = 11
    Case 247
        Report.X247.TextColor = &HC0&
        Report.X247.Font.Size = 11
    Case 248
        Report.X248.TextColor = &HC0&
        Report.X248.Font.Size = 11
    Case 249
        Report.X249.TextColor = &HC0&
        Report.X249.Font.Size = 11
    Case 250
        Report.X250.TextColor = &HC0&
        Report.X250.Font.Size = 11
    Case 251
        Report.X251.TextColor = &HC0&
        Report.X251.Font.Size = 11
    Case 252
        Report.X252.TextColor = &HC0&
        Report.X252.Font.Size = 11
    Case 253
        Report.X253.TextColor = &HC0&
        Report.X253.Font.Size = 11
    Case 254
        Report.X254.TextColor = &HC0&
        Report.X254.Font.Size = 11
    Case 255
        Report.X255.TextColor = &HC0&
        Report.X255.Font.Size = 11
    Case 256
        Report.X256.TextColor = &HC0&
        Report.X256.Font.Size = 11
    Case 257
        Report.X257.TextColor = &HC0&
        Report.X257.Font.Size = 11
    Case 258
        Report.X258.TextColor = &HC0&
        Report.X258.Font.Size = 11
    Case 259
        Report.X259.TextColor = &HC0&
        Report.X259.Font.Size = 11
    Case 260
        Report.X260.TextColor = &HC0&
        Report.X260.Font.Size = 11
    Case 261
        Report.X261.TextColor = &HC0&
        Report.X261.Font.Size = 11
    Case 262
        Report.X262.TextColor = &HC0&
        Report.X262.Font.Size = 11
    Case 263
        Report.X263.TextColor = &HC0&
        Report.X263.Font.Size = 11
    Case 264
        Report.X264.TextColor = &HC0&
        Report.X264.Font.Size = 11
    Case 265
        Report.X265.TextColor = &HC0&
        Report.X265.Font.Size = 11
    Case 266
        Report.X266.TextColor = &HC0&
        Report.X266.Font.Size = 11
    Case 267
        Report.X267.TextColor = &HC0&
        Report.X267.Font.Size = 11
    Case 268
        Report.X268.TextColor = &HC0&
        Report.X268.Font.Size = 11
    Case 269
        Report.X269.TextColor = &HC0&
        Report.X269.Font.Size = 11
    Case 270
        Report.X270.TextColor = &HC0&
        Report.X270.Font.Size = 11
    Case 271
        Report.X271.TextColor = &HC0&
        Report.X271.Font.Size = 11
    Case 272
        Report.X272.TextColor = &HC0&
        Report.X272.Font.Size = 11
    Case 273
        Report.X273.TextColor = &HC0&
        Report.X273.Font.Size = 11
    Case 274
        Report.X274.TextColor = &HC0&
        Report.X274.Font.Size = 11
    Case 275
        Report.X275.TextColor = &HC0&
        Report.X275.Font.Size = 11
    Case 276
        Report.X276.TextColor = &HC0&
        Report.X276.Font.Size = 11
    Case 277
        Report.X277.TextColor = &HC0&
        Report.X277.Font.Size = 11
    Case 278
        Report.X278.TextColor = &HC0&
        Report.X278.Font.Size = 11
    Case 279
        Report.X279.TextColor = &HC0&
        Report.X279.Font.Size = 11
    Case 280
        Report.X280.TextColor = &HC0&
        Report.X280.Font.Size = 11
    Case 281
        Report.X281.TextColor = &HC0&
        Report.X281.Font.Size = 11
    Case 282
        Report.X282.TextColor = &HC0&
        Report.X282.Font.Size = 11
    Case 283
        Report.X283.TextColor = &HC0&
        Report.X283.Font.Size = 11
    Case 284
        Report.X284.TextColor = &HC0&
        Report.X284.Font.Size = 11
    Case 285
        Report.X285.TextColor = &HC0&
        Report.X285.Font.Size = 11
    Case 286
        Report.X286.TextColor = &HC0&
        Report.X286.Font.Size = 11
    Case 287
        Report.X287.TextColor = &HC0&
        Report.X287.Font.Size = 11
    Case 288
        Report.X288.TextColor = &HC0&
        Report.X288.Font.Size = 11
    Case 289
        Report.X289.TextColor = &HC0&
        Report.X289.Font.Size = 11
    Case 290
        Report.X290.TextColor = &HC0&
        Report.X290.Font.Size = 11
    Case 291
        Report.X291.TextColor = &HC0&
        Report.X291.Font.Size = 11
    Case 292
        Report.X292.TextColor = &HC0&
        Report.X292.Font.Size = 11
    Case 293
        Report.X293.TextColor = &HC0&
        Report.X293.Font.Size = 11
    Case 294
        Report.X294.TextColor = &HC0&
        Report.X294.Font.Size = 11
    End Select
End Sub

Private Sub GetPrinter()
    Dim vPort As String, vDriver As String
    CboPrint.Visible = True
    Set WmiObjSet = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery("SELECT * FROM Win32_Printer")
    With CboPrint
        .Clear
        If WmiObjSet.Count > 0 Then
            For Each Wmiobj In WmiObjSet
                vPort = "": vDriver = ""
                vPort = UCase(Left(Trim(Wmiobj.PortName), 3)): vDriver = UCase(Trim(Wmiobj.DriverName))
                If (vPort = "USB" Or vPort = "LPT" Or vPort = "COM" Or vPort = "DOT" Or vPort = "ESD") And _
                    InStr(vDriver, "ADOBE") = 0 And InStr(vDriver, "ACROBAT") = 0 And InStr(vDriver, "PDF") = 0 Then
                        .AddItem Trim(Wmiobj.DeviceID)
                End If
            Next
        End If
    End With
End Sub

