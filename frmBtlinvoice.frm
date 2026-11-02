VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmBtlinvoice 
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Alasan Retur : "
      Height          =   2085
      Left            =   375
      TabIndex        =   2
      Top             =   375
      Width           =   4080
      Begin VB.CommandButton TblProses 
         Caption         =   "&Proses"
         Height          =   390
         Left            =   1830
         TabIndex        =   1
         Top             =   1620
         Width           =   2145
      End
      Begin VB.TextBox Text1 
         Height          =   1260
         Left            =   120
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   3840
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   4471
      StartColor      =   0
      EndColor        =   12632064
      FormatString    =   "Retur Invoice"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   345
      Top             =   45
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxSysButton BtnTutup 
      Height          =   285
      Left            =   4140
      Top             =   45
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      ButtonStyle     =   0
      ButtonFlags     =   16384
   End
End
Attribute VB_Name = "frmBtlinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vKodeLgn As String
Dim vNoRetur As String
Dim vJam As String

Private Sub HapusFormRetur()
    With frmBatalBilling
        .Text(0) = ""
        For i = 1 To 45
            Select Case i
                Case 9, 10, 11, 15, 18, 20, 22, 23, 27, 29, 30, 32, 33, 34, 35, 36, 43, 44
                Case 28
                    .Text(i) = ">>->>->>>>"
                Case Else
                    .Text(i) = ""
            End Select
        Next
        .Text1.Text = ""
        Baris = 1
        For i = 1 To 7
            .DGrid.TextMatrix(i, 1) = ""
            .DGrid.TextMatrix(i, 2) = ""
            .DGrid.TextMatrix(i, 3) = ""
            .DGrid.TextMatrix(i, 4) = ""
            .DGrid.TextMatrix(i, 5) = ""
        Next
        .Text(57) = ""
        .Text(0).SetFocus
    End With
End Sub

Private Sub CetakFormKecil()
    Dim rpt As New RptInvRetur
    
    With frmBatalBilling
        rpt.Label48 = "No Urut Retur : " & vNoRetur
        rpt.Label39 = gSecmanager
        rpt.Label40 = "KEPALA BENGKEL"
        rpt.Label2 = "NPWP : " & gNpwp
        rpt.Label3 = "PKP  : " & gPkp
        rpt.Label44 = "TELP : " & gTelp
        rpt.Label18 = "BIAYA JASA  : " & AddSpace(Format(.Text(39), "#,##0"), 13, 3, " ")
        rpt.Label19 = "SUBLET      : " & AddSpace(Format(.Text(5), "#,##0"), 13, 3, " ")
        rpt.Label20 = "PARTS       : " & AddSpace(Format(.Text(40), "#,##0"), 13, 3, " ")
        rpt.Label21 = "OIL/GRERASE : " & AddSpace(Format(.Text(45), "#,##0"), 13, 3, " ")
        rpt.Label22 = "MATERIAL    : " & AddSpace(Format(.Text(4), "#,##0"), 13, 3, " ")
        rpt.Label23 = "SUBTOTAL    : " & AddSpace(Format(.Text(6), "#,##0"), 13, 3, " ")
        rpt.Label24 = "METERAI     : " & AddSpace(Format(.Text(7), "#,##0"), 13, 3, " ")
        rpt.Label25 = "GRAND TOTAL : " & AddSpace(Format(.Text(41), "#,##0"), 13, 3, " ")
        
        rpt.Label26 = AddSpace("URAIAN BIAYA JASA", 25, 1, " ") + _
                    AddSpace("MKN", 4, 3, " ") + _
                    AddSpace("RUPIAH", 12, 3, " ")
        rpt.Label27 = AddSpace(.DGrid.TextMatrix(1, 1), 25, 1, " ") + _
                    AddSpace(.DGrid.TextMatrix(1, 5), 4, 3, " ") + _
                    AddSpace(Format(.DGrid.TextMatrix(1, 4), "#,##0"), 12, 3, " ")
        rpt.Label28 = AddSpace(.DGrid.TextMatrix(2, 1), 25, 1, " ") + _
                    AddSpace(.DGrid.TextMatrix(2, 5), 4, 3, " ") + _
                    AddSpace(Format(.DGrid.TextMatrix(2, 4), "#,##0"), 12, 3, " ")
        rpt.Label29 = AddSpace(.DGrid.TextMatrix(3, 1), 25, 1, " ") + _
                    AddSpace(.DGrid.TextMatrix(3, 5), 4, 3, " ") + _
                    AddSpace(Format(.DGrid.TextMatrix(3, 4), "#,##0"), 12, 3, " ")
        rpt.Label30 = AddSpace(.DGrid.TextMatrix(4, 1), 25, 1, " ") + _
                    AddSpace(.DGrid.TextMatrix(4, 5), 4, 3, " ") + _
                    AddSpace(Format(.DGrid.TextMatrix(4, 4), "#,##0"), 12, 3, " ")
        rpt.Label31 = AddSpace(.DGrid.TextMatrix(5, 1), 25, 1, " ") + _
                    AddSpace(.DGrid.TextMatrix(5, 5), 4, 3, " ") + _
                    AddSpace(Format(.DGrid.TextMatrix(5, 4), "#,##0"), 12, 3, " ")
        rpt.Label32 = AddSpace(.DGrid.TextMatrix(6, 1), 25, 1, " ") + _
                    AddSpace(.DGrid.TextMatrix(6, 5), 4, 3, " ") + _
                    AddSpace(Format(.DGrid.TextMatrix(6, 4), "#,##0"), 12, 3, " ")
        rpt.Label33 = AddSpace(.DGrid.TextMatrix(7, 1), 25, 1, " ") + _
                    AddSpace(.DGrid.TextMatrix(7, 5), 4, 3, " ") + _
                    AddSpace(Format(.DGrid.TextMatrix(7, 4), "#,##0"), 12, 3, " ")
        
        rpt.Label8 = "NO. WO. : " & .Text(26)
        rpt.Label9 = "FRAME : " & .Text(25)
        rpt.Label10 = "E/G : " & .Text(21)
        rpt.Label11 = "NAMA : " & .Text(1)
        rpt.Label12 = "ALAMAT : " & .Text(2)
        rpt.Label13 = "KOTA : " & .Text(3)
        rpt.Label14 = .Text(19) & "(" & .Text(24) & ")"
        rpt.Label15 = "SA : " & .Text(13)
        rpt.Label16 = "OPT : " & gSA
        
        rpt.Label47 = "Kode Plgn  : " & .Text(57)
        rpt.Label4 = "No. Retur  : " & gNoInvoice & gKodeRetur & "/" & .cboStatus.Text
        rpt.Label5 = "Tanggal    : " & gTanggal
        rpt.Label6 = "Jam        : " & vJam
        rpt.Label7 = "No. Polisi : " & .Text(16)
        
        rpt.Label36 = .Text1.Text
        rpt.Label37 = .Text(37)
        rpt.Label38 = .Text(38)
        
        rpt.Label45 = "Bank     : " & gBank
        rpt.Label46 = "Acc. No. : " & gAcc
    
        rpt.Printer.PaperWidth = 1440 * 8.5
        rpt.Printer.PaperHeight = 1440 * 5.5
        rpt.Printer.Orientation = 1
    End With
        
    If gPreviewCetak = True Then
        rpt.PrintReport False
        Unload rpt
    Else
        rpt.Show vbModal
    End If
End Sub

Private Sub Cetak()
    Dim rpt As New RptFromRetur
    
    With frmBatalBilling
        gKodePlg = .Text(57)
        gInvNo = .Text(42) & "/" & .cboStatus.Text
        gInvTgl = gTanggal
        gInvCust = .Text(1)
        gInvAlmt = .Text(2)
        gInvKota = .Text(3) & "-" & vNoTelp
        gInvNPWP = .Text(14)
        gInvNoPolisi = .Text(16)
        gInvModel = .Text(19) & " (" & .Text(24) & ")"
        gInvNoMesin = .Text(21)
        gInvNoChas = .Text(25)
        gInvNoOrder = .Text(26)
        gInvSA = .Text(13)
        gInvStatus = .cboStatus.Text
        gInvMech = .Text(31)
        gInvPrint = vJam
        gInvBilling = gSA
        gInvUraian1 = .DGrid.TextMatrix(1, 1)
        gInvUraian2 = .DGrid.TextMatrix(2, 1)
        gInvUraian3 = .DGrid.TextMatrix(3, 1)
        gInvUraian4 = .DGrid.TextMatrix(4, 1)
        gInvUraian5 = .DGrid.TextMatrix(5, 1)
        gInvUraian6 = .DGrid.TextMatrix(6, 1)
        gInvUraian7 = .DGrid.TextMatrix(7, 1)
        gInvRate1 = ""
        gInvRate2 = ""
        gInvRate3 = ""
        gInvRate4 = ""
        gInvRate5 = ""
        gInvRate6 = ""
        gInvRate7 = ""
        gInvRp1 = Format(.DGrid.TextMatrix(1, 4), "#,##0")
        gInvRp2 = Format(.DGrid.TextMatrix(2, 4), "#,##0")
        gInvRp3 = Format(.DGrid.TextMatrix(3, 4), "#,##0")
        gInvRp4 = Format(.DGrid.TextMatrix(4, 4), "#,##0")
        gInvRp5 = Format(.DGrid.TextMatrix(5, 4), "#,##0")
        gInvRp6 = Format(.DGrid.TextMatrix(6, 4), "#,##0")
        gInvRp7 = Format(.DGrid.TextMatrix(7, 4), "#,##0")
        gGInvRp = Format(.Text(39), "#,##0")
        gInvTRp1 = Format(.Text(39), "#,##0")
        gInvTRp2 = Format(.Text(5), "#,##0")
        gInvTRp3 = Format(.Text(40), "#,##0")
        gInvTRp4 = Format(.Text(45), "#,##0")
        gInvTRp5 = Format(.Text(4), "#,##0")
        gInvTRp6 = Format(.Text(6), "#,##0")
        gInvTRp7 = Format(.Text(7), "#,##0")
        gGInvTRp = Format(.Text(41), "#,##0")
        gInvTerbilang = Text1.Text
        gInvSS = .Text(37)
        gInvCatatan = .Text(38)
    
        gInvSM = gSecmanager
        rpt.Field67.DataValue = vNoRetur
        rpt.Field66.DataValue = gNoInvoice & gKodeRetur & "/" & .cboStatus.Text
        rpt.Field65.DataValue = gKodePlg
        rpt.Field39.DataValue = gInvNo
        rpt.Field40.DataValue = gInvTgl
    End With

    rpt.Field1.DataValue = gInvCust
    rpt.Field4.DataValue = gInvAlmt
    rpt.Field28.DataValue = gInvKota
    rpt.Field7.DataValue = gInvNPWP

    rpt.Field2.DataValue = gInvNoPolisi
    rpt.Field5.DataValue = gInvModel
    rpt.Field12.DataValue = gInvNoMesin
    rpt.Field14.DataValue = gInvNoChas
    
    rpt.Field29.DataValue = gInvNoOrder
    rpt.Field30.DataValue = gInvSA

    rpt.Field31.DataValue = gInvStatus
    rpt.Field32.DataValue = gInvMech

    rpt.Field33.DataValue = gInvPrint
    rpt.Field38.DataValue = gInvBilling
    
    rpt.Field16.DataValue = gInvUraian1
    rpt.Field17.DataValue = gInvUraian2
    rpt.Field18.DataValue = gInvUraian3
    rpt.Field19.DataValue = gInvUraian4
    rpt.Field20.DataValue = gInvUraian5
    rpt.Field21.DataValue = gInvUraian6
    rpt.Field22.DataValue = gInvUraian7
    
    rpt.Field41.DataValue = gInvRate1
    rpt.Field42.DataValue = gInvRate2
    rpt.Field43.DataValue = gInvRate3
    rpt.Field44.DataValue = gInvRate4
    rpt.Field45.DataValue = gInvRate5
    rpt.Field46.DataValue = gInvRate6
    rpt.Field47.DataValue = gInvRate7
    
    rpt.Field48.DataValue = gInvRp1
    rpt.Field49.DataValue = gInvRp2
    rpt.Field50.DataValue = gInvRp3
    rpt.Field51.DataValue = gInvRp4
    rpt.Field52.DataValue = gInvRp5
    rpt.Field53.DataValue = gInvRp6
    rpt.Field54.DataValue = gInvRp7
    
    rpt.Field55.DataValue = gInvTRp1
    rpt.Field56.DataValue = gInvTRp2
    rpt.Field57.DataValue = gInvTRp3
    rpt.Field58.DataValue = gInvTRp4
    rpt.Field59.DataValue = gInvTRp5
    rpt.Field60.DataValue = gInvTRp6
    rpt.Field61.DataValue = gInvTRp7
    
    rpt.Field23.DataValue = gGInvRp
    rpt.Field27.DataValue = gGInvTRp
    rpt.Field25.DataValue = gInvTerbilang
    
    rpt.Field62.DataValue = gInvSS
    rpt.Field63.DataValue = gInvCatatan
    rpt.Field64.DataValue = gInvSM
    '9=A4
    'kwarto

'    rpt.Printer.PaperSize = 9
    rpt.Printer.PaperSize = 256
    rpt.Printer.PaperHeight = 1440 * 11
    rpt.Printer.PaperWidth = 1440 * 8.5
    
    rpt.Printer.Orientation = 1
    
    If gPreviewCetak = True Then
        rpt.PrintReport False
        Unload rpt
    Else
        rpt.Show vbModal
    End If
End Sub

Private Sub BtnTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterForm(Me, True)
End Sub

Private Sub SimpanAAS()
    'On Error GoTo Err
    On Error Resume Next
    Dim SsqlNis As String
    
    Dim vSyaratByr As String
    Dim vJmlHari As Double
    Dim vTglSkr As Date
    Dim vTglTempo As Date
    Dim vTPenjualan As Double
    Dim vTPenjNonMet As Double
    Dim vTPPN As Double
    Dim vTotal As Double
    Dim vTMeterai As Double
    Dim vKeterangan As String
    
    Dim vPiutangGL As String
    Dim vPenjualanGL As String
    Dim vPPNGL As String
    Dim vMeteraiGL As String
    
    Dim vDelIndex As String
    
    vDelIndex = gKode & Left(McDelIndex(7), 8)
    vTglSkr = gTanggal
    vTglTempo = DateAdd(LCase(Left(frmBatalBilling.Text(12), 1)), Val(Right(frmBatalBilling.Text(12), Len(frmBatalBilling.Text(12)) - 1)), vTglSkr)

    vTotal = Koma2Titik(frmBatalBilling.Text(39)) + Koma2Titik(frmBatalBilling.Text(4)) + Koma2Titik(frmBatalBilling.Text(5))
    vTPenjNonMet = Round(vTotal * (100 / 110), 4)
    vTMeterai = Koma2Titik(frmBatalBilling.Text(7))
    vTPenjualan = vTPenjNonMet + vTMeterai
    vTPPN = Round(vTotal * (10 / 110), 4)

    'Proses Simpan Akses
   ' LogOnNis = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";Password='" & NisPass & "';Initial Catalog= " & NisDb & ";Data Source= " & NisKomp & ""
   '     oConnNis.ConnectionString = LogOnNis
   '     oConnNis.Open
   '     oConnNis.BeginTrans
   '     oConnNis.Errors.Clear
        
   
    'LOGONSTR2 = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=auser;Data Source=" & LetakAcces & "ardata" & Trim(gGenerasi) & ".MDB" & ";Persist Security Info=False;Jet OLEDB:System database=" & SysAcces
    ''LOGONSTR2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";Password='" & NisPass & "';Initial Catalog= " & NisDb & ";Data Source= " & NisKomp & ""
    ''oConnAcces.ConnectionString = LOGONSTR2
    ''oConnAcces.CommandTimeout = 120
    ''oConnAcces.ConnectionTimeout = 120
    ''oConnAcces.Open
    
    If Get_ConnectionAcces(NisDb, NisUSER, NisPass, NisKomp) Then
    oConnAcces.BeginTrans
    oConnAcces.Errors.Clear
    
    If vTPenjualan <> 0 Then
        'Penjualan
        sSql1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
              & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
              & " JumlahTrn, Kodeuser, DelIndex, Kodesumber) values " _
              & " ('" & Trim(frmBatalBilling.Text(57)) & "', '" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & "', 'Q'," _
              & " '" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vTglSkr & "'," _
              & " '" & vTglTempo & "', '" & vTglSkr & "'," _
              & " '" & vTglTempo & "', 'Penjualan Kredit'," _
              & " " & SimpanKoma2Titik(vTPenjualan * -1) & ", '" & gSA & "'," _
              & " '" & vDelIndex & "','" & gKode & "')"
        sSql1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
              & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
              & " JumlahTrn, Kodeuser, DelIndex, Kodesumber) values " _
              & " ('" & Trim(frmBatalBilling.Text(57)) & "', '" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & "', 'Q'," _
              & " '" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
              & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
              & " '" & FormatTgl(vTglTempo) & "', 'Penjualan Kredit'," _
              & " " & SimpanKoma2Titik(vTPenjualan * -1) & ", '" & gSA & "'," _
              & " '" & vDelIndex & "','" & gKode & "')"
        oConnAcces.Execute sSql1
        'oConnNis.Execute SsqlNis
        'PPN
        sSql1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
              & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
              & " JumlahTrn, Kodeuser, DelIndex, Kodesumber) values " _
              & " ('" & Trim(frmBatalBilling.Text(57)) & "', '" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & "', 'S'," _
              & " '" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vTglSkr & "'," _
              & " '" & vTglTempo & "', '" & vTglSkr & "'," _
              & " '" & vTglTempo & "', 'PPN Keluar No. Faktur Pajak : " & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "'," _
              & " " & SimpanKoma2Titik(vTPPN * -1) & ", '" & gSA & "'," _
              & " '" & vDelIndex & "','" & gKode & "')"
        sSql1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
              & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
              & " JumlahTrn, Kodeuser, DelIndex, Kodesumber) values " _
              & " ('" & Trim(frmBatalBilling.Text(57)) & "', '" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & "', 'S'," _
              & " '" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & FormatTgl(vTglSkr) & "'," _
              & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
              & " '" & FormatTgl(vTglTempo) & "', 'PPN Keluar No. Faktur Pajak : " & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "'," _
              & " " & SimpanKoma2Titik(vTPPN * -1) & ", '" & gSA & "'," _
              & " '" & vDelIndex & "','" & gKode & "')"
        oConnAcces.Execute sSql1
        'oConnNis.Execute SsqlNis
    
        'Update Armst Saldo berjalan
        sSql1 = "Update Armst set SaldoJalan=SaldoJalan+" & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(frmBatalBilling.Text(57)) & "'"
        oConnAcces.Execute sSql1
        'oConnNis.Execute sSql1
    End If
    End If
    
    If frmBatalBilling.cboStatus.Text = "INT" Then
        sSQL = "select * from T_AKUN"
        oRs1.Open sSQL, oConn, adOpenStatic
        If Not oRs1.BOF Then
            vPiutangGL = oRs1("FS_KD_RPIUTANG_INT")
            vPenjualanGL = oRs1("FS_KD_RLABOUR_INT")
            vPPNGL = oRs1("FS_KD_RPPN_INT")
            vMeteraiGL = oRs1("FS_KD_RMETERAI_INT")
        End If
        oRs1.Close
    Else
        sSQL = "select * from T_AKUN"
        oRs1.Open sSQL, oConn, adOpenStatic
        If Not oRs1.BOF Then
            vPiutangGL = oRs1("FS_KD_RPIUTANG_EXT")
            vPenjualanGL = oRs1("FS_KD_RLABOUR_EXT")
            vPPNGL = oRs1("FS_KD_RPPN_EXT")
            vMeteraiGL = oRs1("FS_KD_RMETERAI_EXT")
        End If
        oRs1.Close
    End If
    
    'Membuat transaksi di GL
   ' LOGONSTR2 = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=auser;Data Source=" & LetakAcces & "gldata" & Trim(gGenerasi) & ".MDB" & ";Persist Security Info=False;Jet OLEDB:System database=" & SysAcces
   ''LOGONSTR2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";Password='" & NisPass & "';Initial Catalog= " & NisDb & ";Data Source= " & NisKomp & ""
   ''oConnAcces2.ConnectionString = LOGONSTR2
   ''oConnAcces2.CommandTimeout = 120
   '' oConnAcces2.ConnectionTimeout = 120
   '' oConnAcces2.Open
   
   If Get_ConnectionAcces2(NisDb, NisUSER, NisPass, NisKomp) Then
    oConnAcces2.BeginTrans
    oConnAcces2.Errors.Clear
    
    vKeterangan = frmBatalBilling.Text(1) & frmBatalBilling.Text(26)
    
    If vTotal + vTMeterai <> 0 Then
        sSql1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
              & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan) values " _
              & " ('" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vPiutangGL & "', '" & vTglSkr & "', '" & gKode & "', '" & "C" & "', " _
              & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
              & " '" & vDelIndex & "','" & vKeterangan & "')"
        sSql1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
              & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan) values " _
              & " ('" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
              & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
              & " '" & vDelIndex & "','" & vKeterangan & "')"
        oConnAcces2.Execute sSql1
        'oConnNis.Execute SsqlNis
        
        sSql1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan) values " _
              & " ('" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & vTglSkr & "', '" & gKode & "', '" & "C" & "', " _
              & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
              & " '" & vDelIndex & "','" & vKeterangan & "')"
        sSql1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan) values " _
              & " ('" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
              & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
              & " '" & vDelIndex & "','" & vKeterangan & "')"
        oConnAcces2.Execute sSql1
        'oConnNis.Execute SsqlNis
        
        If vTMeterai <> 0 Then
            'sSql1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                  & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan) values " _
                  & " ('" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vMeteraiGL & "', '" & vTglSkr & "', '" & gKode & "', '" & "C" & "', " _
                  & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                  & " '" & vDelIndex & "','" & vKeterangan & "')"
            sSql1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                  & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan) values " _
                  & " ('" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                  & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                  & " '" & vDelIndex & "','" & vKeterangan & "')"
            oConnAcces2.Execute sSql1
            'oConnNis.Execute SsqlNis
        End If
    
        'sSql1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan) values " _
              & " ('" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vPPNGL & "', '" & vTglSkr & "', '" & gKode & "', '" & "C" & "', " _
              & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
              & " '" & vDelIndex & "','" & vKeterangan & "')"
        sSql1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan) values " _
              & " ('" & Trim(gKodeFaktur) & Trim(frmBatalBilling.Text(42)) & Trim(gKodeRetur) & "', '" & vPPNGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
              & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
              & " '" & vDelIndex & "','" & vKeterangan & "')"
       
        oConnAcces2.Execute sSql1
        'oConnNis.Execute SsqlNis
    
        sSql1 = "update glmst set JbulanIni = JbulanIni -" & (vTotal + vTMeterai) & ", " _
              & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
              & " where KodeGl = '" & vPiutangGL & "'"
        oConnAcces2.Execute sSql1
        'oConnNis.Execute sSql1
        
        sSql1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
              & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
              & " where KodeGl = '" & vPenjualanGL & "'"
        'oConnAcces2.Execute sSql1
        oConnNis.Execute sSql1
        
        sSql1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTMeterai) & ", " _
              & " JDebitBi = JDebitBi +" & vTMeterai & " " _
              & " where KodeGl = '" & vMeteraiGL & "'"
        oConnAcces2.Execute sSql1
        'oConnNis.Execute sSql1
        
        sSql1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPPN) & ", " _
              & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPPN) & " " _
              & " where KodeGl = '" & vPPNGL & "'"
        oConnAcces2.Execute sSql1
        'oConnNis.Execute sSql1
    End If
    End If
    'If oConnNis.Errors.Count = 0 Then
    '   oConnNis.CommitTrans
    '       Else
    '   oConnNis.RollbackTrans
    'End If
    Exit Sub
'Err:
'    MsgBox Err.Description
End Sub

Private Sub TblProses_Click()
    Dim NoRetur As Double
    
    If Trim(Text1.Text) = "" Then
        MsgBox "Masukkan Alasan Pembatalannya !", vbOKOnly, "Peringatan"
    Else
        'Cek di AR, ambil delindexnya untuk hapus trans & Adjust
        oConn.BeginTrans
        oConn.Errors.Clear
        vKodeLgn = frmBatalBilling.Text(57)
        
        sSQL = "Select * from T_PROFIL"
        oRs.Open sSQL, oConn, adOpenStatic
        If Not oRs.BOF Then
            NoRetur = Format(oRs("FN_NO_RET"), "00000000")
        End If
        oRs.Close
        
        If gConnecAAS = True Then
            Call SimpanAAS
        End If
'             & " CUSTOMER='" & frmBatalBilling.Text(1) & "', " _
             & " CUST_ADRES='" & frmBatalBilling.Text(2) & "', " _
             & " KOTA='" & frmBatalBilling.Text(3) & "', " _
             & " NPWP='" & frmBatalBilling.Text(14) & "', "

        vJam = Format(Time, "HH:NN")
        sSQL = "update T_INVOICE set " _
             & " FS_NO_RETUR='" & gNoInvoice & gKodeRetur & "'  , " _
             & " FS_LOG_DEL='" & gSA & "', " _
             & " ALASAN_INV='" & Text1.Text & "', " _
             & " FD_RETUR='" & TglYYYYMMDD(gTanggal) & " " & vJam & "', " _
             & " FS_KD_PLG='" & frmBatalBilling.Text(57) & "', " _
             & " FS_KD_WIL='" & Left(frmBatalBilling.Text(57), 3) & "' " _
             & " where NOINV='" & gNoInvoice & "'"
        oConn.Execute sSQL
        vNoRetur = Format(NoRetur, "00000000")
        sSQL = "update T_INVOICE set FS_KD_RETUR = '" & vNoRetur & "'" _
             & " where NOINV='" & gNoInvoice & "'"
        oConn.Execute sSQL
        
        If Val(NoRetur) >= 99999999 Then
            NoRetur = 0
        Else
            NoRetur = NoRetur + 1
        End If
        sSQL = "update T_PROFIL set FN_NO_RET = " & NoRetur
        oConn.Execute sSQL
    
        sSQL = "update JOBDTA set noinv='', INVDATE = '" & gTglDefault & "', RPR_COND='S' where NOINV='" & gNoInvoice & "'"
        oConn.Execute sSQL
        
        sSQL = "DELETE HISTORY where NOINV='" & gNoInvoice & "' and NO_WO='" & Trim(frmBatalBilling.Text(26).Text) & "'"
        oConn.Execute sSQL
    
        sSQL = "DELETE T_PART where NOINV='" & gNoInvoice & "' and NO_WO='" & Trim(frmBatalBilling.Text(26).Text) & "'"
        oConn.Execute sSQL
        
        
'        If gConnecAAS = True Then
            If oConnAcces.Errors.Count = 0 And oConnAcces2.Errors.Count = 0 _
                And oConn.Errors.Count = 0 Then
'        Else
'            If oConnAcces.Errors.Count = 0 Then
'        End If
        
            oConn.CommitTrans
            If gConnecAAS = True Then
                oConnAcces.CommitTrans
                oConnAcces2.CommitTrans
            End If
            If gUkuranRetur Then
                Call Cetak
            Else
                Call CetakFormKecil
            End If
        Else
            oConn.RollbackTrans
            If gUkuranRetur Then
                oConnAcces.RollbackTrans
                oConnAcces2.RollbackTrans
            End If
            MsgBox "Ada Error, Pembatalan Invoice dibatalkan, Batalkan invoice kembali / Hubungi EDP !"
            MsgBox Err.Description
        End If
        If gConnecAAS = True Then
            oConnAcces.Close
            oConnAcces2.Close
        End If
        Unload Me
        
        Call HapusFormRetur
    End If
End Sub

