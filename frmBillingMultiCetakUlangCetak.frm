VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBillingMultiCetakUlangCetak 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CETAK MULTI INVOICE"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   13695
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   3090
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12240
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7395
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13044
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "    Multiple Cetak Invoice"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   4194304
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   6615
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   11655
   End
End
Attribute VB_Name = "frmBillingMultiCetakUlangCetak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmBillingMultiCetakUlangCetak
' DateTime  : 04/06/2024 09:02
' Author    : HO-VM-RUDI
' Purpose   :
'---------------------------------------------------------------------------------------

 
Dim rsGrid As New ADODB.Recordset
Dim rsCetakNewBilling As New ADODB.Recordset

Dim rsTemp As ADODB.Recordset
Dim newRecord As Boolean
Dim pesan As String
Dim boleh As Boolean
Dim SSQL2 As String
Dim statusSBE As String
Dim cols As TrueOleDBGrid70.Columns
 
Dim nilaiPPnInv     As Double
Dim nilaiDppInv     As Double
Dim nilaiTotalInv   As Double
Public noInv As String

Private Sub createRS()
   On Error GoTo createRS_Error

    On Error GoTo vb_error
    
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = Nothing
    With rsGrid
        .Fields.Append "noinv", adVarChar, 50
        .Fields.Append "tglImport", adDate
        .Open
    End With
    Set GridFlate.DataSource = rsGrid
    Exit Sub
    
vb_error:
   MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateRs"), vbCritical

   On Error GoTo 0
   Exit Sub

createRS_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRS of Form frmBillingMultiCetakUlangCetak")
End Sub
 
  
 

Private Sub CmdKeluar_Click()
    Unload Me
End Sub
 

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Call cetakInv1New(noInv, 1)
    
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - 700
    CRViewer1.Width = Me.Width - 500 - asxPowerBanner1.Width
    CRViewer1.Height = asxPowerBanner1.Height
End Sub
 

Private Sub cetakInv1New(noInv As String, No As Integer)
    Dim rsRepo As ADODB.Recordset
    Dim rsAddJob As ADODB.Recordset
    Dim rsProfil As ADODB.Recordset
    Dim rsRinciSBE As ADODB.Recordset
    Dim rpt As New rtpNewInvoice
    Dim hlp1 As String
    Dim hlp2 As String
    Dim hlp3 As String
    Dim hlp4 As String
    Dim hlp6 As String
    Dim gcat As String
    Dim infoBill As String
    Dim no_wo As String
    Dim nilaiMAT As Double
    Dim tglInv  As String
    
    
    Dim onlySublet      As Boolean
    Dim isSBE           As Boolean
    Dim isFixPrice      As Boolean
    Dim discFixPrice    As Double
    Dim total           As Double
    
    If Log.notifInvNCM = 1 Then
        rpt.boxNCM.Suppress = True
        rpt.txtNCM1.Suppress = True
        rpt.txtNCM2.Suppress = True
        
        If isJualNRM = True Then
            Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NCM'")
            If Not rsTemp.EOF Then
            
                rpt.boxNCM.Suppress = False
                rpt.txtNCM1.Suppress = False
                rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                
                rpt.txtNCM2.Suppress = False
                rpt.txtNCM2.SetText Trim(rsTemp!Text2)
             
            End If
        Else
            Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NonNCM'")
            If Not rsTemp.EOF Then
            
                rpt.boxNCM.Suppress = False
                rpt.txtNCM1.Suppress = False
                rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                
                rpt.txtNCM2.Suppress = False
                rpt.txtNCM2.SetText Trim(rsTemp!Text2)
            End If
        End If
    End If
    nilaiMAT = 0
    createTableNewBilling
    sSQL = "select a.NOINV, a.NO_WO,a.FS_KD_PLG,a.CUSTOMER,a.CUST_ADRES,"
    sSQL = sSQL & " a.NO_POLISI,a.NO_CHASSIS,a.FMDL,a.NO_ENGINE,a.NO_WO,a.FRONT,A.STATUS,A.MECH,A.INVDATE,A.INVTIME,"
    sSQL = sSQL & " A.REP1,A.REP2,A.REP3,A.REP4,A.REP5,A.REP6,A.REP7,A.REP8,"
    sSQL = sSQL & " A.RTE1,A.RTE2,A.RTE3,A.RTE4,A.RTE5,A.RTE6,A.RTE7,A.RTE8,"
    sSQL = sSQL & " A.OFF1,A.OFF2,A.OFF3,A.OFF4,A.OFF5,A.OFF6,A.OFF7,A.OFF8,"
    sSQL = sSQL & " A.DISC1,A.DISC2,A.DISC3,A.DISC4,A.DISC5,A.DISC6,A.DISC7,A.DISC8,"
    sSQL = sSQL & " PART_G,PART_NG,MTRL,OIL_G,OIL_NG,MET,SBLT,LABOR,BODY,TOTAL-(FN_PAJAK+met) AS DPP,FN_PAJAK,TOTAL,FS_RC,DATE,PPn"
    sSQL = sSQL & " from T_INVOICE a with (nolock) "
    sSQL = sSQL & " where NOINV = '" & noInv & "'"
    sSQL = sSQL & " order by INVDATE desc"
    Set rsRepo = objO.FetchData(sSQL)
    
    sSQL = "SELECT A.* FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_INV = B.NOINV AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  WHERE   A.FS_LOG_DEL = '' AND B.NOINV LIKE '" & Trim(noInv) & "%'"
    Set rsAddJob = objO.FetchData(sSQL)
    
    sSQL = "select FS_NM_PERUSH,FS_ALM_PERUSH,'No Telpon : ' + FS_TELP + ', Fax : ' + FS_FAX AS FS_TELP,'Email : ' + FS_EMAIL as FS_EMAIL,'NPWP : ' + FS_NPWP + ' Tanggal : ' as FS_NPWP,'PKP : ' + FS_PKP as FS_PKP,bankname1,bankname2,bankname3,bankno1,bankno2,bankno3 from t_profil"
    Set rsProfil = objO.FetchData(sSQL)
    If rsProfil.EOF Then
        MsgBox "Gagal memuat data profil dealer", vbInformation
        Exit Sub
    End If
    If Not rsRepo.EOF Then

        nilaiPPnInv = NulSID(rsRepo("PPn"))
        nilaiDppInv = 100
        nilaiTotalInv = nilaiDppInv + nilaiPPnInv
        nilaiMAT = rsRepo!met
        rsRepo.MoveFirst
        isSBE = False
        no_wo = NulSID(rsRepo!no_wo)
        tglInv = TglDDMMMMYYYY(rsRepo!InvDate, 1) & " " & rsRepo!INVTIME
        If NulSID(rsRepo!FS_RC) <> "" Then
            isSBE = True
        End If
        onlySublet = True
        Do While Not rsRepo.EOF
            For i = 1 To 8
                hlp1 = "REP" & i
                hlp2 = "RTE" & i
                hlp3 = "OFF" & i
                hlp4 = "DISC" & i
                
                '' -------------------------------------------------
                '' 20200316 || cetakan invoice untuk fix price ditampilkan disc
                isFixPrice = False
                discFixPrice = 0
                If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(rsRepo!Date, "DD/MM/YYYY")) >= 0 Then
                    If objO.FetchData_String("select no_wo from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsRepo(hlp1)) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'") <> "" Then
                        isFixPrice = True
                        discFixPrice = objO.FetchData_Double("select disc from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsRepo(hlp1)) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'")
                    End If
                End If
                
                If Trim(rsRepo(hlp1)) <> "" Then
                    With rsCetakNewBilling
                        .AddNew
                        !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                        !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                        !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                        !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                        !alamatperusahaanpajak = NulSID("")
                        !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                        !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                        !noInv = NulSID(rsRepo!noInv)
                        !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                        !customer = NulSID(rsRepo!customer)
                        !customeralamat = NulSID(rsRepo!cust_adres)
                        !NoPolisi = NulSID(rsRepo!NO_POLISI)
                        !model = NulSID(rsRepo!Fmdl)
                        !noRangka = NulSID(rsRepo!NO_CHASSIS)
                        !noMesin = NulSID(rsRepo!No_Engine)
                        !noOrder = NulSID(rsRepo!no_wo)
                        !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                        !tipecustomer = NulSID(rsRepo!Status)
                        !mekanik = NulSID(rsRepo!Mech)
                        !billing = NulSID(rsRepo!FRONT) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                        !job = Left(IIf(InStr(NulSID(rsRepo(hlp1)), ".000") > 0, " ", "") & NulSID(rsRepo(hlp1)), 50)
                        !jobdetail = NulSID("")
                        !jobRate = k2t(NulSID(rsRepo(hlp2)))
                        !jobharga = NulSID(rsRepo(hlp3))
                        '!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
                        If isFixPrice = True Then
                            !jobdisc = NulSID(discFixPrice & " %")
                        Else
                            !jobdisc = NulSID(rsRepo(hlp4) & " %")
                        End If
                        !jobsubtotal = k2t(NulSID((rsRepo(hlp3) * (100 - NulSID(rsRepo(hlp4)))) / 100)) ''NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
                        !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                        !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                        !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                        !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                        !PPNtotal = NulA(rsRepo!fn_pajak)
                        !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                        .Update
                    End With
                    onlySublet = False
                    If isSBE Then
                        Dim nilaisbe As Integer
                        nilaisbe = CInt(Right(NulSID(rsRepo!FS_RC), 3))
                        If InStr(NulSID(rsRepo(hlp1)), nilaisbe & ".000") > 0 Then
                            Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe, String(3, "0")) & "' order by Status")
                            If Not rsRinciSBE.EOF Then
                                rsRinciSBE.MoveFirst
                                Do While Not rsRinciSBE.EOF
                                    With rsCetakNewBilling
                                        .AddNew
                                        !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                                        !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                                        !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                                        !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                                        !alamatperusahaanpajak = NulSID("")
                                        !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                                        !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                                        !noInv = NulSID(rsRepo!noInv)
                                        !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                                        !customer = NulSID(rsRepo!customer)
                                        !customeralamat = NulSID(rsRepo!cust_adres)
                                        !NoPolisi = NulSID(rsRepo!NO_POLISI)
                                        !model = NulSID(rsRepo!Fmdl)
                                        !noRangka = NulSID(rsRepo!NO_CHASSIS)
                                        !noMesin = NulSID(rsRepo!No_Engine)
                                        !noOrder = NulSID(rsRepo!no_wo)
                                        !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                                        !tipecustomer = NulSID(rsRepo!Status)
                                        !mekanik = NulSID(rsRepo!Mech)
                                        !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                                        !job = Left(" " & NulSID(rsRepo(hlp1)), 50)
                                        !jobdetail = NulSID(rsRinciSBE!Detail)
                                        !jobRate = k2t(NulSID(rsRepo(hlp2)))
                                        !jobharga = NulSID(rsRepo(hlp3))
                                        ''!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
                                        If isFixPrice = True Then
                                            !jobdisc = NulSID(discFixPrice & " %")
                                        Else
                                            !jobdisc = NulSID(rsRepo(hlp4) & " %")
                                        End If
                                        !jobsubtotal = k2t(NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100)))
                                        !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                                        !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                                        !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                                        !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                                        !PPNtotal = NulA(rsRepo!fn_pajak)
                                        !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                                        .Update
                                    End With
                                    rsRinciSBE.MoveNext
                                Loop
                            End If
                        End If
                    End If
                End If
            Next
            
            If onlySublet Then
                With rsCetakNewBilling
                    .AddNew
                    !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                    !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                    !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                    !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                    !alamatperusahaanpajak = NulSID("")
                    !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                    !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                    !noInv = NulSID(rsRepo!noInv)
                    !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                    !customer = NulSID(rsRepo!customer)
                    !customeralamat = NulSID(rsRepo!cust_adres)
                    !NoPolisi = NulSID(rsRepo!NO_POLISI)
                    !model = NulSID(rsRepo!Fmdl)
                    !noRangka = NulSID(rsRepo!NO_CHASSIS)
                    !noMesin = NulSID(rsRepo!No_Engine)
                    !noOrder = NulSID(rsRepo!no_wo)
                    !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                    !tipecustomer = NulSID(rsRepo!Status)
                    !mekanik = NulSID(rsRepo!Mech)
                    !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                    !job = NulSID("Order Sublet")
                    !jobdetail = NulSID("")
                    !jobRate = 0 'NulSID(rsRepo(hlp2))
                    !jobharga = 0 'NulSID(rsRepo(hlp3))
                    !jobdisc = 0 'NulSID(rsRepo(hlp4) & " %")
                    !jobsubtotal = 0 ' NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
                    !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                    !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                    !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                    !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                    !PPNtotal = NulA(rsRepo!fn_pajak)
                    !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                    .Update
                End With
            End If
            rsRepo.MoveNext
        Loop
        
        If Not rsAddJob.EOF Then
            rsAddJob.MoveFirst
            rsRepo.MoveFirst
            Do While Not rsAddJob.EOF
                '' -------------------------------------------------
                '' 20200316 || cetakan invoice untuk fix price ditampilkan disc
                isFixPrice = False
                discFixPrice = 0
                If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(rsRepo!Date, "DD/MM/YYYY")) >= 0 Then
                    If objO.FetchData_String("select no_wo from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsAddJob!rep) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'") <> "" Then
                        isFixPrice = True
                        discFixPrice = objO.FetchData_Double("select disc from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsAddJob!rep) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'")
                    End If
                End If
                With rsCetakNewBilling
                    .AddNew
                    !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                    !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                    !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                    !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                    !alamatperusahaanpajak = NulSID("")
                    !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                    !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                    !noInv = NulSID(rsRepo!noInv)
                    !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                    !customer = NulSID(rsRepo!customer)
                    !customeralamat = NulSID(rsRepo!cust_adres)
                    !NoPolisi = NulSID(rsRepo!NO_POLISI)
                    !model = NulSID(rsRepo!Fmdl)
                    !noRangka = NulSID(rsRepo!NO_CHASSIS)
                    !noMesin = NulSID(rsRepo!No_Engine)
                    !noOrder = NulSID(rsRepo!no_wo)
                    !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                    !tipecustomer = NulSID(rsRepo!Status)
                    !mekanik = NulSID(rsRepo!Mech)
                    !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                    !job = Left("Add. " & NulSID(rsAddJob!rep), 50)
                    !jobdetail = NulSID("")
                    !jobRate = k2t(NulSID(rsAddJob!rte))
                    !jobharga = NulSID(rsAddJob!offn)
                    ''!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "") 'NulSID(rsAddJob!disc) & " %"
                    If isFixPrice = True Then
                        !jobdisc = NulSID(discFixPrice & " %")
                    Else
                        !jobdisc = NulSID(rsAddJob!disc) & " %"
                    End If
                    !jobsubtotal = Round(NulSID(rsAddJob!offn * ((100 - NulSID(rsAddJob!disc)) / 100)))
                    !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                    !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                    !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                    !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                    !PPNtotal = NulA(rsRepo!fn_pajak)
                    !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                    .Update
                End With
                If isSBE Then
                    Dim nilaisbe2 As Integer
                    nilaisbe2 = CInt(Right(NulSID(rsRepo!FS_RC), 3))
                    If InStr(NulSID(rsRepo(hlp1)), nilaisbe2 & ".000") > 0 Then
                        Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe2, String(3, "0")) & "' order by Status")
                        If Not rsRinciSBE.EOF Then
                            rsRinciSBE.MoveFirst
                            Do While Not rsRinciSBE.EOF
                                With rsCetakNewBilling
                                    .AddNew
                                    !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                                    !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                                    !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                                    !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                                    !alamatperusahaanpajak = NulSID("")
                                    !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                                    !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                                    !noInv = NulSID(rsRepo!noInv)
                                    !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                                    !customer = NulSID(rsRepo!customer)
                                    !customeralamat = NulSID(rsRepo!cust_adres)
                                    !NoPolisi = NulSID(rsRepo!NO_POLISI)
                                    !model = NulSID(rsRepo!Fmdl)
                                    !noRangka = NulSID(rsRepo!NO_CHASSIS)
                                    !noMesin = NulSID(rsRepo!No_Engine)
                                    !noOrder = NulSID(rsRepo!no_wo)
                                    !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                                    !tipecustomer = NulSID(rsRepo!Status)
                                    !mekanik = NulSID(rsRepo!Mech)
                                    !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                                    !job = Left(" " & "Add. " & NulSID(rsAddJob!rep), 50)
                                    !jobdetail = NulSID(rsRinciSBE!Detail)
                                    !jobRate = k2t(NulSID(rsAddJob!rte))
                                    !jobharga = NulSID(rsAddJob!offn)
                                    ''!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "") 'NulSID(rsAddJob!disc) & " %"
                                    If isFixPrice = True Then
                                        !jobdisc = NulSID(discFixPrice & " %")
                                    Else
                                        !jobdisc = NulSID(rsAddJob!disc) & " %"
                                    End If
                                    !jobsubtotal = k2t(NulSID(rsAddJob!offn * ((100 - NulSID(rsAddJob!disc)) / 100)))
                                    !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                                    !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                                    !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                                    !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                                    !PPNtotal = NulA(rsRepo!fn_pajak)
                                    !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                                    .Update
                                End With
                                rsRinciSBE.MoveNext
                            Loop
                        End If
                    End If
                End If
                rsAddJob.MoveNext
            Loop
        End If
        gcat = ""
        If nRC Then
            Dim rsHlp As ADODB.Recordset
            Dim nBulan As String
            Dim nSbe As String
            nSbe = ""
            Set rsHlp = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS JT_TGL, SBENEXT FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & Trim(rsRepo!NO_CHASSIS) & "' AND TGLSVCNOW = '" & Format(rsRepo!Date, "MM/DD/YYYY") & "'")
            If Not rsHlp.EOF Then
                nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
                nSbe = Trim(rsHlp("SBENEXT"))
            End If
            rsHlp.Close
            If Trim(nSbe) <> "" Then
                'gcat = IIf(nHeavy, "Kembali untuk service berkala setelah kelipatan 5.000 KM atau bulan " & nBulan, "Kembali untuk service berkala setelah kelipatan 10.000 KM atau bulan " & nBulan)
                ''gcat = "Kembali untuk service berkala " & Format(CDbl(nSbe), "#,#0") & " KM atau bulan " & nBulan
                gcat = "Kembali untuk service berkala bulan " & nBulan & " atau " & Format(CDbl(nSbe), "#,#0") & " KM "
                gcat = gcat + IIf(gcat <> "", " (mana yang tercapai lebih dahulu).", "")
            Else
               '' Call objO.InsertData("INSERT INTO ESTIMASISALAH (norangka, tglservice, kmsvc, sbesvc, kelipatan, noinv, nowo,customer, alamat,nopolisi, telpon, gsm,cdma) VALUES " & _
                               "('" & Trim(Text(17)) & "','" & Format(Text(43), "MM/DD/YYYY") & "'," & CDbl(Text(37)) & ",'" & Trim(Text(33)) & "','" & IIf(nHeavy, 5000, 10000) & "', " & _
                               "'" & Trim(Text(20)) & "','" & Trim(Trim(Text(0))) & "','" & Trim(Text(4)) & "','" & Trim(txt5) & "','" & Trim(Text(16)) & "','" & Trim(Text(7)) & "','" & Trim(Text(8)) & "','" & Trim(Text(9)) & "')")
            End If
        End If
        
        rpt.txtTerbilang.SetText "Terbilang (Rupiah) " & Trim(Angka2Huruf(rsCetakNewBilling!grandTotal, 1))
        ''20240829 revisi menutup kembali service
        rpt.txtNextService.SetText "" 'gcat
        
        '' inspection ====================
        If Log.CR7 = 1 Then
            Dim rsInspection As ADODB.Recordset
            Dim total1 As Double
            Dim total2 As Double
            Dim total3 As Double
            Dim total4 As Double
            
            sSQL = "select sp1,sp2,sp3,sp4,estn1,estn2,estn3,estn4,estj1,estj2,estj3,estj4,estp1,estp2,estp3,estp4,ests1,ests2,ests3,ests4 from INSPECTION with (nolock) where NoWo = '" & no_wo & "' order by tglentry desc"
            Set rsInspection = objO.FetchData(sSQL)
            If Not rsInspection.EOF Then
                rsInspection.MoveFirst
                rpt.txtSaran1.SetText NulSID(rsInspection!sp1)
                rpt.txtSaran2.SetText NulSID(rsInspection!sp2)
                rpt.txtSaran3.SetText NulSID(rsInspection!sp3)
                rpt.txtSaran4.SetText NulSID(rsInspection!sp4)
                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtWaktuSVC1.SetText Format(NulSID(rsInspection!estn1), "dd-mm-yyyy")
                Else
                    rpt.txtWaktuSVC1.SetText ""
                End If
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtWaktuSVC2.SetText Format(NulSID(rsInspection!estn2), "dd-mm-yyyy")
                Else
                    rpt.txtWaktuSVC2.SetText ""
                End If
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtWaktuSVC3.SetText Format(NulSID(rsInspection!estn3), "dd-mm-yyyy")
                Else
                    rpt.txtWaktuSVC3.SetText ""
                End If
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtWaktuSVC4.SetText Format(NulSID(rsInspection!estn4), "dd-mm-yyyy")
                Else
                    rpt.txtWaktuSVC4.SetText ""
                End If
                                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtEstJasa1.SetText Format(NulSID(rsInspection!estj1), "#,0")
                Else
                    rpt.txtEstJasa1.SetText ""
                End If
                total1 = total1 + NulSID(rsInspection!estj1)
                
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtEstJasa2.SetText Format(NulSID(rsInspection!estj2), "#,0")
                Else
                    rpt.txtEstJasa2.SetText ""
                End If
                total2 = total2 + NulSID(rsInspection!estj2)
                
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtEstJasa3.SetText Format(NulSID(rsInspection!estj3), "#,0")
                Else
                    rpt.txtEstJasa3.SetText ""
                End If
                total3 = total3 + NulSID(rsInspection!estj3)
                
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtEstJasa4.SetText Format(NulSID(rsInspection!estj4), "#,0")
                Else
                    rpt.txtEstJasa4.SetText ""
                End If
                total4 = total4 + NulSID(rsInspection!estj4)
                                                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtEstPart1.SetText Format(NulSID(rsInspection!estp1), "#,0")
                Else
                    rpt.txtEstPart1.SetText ""
                End If
                total1 = total1 + NulSID(rsInspection!estp1)
                                                
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtEstPart2.SetText Format(NulSID(rsInspection!estp2), "#,0")
                Else
                    rpt.txtEstPart2.SetText ""
                End If
                total2 = total2 + NulSID(rsInspection!estp2)
                                                
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtEstPart3.SetText Format(NulSID(rsInspection!estp3), "#,0")
                Else
                    rpt.txtEstPart3.SetText ""
                End If
                total3 = total3 + NulSID(rsInspection!estp3)
                                                
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtEstPart4.SetText Format(NulSID(rsInspection!estp4), "#,0")
                Else
                    rpt.txtEstPart4.SetText ""
                End If
                total4 = total4 + NulSID(rsInspection!estp4)
                                
                                                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtEstSublet1.SetText Format(NulSID(rsInspection!ests1), "#,0")
                Else
                    rpt.txtEstSublet1.SetText ""
                End If
                total1 = total1 + NulSID(rsInspection!ests1)
                                                
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtEstSublet2.SetText Format(NulSID(rsInspection!ests2), "#,0")
                Else
                    rpt.txtEstSublet2.SetText ""
                End If
                total2 = total2 + NulSID(rsInspection!ests2)
                                                
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtEstSublet3.SetText Format(NulSID(rsInspection!ests3), "#,0")
                Else
                    rpt.txtEstSublet3.SetText ""
                End If
                total3 = total3 + NulSID(rsInspection!ests3)
                                                
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtEstSublet4.SetText Format(NulSID(rsInspection!ests4), "#,0")
                Else
                    rpt.txtEstSublet4.SetText ""
                End If
                total4 = total4 + NulSID(rsInspection!ests4)
                                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtEstTotal1.SetText Format(NulSID(total1), "#,0")
                Else
                    rpt.txtEstTotal1.SetText ""
                End If
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtEstTotal2.SetText Format(NulSID(total2), "#,0")
                Else
                    rpt.txtEstTotal2.SetText ""
                End If
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtEstTotal3.SetText Format(NulSID(total3), "#,0")
                Else
                    rpt.txtEstTotal3.SetText ""
                End If
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtEstTotal4.SetText Format(NulSID(total4), "#,0")
                Else
                    rpt.txtEstTotal4.SetText ""
                End If
                
            End If
        End If
        '' akhir inspection ==============
                         
        rpt.txtBankNama1.SetText NulSID(rsProfil!bankname1)
        rpt.txtBankNama2.SetText NulSID(rsProfil!bankname2)
        rpt.txtBankNama3.SetText NulSID(rsProfil!bankname3)
        
        'rpt.txtBankNo1.SetText NulSID(rsProfil!bankno1)
        'rpt.txtBankNo2.SetText NulSID(rsProfil!bankno2)
        'rpt.txtBankNo3.SetText NulSID(rsProfil!bankno3)
        
        rpt.txtInfoMaterai.SetText ""
        rpt.txtLunasMaterai.SetText ""
        
        If Log.fakturMaterai = 1 Then
            Set rsTemp = objO.FetchData("SELECT * FROM t_invoiceMaterai WITH (NOLOCK) WHERE NOINV = '" & noInv & "'")
            If Not rsTemp.EOF Then
                rpt.txtInfoMaterai.SetText "Ijin Pembubuhan Materai No. " & NulSID(rsTemp!NOIJIN)
                rpt.txtLunasMaterai.SetText "BEA MATERAI LUNAS " & NulSID(rsTemp!nilai)
            End If
        End If
        rsCetakNewBilling.MoveFirst
        
        rpt.Text57.SetText NulSID(objO.FetchData_String("SELECT FS_KA_BENGKEL FROM T_PROFIL"))
        rpt.Text8.SetText NulSID(objO.FetchData_String("SELECT FS_KOTA FROM T_PROFIL"))
        
        rpt.txtMaterai.SetText IIf(nilaiMAT = 0, "0,00", Format(NulSID(nilaiMAT), "#,###.00")) ''Format(NulSID(nilaiMAT), "#,###.00")
        rpt.txtTglBilling.SetText tglInv
        rpt.Text41.SetText "PPn " & nilaiPPnInv & " %"
        rpt.Text55.SetText "PPn " & nilaiPPnInv & " %"
        rpt.database.SetDataSource rsCetakNewBilling
        If No = 1 Then
            With CRViewer1
                .ReportSource = rpt
                .ViewReport
                .Zoom 100
                .Top = 0
                .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
                .Width = Frame1.Width + 100
                .Height = asxPowerBanner1.Height
                
                cmdTutup.Top = .Top + 430
                cmdTutup.Left = .Left
                cmdTutup.Visible = True
                cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 5
                cmdSettingPrinter.Visible = True
                cmdSettingPrinter.Top = cmdTutup.Top
                .Visible = True
            End With
            DoEvents
            cmdTutup.Visible = True
        
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
        
    Else
        MsgBox "Data invoice tidak ditemukan", vbInformation
    End If
End Sub



Private Sub createTableNewBilling()
1        On Error GoTo createTableNewBilling_Error
2         If rsCetakNewBilling.State Then rsCetakNewBilling.Close
3         Set rsCetakNewBilling = Nothing
          
4         With rsCetakNewBilling
5             .Fields.Append "namaperusahaan", adVarChar, 200
6             .Fields.Append "alamatperusahaan", adVarChar, 200
7             .Fields.Append "teleponperusahaan", adVarChar, 200
8             .Fields.Append "emailperusahaan", adVarChar, 200
9             .Fields.Append "alamatperusahaanpajak", adVarChar, 200
10            .Fields.Append "npwpperusahaan", adVarChar, 200
11            .Fields.Append "pkpperusahaan", adVarChar, 200
12            .Fields.Append "noinv", adVarChar, 20
13            .Fields.Append "kodelgn", adVarChar, 30
14            .Fields.Append "customer", adVarChar, 100
15            .Fields.Append "customeralamat", adVarChar, 200
16            .Fields.Append "nopolisi", adVarChar, 12
17            .Fields.Append "model", adVarChar, 50
18            .Fields.Append "norangka", adVarChar, 25
19            .Fields.Append "nomesin", adVarChar, 15
20            .Fields.Append "noorder", adVarChar, 40
21            .Fields.Append "namasa", adVarChar, 40
22            .Fields.Append "tipecustomer", adVarChar, 40
23            .Fields.Append "mekanik", adVarChar, 40
24            .Fields.Append "printed", adVarChar, 40
              ''.Fields.Append "billing", adDate
              ''20191108 REQ MB THANIA BILLING DIBUAT INISIAL FRONT
25            .Fields.Append "billing", adVarChar, 10
26            .Fields.Append "job", adVarChar, 50
27            .Fields.Append "jobdetail", adVarChar, 100
28            .Fields.Append "jobrate", adDouble
29            .Fields.Append "jobharga", adDouble
30            .Fields.Append "jobdisc", adVarChar, 10
31            .Fields.Append "jobsubtotal", adDouble
32            .Fields.Append "dppJasa", adDouble
33            .Fields.Append "dppPart", adDouble
34            .Fields.Append "dppSublet", adDouble
35            .Fields.Append "dppTotal", adDouble
36            .Fields.Append "ppnTotal", adDouble
37            .Fields.Append "grandTotal", adDouble
38            .Open
39        End With

40       On Error GoTo 0
41       Exit Sub
createTableNewBilling_Error:
42        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createTableNewBilling of Form frmBilling"
End Sub
