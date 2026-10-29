VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "actbar.ocx"
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Layar 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PT New Ratna Motor, NASMOCO INTEGRATED SYSTEM - SERVICE"
   ClientHeight    =   9420
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "Layar.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar1 
      Align           =   1  'Align Top
      Height          =   8865
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11400
      _LayoutVersion  =   1
      _ExtentX        =   20108
      _ExtentY        =   15637
      _DataPath       =   ""
      Bands           =   "Layar.frx":628A
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6120
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1279
         ImageHeight     =   699
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Layar.frx":2BD49
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ActiveBarLibraryCtl.ActiveBar ActiveBar12 
         Left            =   600
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Bands           =   "Layar.frx":2BB29B
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1320
      Top             =   1800
   End
   Begin MSComctlLib.StatusBar pic4text 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8865
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19579
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   9120
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   13
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "0 IP : 10.10.10.100  HO:0"
            TextSave        =   "0 IP : 10.10.10.100  HO:0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2611
            MinWidth        =   2611
            Text            =   "01 Desember 2000"
            TextSave        =   "01 Desember 2000"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "24:00"
            TextSave        =   "24:00"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   88
            MinWidth        =   88
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "UE : 000"
            TextSave        =   "UE : 000"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   397
            MinWidth        =   176
            Text            =   "..."
            TextSave        =   "..."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   185
            MinWidth        =   176
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   159
            MinWidth        =   88
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel10 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   185
            MinWidth        =   176
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel11 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel12 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel13 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Layar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim wktAkhir As String
 Dim support As New clsFormDealer
 Dim nUnit As Integer
 Dim gTglTampung As Date
 Dim rsDt As ADODB.Recordset
 Dim rsHlp As ADODB.Recordset
 Dim rsHlp2 As ADODB.Recordset
 Dim rsHlp4 As ADODB.Recordset
Private objAR As New ClsDataPelanggan

'ByVal Tool As ActiveBarLibraryCtl.Tool
Private Sub ActiveBar1_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    gTglTampung = isnow
     
     If Tool.Visible = True Then
        'MsgBox Tool.Name
        'Exit Sub
        
        Select Case Tool.Name
        
        Case "mnu3"
            frmLogin.show
        Case "mnu2"
            viewLayar.Visible = True
            viewLayar.Top = ActiveBar1.Top + ActiveBar1.Height + 20
            viewLayar.Left = 1440
            Call panggilLayar
        ' booking
        Case "mnu1_1_1"
            If gFiskal = Format(Month(gTglTampung), "00") Then
'                If Log.DmsNonDss = "1" Then
'                    '-----------------------------------
'                    '.:: SET AS CHILD ::.
'                    '-----------------------------------
'                    TOChild Layar.hWnd, frmBookingDMSNonDSS.hWnd
'                    '-----------------------------------
'                    '-----------------------------------
'                    isBookByMrs = False
'                    frmBookingDMSNonDSS.Height = frmBookingDMSNonDSS.asxPowerBanner1.Height
'                    frmBookingDMSNonDSS.Show
'                Else
                    '-----------------------------------
                     '.:: SET AS CHILD ::.
                    '-----------------------------------
                    TOChild Layar.hWnd, frmBooking.hWnd
                    '-----------------------------------
                    '-----------------------------------
                    isBookByMrs = False
                    frmBooking.Height = frmBooking.asxPowerBanner1.Height
                    frmBooking.show
'                End If
            Else
                MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
            End If
        Case "mnu1_1_2"
            TOChild Layar.hWnd, frmBtlBooking.hWnd
            frmBtlBooking.show
            frmBtlBooking.WindowState = vbMaximized
            
        Case "mnu1_1_3"
            TOChild Layar.hWnd, frmEstimasi.hWnd
            frmEstimasi.show
            frmEstimasi.WindowState = vbMaximized
        Case "mnu1_1_4"
            'frmStallDisplay.DtTanggal = Format(isnow, "DD/MM/YYYY")
            TOChild Layar.hWnd, frmStallDisplay.hWnd
            frmStallDisplay.show
            frmStallDisplay.WindowState = vbMaximized
        Case "mnu1_1_5"
            TOChild Layar.hWnd, frmStallJamTerpakai.hWnd
            frmStallJamTerpakai.show
            frmStallJamTerpakai.WindowState = vbMaximized
        Case "mnu1_1_6"
            TOChild Layar.hWnd, frmBookingMenuOnline.hWnd
            frmBookingMenuOnline.show
            frmBookingMenuOnline.WindowState = vbMaximized
        Case "mnu1_1_7"
            If Log.NHS = 1 Then
                If NulSID(gTipeUser) = "KOORD NHS" Or NulSID(gTipeUser) = "ASSKABENG" Or NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                    TOChild Layar.hWnd, frmTPlottingBookingNHS.hWnd
                    frmTPlottingBookingNHS.show
                    frmTPlottingBookingNHS.WindowState = vbMaximized
                Else
                    MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
                End If
            Else
                MsgBox "Project NHS belum diaktifkan.", vbInformation
            End If
        Case "mnu1_1_8"
            TOChild Layar.hWnd, frmBookingTCO.hWnd
            frmBookingTCO.show
            frmBookingTCO.WindowState = vbMaximized
            
            
        ' order
        Case "mnu1_2_1"
            If gFiskal = Format(Month(gTglTampung), "00") Then
                'nUnit = isiCustFU(gSA)
                    
                If gNIK = "" Then
                    MsgBox "Silahkan lengkapi data NIK anda", vbInformation
                    Exit Sub
                End If
                If Log.followUp = 1 And Trim(gTipeUser) = "SA" Then
                  ''2016-07-11 || pengecekan h+1 follow up berdasarkan kode sa tersebut
                  If Trim(gTipeUser) = "SA" And Log.followUp = 1 Then
                    Dim rsFollow As ADODB.Recordset
                    Dim rsOrder As ADODB.Recordset
                    Dim selisih As Integer
                    ''jumlah order untuk follow up +1 hari
                    Set rsOrder = objO.FetchData("select count(l.no_wo) total from JOBDTL l with(nolock)  inner join jobdta a with(nolock) on a.no_wo = l.no_wo and a.no_chassis = l.no_chassis and a.fs_log_del = '' where DATEDIFF(day,tglfu,getdate()) = 1 AND a.FRONT = '" & Trim(gSA) & "'")
                    
                    ''jumlah data follow untuk tglfu +1 hari
                    Set rsFollow = objO.FetchData("select COUNT(NO_WO) total from DTLFU where NO_WO in (select l.no_wo from JOBDTL l with(nolock)  inner join jobdta a with(nolock) on a.no_wo = l.no_wo and a.no_chassis = l.no_chassis and a.fs_log_del = '' where DATEDIFF(day,tglfu,getdate()) = 1 AND a.FRONT = '" & Trim(gSA) & "')")
                    selisih = 0
                    
                    If Not rsOrder.EOF Then
                       selisih = CInt(rsOrder!total) ''CInt(rsOrder.RecordCount) - CInt(rsFollow.RecordCount)
                    End If
                    
                    If Not rsFollow.EOF Then
                       selisih = selisih - CInt(rsFollow!total)
                    End If
                    
                    If selisih > 0 Then
                       ''munculkan peringatan follow up
                       MsgBox "Menu order tidak dapat dibuka." & vbCrLf & "Masih ada " & selisih & " no wo yang belum difollow up", vbInformation
                       ''Unload Me
                       Exit Sub
                    End If
                  End If
                  frmOrder.noRangkaUnitCome = ""
                  frmOrder.noPolisiUnitCome = ""
                  frmOrder.tanggalInUnitCome = ""
                  frmOrder.show
                  frmOrder.WindowState = vbMaximized
                  'TOChild Layar.hWnd, frmOrder.hWnd
                Else
                  If gLimitFUDealer - nUnit <= 10 And gLimitFUDealer - nUnit > 0 Then
                      MsgBox "Hanya bisa menerima " & CDbl(gLimitFUDealer - nUnit) & " Order lagi. " & vbCr & "Lakukan Follow Up Terlebih dahulu !", vbInformation, "Limit Follow Up"
                  ElseIf gLimitFUDealer <= nUnit Then
                      MsgBox "Tidak bisa menerima Order lagi. " & vbCr & "Lakukan Follow Up Terlebih dahulu !", vbInformation, "Limit Follow Up"
                  End If
                  If nUnit < gLimitFUDealer Then
                      frmOrder.noRangkaUnitCome = ""
                      frmOrder.noPolisiUnitCome = ""
                      frmOrder.tanggalInUnitCome = ""
                      frmOrder.show
                      frmOrder.WindowState = vbMaximized
                      'TOChild Layar.hWnd, frmOrder.hWnd
                  ElseIf nUnit > gLimitFUDealer Then
                      MsgBox "Customer yang belum di-Follow telah mencapai batas " & gLimitFUDealer & " unit" & vbCr & "Lakukan Follow Up Terlebih dahulu !", vbInformation, "Limit Follow Up"
                  End If
                End If
            Else
                MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
            End If
        Case "mnu1_2_2"
            TOChild Layar.hWnd, frmBtlOrd.hWnd
            frmBtlOrd.show
            frmBtlOrd.WindowState = vbMaximized
        Case "mnu1_2_3"
            TOChild Layar.hWnd, frmInfoAllRepairGraph.hWnd
            frmInfoAllRepairGraph.show
            frmInfoAllRepairGraph.WindowState = vbMaximized
        Case "mnu1_2_4"
            TOChild Layar.hWnd, frm3DFU.hWnd
            frm3DFU.show
            frm3DFU.WindowState = vbMaximized
        Case "mnu1_2_5"
            frmOrderValidasi.show
            frmOrderValidasi.WindowState = vbMaximized
        Case "mnu1_2_6"
            If gFiskal = Format(Month(gTglTampung), "00") Then
                ' menu validasi Sublet
                If Log.standarisasiSublet = 1 Then
                    frmOSubletSentralValid.show
                    frmOSubletSentralValid.WindowState = vbMaximized
                Else '
                    frmOSubletValid.show
                    frmOSubletValid.WindowState = vbMaximized
                End If
                
            Else
                MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
            End If
        Case "mnu1_2_7"
            If gFiskal = Format(Month(gTglTampung), "00") Then
                ' menu retur sublet
                If Log.standarisasiSublet = 1 Then
                    frmOSubletSentralRetur.show
                    frmOSubletSentralRetur.WindowState = vbMaximized
                Else
                    frmOSubletRetur.show
                    frmOSubletRetur.WindowState = vbMaximized
                End If
                
            Else
                MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
            End If
        Case "mnu1_2_8"
            TOChild Layar.hWnd, frmValidasiTWC.hWnd
            frmValidasiTWC.show
            frmValidasiTWC.WindowState = vbMaximized
        Case "mnu1_2_9"
            TOChild Layar.hWnd, FrmOrderPenerimaanUnit.hWnd
            FrmOrderPenerimaanUnit.show
            FrmOrderPenerimaanUnit.WindowState = vbMaximized
        Case "mnu1_2_10"
            TOChild Layar.hWnd, FrmOrderListUnitCome.hWnd
            FrmOrderListUnitCome.show
            FrmOrderListUnitCome.WindowState = vbMaximized
        Case "mnu1_2_11"
            TOChild Layar.hWnd, frmOrderCuci.hWnd
            frmOrderCuci.show
            frmOrderCuci.WindowState = vbMaximized
        Case "mnu1_2_12"
            Set rsHlp = objO.FetchData("SELECT * FROM QTY_STALLWASHING WHERE TIPESTALL = 'Pengering' AND AKTIF = '1' AND ISDEL = '0'")
            If rsHlp.RecordCount > 0 Then
                TOChild Layar.hWnd, frmOrderPengeringCuci.hWnd
                frmOrderPengeringCuci.show
                frmOrderPengeringCuci.WindowState = vbMaximized
            Else
                MsgBox "Stall Pengering tidak aktif", vbInformation
                Exit Sub
            End If
        Case "mnu1_2_13"
            TOChild Layar.hWnd, frmOrderListCuci.hWnd
            frmOrderListCuci.show
            frmOrderListCuci.WindowState = vbMaximized
        Case "mnu1_2_14"
            TOChild Layar.hWnd, frmSupportTL.hWnd
            frmSupportTL.show
            frmSupportTL.WindowState = vbMaximized
        Case "mnu1_2_15"
            TOChild Layar.hWnd, frmOrderComplain.hWnd
            frmOrderComplain.show
            frmOrderComplain.WindowState = vbMaximized
        Case "mnu1_2_16"
            If Log.FIBeforeSublet = 1 Then
                TOChild Layar.hWnd, frmOrderListSublet.hWnd
                frmOrderListSublet.show
                frmOrderListSublet.WindowState = vbMaximized
            End If
        Case "mnu1_2_17"
            If Log.loyalityCustomer = 1 Then
                TOChild Layar.hWnd, frmLoyalitiVerifikasiCustomer.hWnd
                frmLoyalitiVerifikasiCustomer.show
                frmLoyalitiVerifikasiCustomer.WindowState = vbMaximized
            End If
        Case "mnu1_2_18"
            If Log.loyalityCustomer = 1 Then
                TOChild Layar.hWnd, frmLoyalityGantiUnit.hWnd
                frmLoyalityGantiUnit.show
                frmLoyalityGantiUnit.WindowState = vbMaximized
            End If
        'progress
        Case "mnu1_3_1"
            TOChild Layar.hWnd, frmProgress.hWnd
            frmProgress.show
            frmProgress.WindowState = vbMaximized
        Case "mnu1_3_2"
'            frmProgressGraph.Show
'            frmProgressGraph.WindowState = vbMaximized
        Case "mnu1_3_3"
            'frmStallProgress.WindowState = vbMaximized
            'frmStallProgress.Show
        Case "mnu1_3_4"
            TOChild Layar.hWnd, frmSProgresCustomer.hWnd
            frmSProgresCustomer.show
            frmSProgresCustomer.WindowState = vbMaximized
        Case "mnu1_3_5"
            
        'billing
        Case "mnu1_4_1"
            If gFiskal = Format(Month(gTglTampung), "00") Then
                TOChild Layar.hWnd, frmBilling.hWnd
                frmBilling.show
                frmBilling.WindowState = vbMaximized
            Else
                MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
            End If
        Case "mnu1_4_2"
            If gFiskal = Format(Month(gTglTampung), "00") Then
                TOChild Layar.hWnd, frmBtlInv.hWnd
                frmBtlInv.show
                frmBtlInv.WindowState = vbMaximized
            Else
                MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
            End If
        Case "mnu1_4_3"
            TOChild Layar.hWnd, frmCatatKredit.hWnd
            frmCatatKredit.show
            frmCatatKredit.WindowState = vbMaximized
        Case "mnu1_4_4"
            TOChild Layar.hWnd, frmBatalCatatKredit.hWnd
            frmBatalCatatKredit.show
            frmBatalCatatKredit.WindowState = vbMaximized
        Case "mnu1_4_5"
            TOChild Layar.hWnd, frmEditBilling.hWnd
            frmEditBilling.show
            frmEditBilling.WindowState = vbMaximized
        Case "mnu1_4_6"
            TOChild Layar.hWnd, frmEditRate.hWnd
            frmEditRate.show
            frmEditRate.WindowState = vbMaximized
        Case "mnu1_4_7"
            TOChild Layar.hWnd, frmBillingCtk.hWnd
            frmBillingCtk.show
            frmBillingCtk.WindowState = vbMaximized
        Case "mnu1_4_8"
            TOChild Layar.hWnd, frmBillRekap.hWnd
            frmBillRekap.show
            frmBillRekap.WindowState = vbMaximized
        Case "mnu1_4_9"
            TOChild Layar.hWnd, frmBillingUlang.hWnd
            frmBillingUlang.show
            frmBillingUlang.WindowState = vbMaximized
        Case "mnu1_4_10"
            TOChild Layar.hWnd, frmFakturPajak.hWnd
            frmFakturPajak.show
            frmFakturPajak.WindowState = vbMaximized
        Case "mnu1_4_11"
            TOChild Layar.hWnd, frmPiutangAR.hWnd
            frmPiutangAR.show
            frmPiutangAR.WindowState = vbMaximized
        Case "mnu1_4_12"
            If Log.freepart = 1 Then
                TOChild Layar.hWnd, frmInputWCARTAM.hWnd
                frmInputWCARTAM.show
                frmInputWCARTAM.WindowState = vbMaximized
            Else
                MsgBox "Project 3 invoice belum diaktifkan.", vbInformation
            End If
        Case "mnu1_4_13"
            TOChild Layar.hWnd, frmBillingMultiCetakUlang.hWnd
            frmBillingMultiCetakUlang.show
            frmBillingMultiCetakUlang.WindowState = vbMaximized
        'table master
        Case "mnu1_5_1"
            TOChild Layar.hWnd, frmtJPekerjaan.hWnd
            frmtJPekerjaan.show
            frmtJPekerjaan.WindowState = vbMaximized
        Case "mnu1_5_2"
            'frmtModelCbg.Show
            'frmtModelCbg.WindowState = vbMaximized
        Case "mnu1_5_3"
            TOChild Layar.hWnd, frmtMekanik.hWnd
            frmtMekanik.show
            frmtMekanik.WindowState = vbMaximized
        Case "mnu1_5_4"
            TOChild Layar.hWnd, frmtJobList.hWnd
            frmtJobList.show
            frmtJobList.WindowState = vbMaximized
        Case "mnu1_5_5"
            TOChild Layar.hWnd, frmtMeterai.hWnd
            frmtMeterai.show
            frmtMeterai.WindowState = vbMaximized
        Case "mnu1_5_6"
            TOChild Layar.hWnd, frmtDealer.hWnd
            frmtDealer.show
            frmtDealer.WindowState = vbMaximized
        Case "mnu1_5_7"
            '---------------------
            ' .:: SET AS CHILD ::.
            '---------------------
'            TOChild Layar.hWnd, frmtCustomer.hWnd
'            frmtCustomer.Show
'            frmtCustomer.WindowState = vbMaximized
            TOChild Layar.hWnd, frmSIDCustomer.hWnd
            frmSIDCustomer.show
            frmSIDCustomer.WindowState = vbMaximized
        Case "mnu1_5_8"
            TOChild Layar.hWnd, frmtHobi.hWnd
            frmtHobi.show
            frmtHobi.WindowState = vbMaximized
        Case "mnu1_5_9"
            '---------------------
            ' .:: SET AS CHILD ::.
            '---------------------
'            TOChild Layar.hWnd, frmtPlgn.hWnd
'            frmtPlgn.Show
'            frmtPlgn.WindowState = vbMaximized
        Case "mnu1_5_10"
            TOChild Layar.hWnd, frmtWil.hWnd
            frmtWil.show
            frmtWil.WindowState = vbMaximized
        Case "mnu1_5_11"
            TOChild Layar.hWnd, frmtAgama.hWnd
            frmtAgama.show
            frmtAgama.WindowState = vbMaximized
        Case "mnu1_5_12"
            TOChild Layar.hWnd, frmtAsuransi.hWnd
            frmtAsuransi.show
            frmtAsuransi.WindowState = vbMaximized
        Case "mnu1_5_13"
            TOChild Layar.hWnd, frmtActRate.hWnd
            frmtActRate.show
            frmtActRate.WindowState = vbMaximized
        Case "mnu1_5_14"
            TOChild Layar.hWnd, frmDealer.hWnd
            frmDealer.show
            frmDealer.WindowState = vbMaximized
        Case "mnu1_5_15"
            If Log.NHS = 1 Then
                
                If NulSID(gTipeUser) = "KOORD NHS" Or NulSID(gTipeUser) = "ASSKABENG" Or NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                    TOChild Layar.hWnd, frmTUnitNHS.hWnd
                    frmTUnitNHS.show
                    frmTUnitNHS.WindowState = vbMaximized
                Else
                    MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
                End If
            Else
                MsgBox "Project NHS belum diaktifkan", vbInformation
            End If
        Case "mnu1_5_16"
            If Log.NHS = 1 Then
                
                If NulSID(gTipeUser) = "KOORD NHS" Or NulSID(gTipeUser) = "ASSKABENG" Or NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                    TOChild Layar.hWnd, frmTPlottingTeknisiNHS.hWnd
                    frmTPlottingTeknisiNHS.show
                    frmTPlottingTeknisiNHS.WindowState = vbMaximized
                Else
                    MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
                End If
            Else
                MsgBox "Project NHS belum diaktifkan", vbInformation
            End If
        Case "mnu1_5_17"
            If Log.p2Didx = 1 Then
                TOChild Layar.hWnd, frmUpdateMtoyota.hWnd
                    frmUpdateMtoyota.show
                    frmUpdateMtoyota.WindowState = vbMaximized
            Else
                MsgBox "Project P2 Didx belum diaktifkan", vbInformation
            End If
        ' Setting
        Case "mnu1_6_1"
            TOChild Layar.hWnd, frmtprofil.hWnd
            frmtprofil.show
            frmtprofil.WindowState = vbMaximized
        Case "mnu1_6_2"
            TOChild Layar.hWnd, frmSetupPetugas.hWnd
            frmSetupPetugas.show
            frmSetupPetugas.WindowState = vbMaximized
        Case "mnu1_6_3"
            oConn.Execute ("BACKUP DATABASE " & gDb & " TO DISK = '" & gFBackup & gDb & "_" & Format(isnow, "DDMMYY") & "-" & Format(isnow, "HHMM") & "_" & gSA & ".NIS'")
            MsgBox "Backup berhasil. File di Folder Server " & gFBackup & gDb & "_" & Format(isnow, "DDMMYY") & "-" & Format(isnow, "HHMM") & ".NIS"
        Case "mnu1_6_4"
            TOChild Layar.hWnd, frmSetupPassword.hWnd
            frmSetupPassword.show
            frmSetupPassword.WindowState = vbMaximized
        Case "mnu1_6_5"
            TOChild Layar.hWnd, frmInfoNoPolisi.hWnd
            frmInfoNoPolisi.show
            frmInfoNoPolisi.WindowState = vbMaximized
        Case "mnu1_6_6"
            TOChild Layar.hWnd, frmInfoBackupHist.hWnd
            frmInfoBackupHist.show
            frmInfoBackupHist.WindowState = vbMaximized
        Case "mnu1_6_7"
            TOChild Layar.hWnd, frmtAkun.hWnd
            frmtAkun.show
            frmtAkun.WindowState = vbMaximized
        Case "mnu1_6_8"
            TOChild Layar.hWnd, frmProfilAAS.hWnd
            frmProfilAAS.show
            frmProfilAAS.WindowState = vbMaximized
        Case "mnu1_6_9"
            TOChild Layar.hWnd, frmtoli.hWnd
            frmtoli.show
            frmtoli.WindowState = vbMaximized
        Case "mnu1_6_10"
            TOChild Layar.hWnd, frmSetup.hWnd
            frmSetup.show
            frmSetup.WindowState = vbMaximized
        Case "mnu1_6_11"
            TOChild Layar.hWnd, frmtmaintenance.hWnd
            frmtmaintenance.show
            frmtmaintenance.WindowState = vbMaximized
        Case "mnu1_6_12"
            TOChild Layar.hWnd, frmtPeriode.hWnd
            frmtPeriode.show
            frmtPeriode.WindowState = vbMaximized
        Case "mnu1_6_13"
            TOChild Layar.hWnd, frmtConfigTMS.hWnd
            frmtConfigTMS.show
            frmtConfigTMS.WindowState = vbMaximized
        ' Registrasi
        Case "mnu1_8_1"
            TOChild Layar.hWnd, frmRegistrasi.hWnd
            frmRegistrasi.show
            frmRegistrasi.WindowState = vbMaximized
        Case "mnu1_8_2"
            TOChild Layar.hWnd, frmBtlRegistrasi.hWnd
            frmBtlRegistrasi.show
            frmBtlRegistrasi.WindowState = vbMaximized
            'frmBatalRegistrasi.Show
            'frmBatalRegistrasi.WindowState = vbMaximized
        'Informasi
        Case "mnu1_9_1" 'History
            TOChild Layar.hWnd, frmInfoHistory.hWnd
            frmInfoHistory.show
            frmInfoHistory.WindowState = vbMaximized
        Case "mnu1_9_2"
            TOChild Layar.hWnd, frmInfoServicePerTgl.hWnd
            frmInfoServicePerTgl.show
            frmInfoServicePerTgl.WindowState = vbMaximized
        Case "mnu1_9_3" 'Day on JOb Progress
            TOChild Layar.hWnd, frmInfoDayOnJob.hWnd
            frmInfoDayOnJob.show
            frmInfoDayOnJob.WindowState = vbMaximized
        Case "mnu1_9_4" 'All Rep
            TOChild Layar.hWnd, frmInfoAllRepair.hWnd
            frmInfoAllRepair.show
            frmInfoAllRepair.WindowState = vbMaximized
        Case "mnu1_9_5" 'batal WO
            TOChild Layar.hWnd, frmInfoBatalWO.hWnd
            frmInfoBatalWO.show
            frmInfoBatalWO.WindowState = vbMaximized
        Case "mnu1_9_6" 'batal Inv
            TOChild Layar.hWnd, frmInfoBatalInv.hWnd
            frmInfoBatalInv.show
            frmInfoBatalInv.WindowState = vbMaximized
        Case "mnu1_9_7" 'batal Inv
            TOChild Layar.hWnd, frmInfoBook.hWnd
            frmInfoBook.show
            frmInfoBook.WindowState = vbMaximized
        Case "mnu1_9_8" 'Hobi
            TOChild Layar.hWnd, frmInfoHobi.hWnd
            frmInfoHobi.show
            frmInfoHobi.WindowState = vbMaximized
        Case "mnu1_9_9" 'Type Kendaraan
            TOChild Layar.hWnd, frmInfoTypeKend.hWnd
            frmInfoTypeKend.show
            frmInfoTypeKend.WindowState = vbMaximized
        Case "mnu1_9_10" 'Per Group To day On Job Progress
            TOChild Layar.hWnd, frmInfoDayOnJobGroup.hWnd
            frmInfoDayOnJobGroup.show
            frmInfoDayOnJobGroup.WindowState = vbMaximized
        Case "mnu1_9_11"
'            frmLihatOrder.Show
'            frmLihatOrder.WindowState = vbMaximized
        Case "mnu1_9_12"
            TOChild Layar.hWnd, frmInfoTypePerTgl.hWnd
            frmInfoTypePerTgl.show
            frmInfoTypePerTgl.WindowState = vbMaximized
        Case "mnu1_9_13"
            TOChild Layar.hWnd, frmInfoTypePerBln.hWnd
            frmInfoTypePerBln.show
            frmInfoTypePerBln.WindowState = vbMaximized
        Case "mnu1_9_14"
            TOChild Layar.hWnd, frmRepoMap.hWnd
            frmRepoMap.show
            frmRepoMap.WindowState = vbMaximized
        Case "mnu1_9_15"
            TOChild Layar.hWnd, frmKonfirmasiSublet.hWnd
            frmKonfirmasiSublet.show
            frmKonfirmasiSublet.WindowState = vbMaximized
        Case "mnu1_9_16"
            TOChild Layar.hWnd, frmValidasiBilling.hWnd
            frmValidasiBilling.show
            frmValidasiBilling.WindowState = vbMaximized
        Case "mnu1_9_17"
            TOChild Layar.hWnd, frmAuditService.hWnd
            frmAuditService.show
            frmAuditService.WindowState = vbMaximized
        Case "mnu1_9_18"
            TOChild Layar.hWnd, frmInfoTagihCuci.hWnd
            frmInfoTagihCuci.show
            frmInfoTagihCuci.WindowState = vbMaximized
        Case "mnu1_9_19"
            TOChild Layar.hWnd, frmInfoPelanggan.hWnd
            frmInfoPelanggan.show
            frmInfoPelanggan.WindowState = vbMaximized
        Case "mnu1_9_20"
            TOChild Layar.hWnd, frmRequestKacab.hWnd
            frmRequestKacab.show
            frmRequestKacab.WindowState = vbMaximized
        Case "mnu1_9_21"
            TOChild Layar.hWnd, frmInfoPelangganRetention.hWnd
            frmInfoPelangganRetention.show
            frmInfoPelangganRetention.WindowState = vbMaximized
        Case "mnu1_9_22"
            TOChild Layar.hWnd, frmInfoPelangganRecoveryMapping.hWnd
            frmInfoPelangganRecoveryMapping.show
            frmInfoPelangganRecoveryMapping.WindowState = vbMaximized
        Case "mnu1_9_23"
'            If gTipeUser = "MRS" Then
'
'            End If
            TOChild Layar.hWnd, frmReportLS.hWnd
            frmReportLS.show
            frmReportLS.WindowState = vbMaximized
            
        Case "mnu1_9_24"
            If Log.hansaRenkei = 1 Then
                TOChild Layar.hWnd, frmInjectMEstimasi.hWnd
                frmInjectMEstimasi.show
                frmInjectMEstimasi.WindowState = vbMaximized
            Else
                MsgBox "Program hansa renkei belum diaktifkan.", vbInformation
            End If
        Case "mnu1_9_25"
            If Log.vin2vinRev1 = 1 Then
                TOChild Layar.hWnd, frmInfoRetensiUnitRetail.hWnd
                frmInfoRetensiUnitRetail.show
                frmInfoRetensiUnitRetail.WindowState = vbMaximized
            Else
                MsgBox "Program vin by vin rev 1 belum diaktifkan.", vbInformation
            End If
        Case "mnu1_9_26"
            If Log.NHS = 1 Then
                If NulSID(gTipeUser) = "KOORD NHS" Or NulSID(gTipeUser) = "ASSKABENG" Or NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                    TOChild Layar.hWnd, frmValidasiPekerjaanNHS.hWnd
                    frmValidasiPekerjaanNHS.show
                    frmValidasiPekerjaanNHS.WindowState = vbMaximized
                Else
                    MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
                End If
            Else
                MsgBox "Project NHS belum diaktifkan.", vbInformation
            End If
        Case "mnu1_9_27"
            If NulSID(gTipeUser) = "ASSKABENG" Or NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoValidasiPartTMS.hWnd
                frmInfoValidasiPartTMS.show
                frmInfoValidasiPartTMS.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
        Case "mnu1_9_28"
            If NulSID(gTipeUser) = "ASSKABENG" Or NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoValidasiPartTMSRetur.hWnd
                frmInfoValidasiPartTMSRetur.show
                frmInfoValidasiPartTMSRetur.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
        Case "mnu1_29_1"
            If NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoPerhitunganKPICabang.hWnd
                frmInfoPerhitunganKPICabang.show
                frmInfoPerhitunganKPICabang.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
        Case "mnu1_29_2"
            If NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoPerhitunganKPIManpower.hWnd
                frmInfoPerhitunganKPIManpower.show
                frmInfoPerhitunganKPIManpower.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
        Case "mnu1_29_3"
            If NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoPenilaianKPIProses.hWnd
                frmInfoPenilaianKPIProses.show
                frmInfoPenilaianKPIProses.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
        Case "mnu1_29_4"
            If NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoKetersediaanInsentifCabang.hWnd
                frmInfoKetersediaanInsentifCabang.show
                frmInfoKetersediaanInsentifCabang.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
        Case "mnu1_29_5"
            If NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoPerhitunganKPIManpowerProses.hWnd
                frmInfoPerhitunganKPIManpowerProses.show
                frmInfoPerhitunganKPIManpowerProses.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
        Case "mnu1_29_6"
            If NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoKetersediaanInsentifProduk.hWnd
                frmInfoKetersediaanInsentifProduk.show
                frmInfoKetersediaanInsentifProduk.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
        Case "mnu1_29_7"
            If NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoKetersediaanInsentifProdukNonSA.hWnd
                frmInfoKetersediaanInsentifProdukNonSA.show
                frmInfoKetersediaanInsentifProdukNonSA.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
        Case "mnu1_29_8"
            If NulSID(gTipeUser) = "KABENG" Or NulSID(gTipeUser) = "ADMIN" Or NulSID(gTipeUser) = "HO" Then
                TOChild Layar.hWnd, frmInfoInsentifCalculation.hWnd
                frmInfoInsentifCalculation.show
                frmInfoInsentifCalculation.WindowState = vbMaximized
            Else
                MsgBox "Tipe User Anda tidak diperkenankan mengakses menu ini.", vbInformation
            End If
            
        'Laporan Order
        Case "mnu1_10_1_1"
            'SO Masuk
            TOChild Layar.hWnd, frmLapHarianSOMasuk.hWnd
            frmLapHarianSOMasuk.show
            frmLapHarianSOMasuk.WindowState = vbMaximized
        Case "mnu1_10_1_2"
            'Unit Masuk/SA
            TOChild Layar.hWnd, frmLapHarianSOMasukPerSA.hWnd
            frmLapHarianSOMasukPerSA.show
            frmLapHarianSOMasukPerSA.WindowState = vbMaximized
        Case "mnu1_10_1_3"
            'SO dlm Proses
            TOChild Layar.hWnd, frmLapBulanSemuaSOProses.hWnd
            frmLapBulanSemuaSOProses.show
            frmLapBulanSemuaSOProses.WindowState = vbMaximized
        Case "mnu1_10_1_4"
            'Audit bengkel
            TOChild Layar.hWnd, frmLapBulanWoInv.hWnd
            frmLapBulanWoInv.show
            frmLapBulanWoInv.WindowState = vbMaximized
        Case "mnu1_10_1_5"
            'Unit
            TOChild Layar.hWnd, frmLapBulanJumlahKaliMasuk.hWnd
            frmLapBulanJumlahKaliMasuk.show
            frmLapBulanJumlahKaliMasuk.WindowState = vbMaximized
        Case "mnu1_10_1_6"
            'Unit
            TOChild Layar.hWnd, frmLapBulanSOPerSA.hWnd
            frmLapBulanSOPerSA.show
            frmLapBulanSOPerSA.WindowState = vbMaximized
        Case "mnu1_10_1_7"
            'Unit
            TOChild Layar.hWnd, frmLapBulanOrderBengkel.hWnd
            frmLapBulanOrderBengkel.show
            frmLapBulanOrderBengkel.WindowState = vbMaximized
        'Laporan Invoice
        Case "mnu1_10_2_1"
            TOChild Layar.hWnd, frmLapPenjService.hWnd
            frmLapPenjService.show
            frmLapPenjService.WindowState = vbMaximized
        Case "mnu1_10_2_2"
            TOChild Layar.hWnd, frmLapPenjServiceHead.hWnd
            frmLapPenjServiceHead.show
            frmLapPenjServiceHead.WindowState = vbMaximized
        Case "mnu1_10_2_3"
            'invoice/SA
            TOChild Layar.hWnd, frmLapHarianPenjualanSA.hWnd
            frmLapHarianPenjualanSA.show
            frmLapHarianPenjualanSA.WindowState = vbMaximized
        Case "mnu1_10_2_4"
            'Penjualan
            TOChild Layar.hWnd, frmLapHarianPenjualan.hWnd
            frmLapHarianPenjualan.show
            frmLapHarianPenjualan.WindowState = vbMaximized
        Case "mnu1_10_2_5"
            'Detail Penjualan
            TOChild Layar.hWnd, frmLapHarianPenjualanBengkel.hWnd
            frmLapHarianPenjualanBengkel.show
            frmLapHarianPenjualanBengkel.WindowState = vbMaximized
        Case "mnu1_10_2_6"
            'Pasif/Hari
            TOChild Layar.hWnd, frmLapBulanPasif.hWnd
            frmLapBulanPasif.show
            frmLapBulanPasif.WindowState = vbMaximized
        Case "mnu1_10_2_7"
            'Cust IRC 1000/Hr
            TOChild Layar.hWnd, frmLapBulanPasifIRC1000.hWnd
            frmLapBulanPasifIRC1000.show
            frmLapBulanPasifIRC1000.WindowState = vbMaximized
        Case "mnu1_10_2_8"
            'Cust IRC 5000/Hr
            TOChild Layar.hWnd, frmLapBulanPasifIRC5000.hWnd
            frmLapBulanPasifIRC5000.show
            frmLapBulanPasifIRC5000.WindowState = vbMaximized
'        Case "mnu1_10_12"
'            'Aktivitas Bengkel
'            '(belum)
        Case "mnu1_10_2_9"
            'Bln penjualan
'            frmLapPenjServiceBln.Show -> pakai kertas besar
            TOChild Layar.hWnd, frmLapBulanPenjualan.hWnd
            frmLapBulanPenjualan.show
            frmLapBulanPenjualan.WindowState = vbMaximized
        Case "mnu1_10_2_10"
            'Detail Penjualan
            TOChild Layar.hWnd, frmLapBulanPenjualanBengkel.hWnd
            frmLapBulanPenjualanBengkel.show
            frmLapBulanPenjualanBengkel.WindowState = vbMaximized
        Case "mnu1_10_2_11"
            'Regular check
            TOChild Layar.hWnd, frmLapBulanMointorUnitRC.hWnd
            frmLapBulanMointorUnitRC.show
            frmLapBulanMointorUnitRC.WindowState = vbMaximized
        Case "mnu1_10_2_12"
            'rate Group
            TOChild Layar.hWnd, frmLapBulanFlateRateGroup.hWnd
            frmLapBulanFlateRateGroup.show
            frmLapBulanFlateRateGroup.WindowState = vbMaximized
        Case "mnu1_10_2_13"
            'rate Mekanik
            TOChild Layar.hWnd, frmLapBulanFlateRateMekanik.hWnd
            frmLapBulanFlateRateMekanik.show
            frmLapBulanFlateRateMekanik.WindowState = vbMaximized
        Case "mnu1_10_2_14"
            'IRC 1000 pd bln
            TOChild Layar.hWnd, frmLapBulanIRC1000.hWnd
            frmLapBulanIRC1000.show
            frmLapBulanIRC1000.WindowState = vbMaximized
        Case "mnu1_10_2_15"
            'IRC 5000 pd bln
            TOChild Layar.hWnd, frmLapBulanIRC5000.hWnd
            frmLapBulanIRC5000.show
            frmLapBulanIRC5000.WindowState = vbMaximized
        Case "mnu1_10_2_16"
            'Regular cek per No Polisi
            TOChild Layar.hWnd, frmLapBulanMointorUnitRCPol.hWnd
            frmLapBulanMointorUnitRCPol.show
            frmLapBulanMointorUnitRCPol.WindowState = vbMaximized
        Case "mnu1_10_2_17"
            'Regular cek per No Polisi
            TOChild Layar.hWnd, frmLapBulanPenjualanSA.hWnd
            frmLapBulanPenjualanSA.show
            frmLapBulanPenjualanSA.WindowState = vbMaximized
        Case "mnu1_10_2_18"
            'Regular cek per No Polisi
            TOChild Layar.hWnd, frmLapBulanPerJenisP.hWnd
            frmLapBulanPerJenisP.show
            frmLapBulanPerJenisP.WindowState = vbMaximized
        Case "mnu1_10_2_19"
            'Regular cek per No Polisi
            TOChild Layar.hWnd, frmLapBulanRtj.hWnd
            frmLapBulanRtj.show
            frmLapBulanRtj.WindowState = vbMaximized
        Case "mnu1_10_2_20"
            'Regular cek per No Polisi
            TOChild Layar.hWnd, frmLapBulanPenjualanPerPlg.hWnd
            frmLapBulanPenjualanPerPlg.show
            frmLapBulanPenjualanPerPlg.WindowState = vbMaximized
        Case "mnu1_10_2_21"
            'Regular cek per No Polisi
            TOChild Layar.hWnd, frmLapHarianPenjualanPerPlg.hWnd
            frmLapHarianPenjualanPerPlg.show
            frmLapHarianPenjualanPerPlg.WindowState = vbMaximized
            
        ' LAPORAN JML UNIT
        Case "mnu1_10_3_1"
            TOChild Layar.hWnd, frmRepoJmlUnit.hWnd
            frmRepoJmlUnit.show
            frmRepoJmlUnit.WindowState = vbMaximized
        Case "mnu1_10_3_2"
            TOChild Layar.hWnd, frmRepoSOdalamProses.hWnd
            frmRepoSOdalamProses.show
            frmRepoSOdalamProses.WindowState = vbMaximized
        Case "mnu1_10_3_3"
            TOChild Layar.hWnd, frmRepoUnitOut.hWnd
            frmRepoUnitOut.show
            frmRepoUnitOut.WindowState = vbMaximized
            
        ' LAPORAN REVENUE UNIT
        Case "mnu1_10_4_1"
'            frmRepoJmlInv.Show
'            frmRepoJmlInv.WindowState = vbMaximized
            TOChild Layar.hWnd, frmRepoInvRinci.hWnd
            frmRepoInvRinci.show
            frmRepoInvRinci.WindowState = vbMaximized
        Case "mnu1_10_4_2"
            TOChild Layar.hWnd, frmRepoInvTot.hWnd
            frmRepoInvTot.show
            frmRepoInvTot.WindowState = vbMaximized
        Case "mnu1_10_4_3"
            TOChild Layar.hWnd, frmRepoInvoiceDiskon.hWnd
            frmRepoInvoiceDiskon.show
            frmRepoInvoiceDiskon.WindowState = vbMaximized
            
        ' LAPORAN MEKANIK
        Case "mnu1_10_5_1"
            TOChild Layar.hWnd, frmRepoProdTech.hWnd
            frmRepoProdTech.show
            frmRepoProdTech.WindowState = vbMaximized
        Case "mnu1_10_5_2"
            TOChild Layar.hWnd, frmRepoHarianTech.hWnd
            frmRepoHarianTech.show
            frmRepoHarianTech.WindowState = vbMaximized
            
        ' LAPORAN REKAP
        Case "mnu1_10_6_1"
            TOChild Layar.hWnd, frmRepoHarian.hWnd
            frmRepoHarian.show
            frmRepoHarian.WindowState = vbMaximized
        Case "mnu1_10_6_2"
            TOChild Layar.hWnd, frmRepoBulanan.hWnd
            frmRepoBulanan.show
            frmRepoBulanan.WindowState = vbMaximized
        Case "mnu1_10_6_3"
            TOChild Layar.hWnd, frmRepo3DFU.hWnd
            frmRepo3DFU.show
            frmRepo3DFU.WindowState = vbMaximized
        Case "mnu1_10_6_4"
            TOChild Layar.hWnd, frmCetakKonfirmasi.hWnd
            frmCetakKonfirmasi.show
            frmCetakKonfirmasi.WindowState = vbMaximized
        Case "mnu1_10_6_5"
            TOChild Layar.hWnd, frmRepoICSK.hWnd
            frmRepoICSK.show
            frmRepoICSK.WindowState = vbMaximized
        Case "mnu1_10_6_6"
            TOChild Layar.hWnd, frmRepoPenyerahan.hWnd
            frmRepoPenyerahan.show
            frmRepoPenyerahan.WindowState = vbMaximized
        Case "mnu1_10_6_7"
            TOChild Layar.hWnd, frmRepoCuci.hWnd
            frmRepoCuci.show
            frmRepoCuci.WindowState = vbMaximized
        Case "mnu1_10_6_8"
            TOChild Layar.hWnd, frmCetakKonfirmasiCuci.hWnd
            frmCetakKonfirmasiCuci.show
            frmCetakKonfirmasiCuci.WindowState = vbMaximized
        Case "mnu1_10_6_9"
            TOChild Layar.hWnd, frmRepoTestDrive.hWnd
            frmRepoTestDrive.show
            frmRepoTestDrive.WindowState = vbMaximized
        Case "mnu1_10_6_10"
            If Log.logRetention = 1 Then
                If Log.vin2vinRev1 = 1 Then
                    TOChild Layar.hWnd, frmRepoSBRateRetentionRev1.hWnd
                    frmRepoSBRateRetentionRev1.show
                    frmRepoSBRateRetentionRev1.WindowState = vbMaximized
                Else
                    TOChild Layar.hWnd, frmRepoSBRateRetention.hWnd
                    frmRepoSBRateRetention.show
                    frmRepoSBRateRetention.WindowState = vbMaximized
                End If
            Else
                TOChild Layar.hWnd, frmRepoSBRate.hWnd
                frmRepoSBRate.show
                frmRepoSBRate.WindowState = vbMaximized
            End If
        Case "mnu1_10_6_11"
            If Log.discKacab = 1 Then
                TOChild Layar.hWnd, frmRepoLockDiscount.hWnd
                frmRepoLockDiscount.show
                frmRepoLockDiscount.WindowState = vbMaximized
                
            End If
        Case "mnu1_10_6_12"
            If Log.standarisasiSublet = 1 Then
                TOChild Layar.hWnd, frmRepoSentralSublet.hWnd
                frmRepoSentralSublet.show
                frmRepoSentralSublet.WindowState = vbMaximized
                
            End If
        Case "mnu1_10_6_13"
            If Log.standarisasiSublet = 1 Then
                TOChild Layar.hWnd, frmRepoSentralMstSublet.hWnd
                frmRepoSentralMstSublet.show
                frmRepoSentralMstSublet.WindowState = vbMaximized
                
            End If
            
        ' LAPORAN MRS
        Case "mnu1_10_7_1"
            TOChild Layar.hWnd, frmRepoMrsDatang.hWnd
            frmRepoMrsDatang.show
            frmRepoMrsDatang.WindowState = vbMaximized
        Case "mnu1_10_7_2"
            frmRepoMrsTakTerhubung.show
            frmRepoMrsTakTerhubung.WindowState = vbMaximized
        Case "mnu1_10_7_3"
            frmListSBI.show
            frmListSBI.WindowState = vbMaximized
        Case "mnu1_10_7_4"
            frmMapping.show
            frmMapping.WindowState = vbMaximized
        Case "mnu1_10_7_5"
            frmRepoMrsRecordVoiceOfCustomer.show
            frmRepoMrsRecordVoiceOfCustomer.WindowState = vbMaximized
        Case "mnu1_10_7_6"
            TOChild Layar.hWnd, frmRepoHansaRenkei.hWnd
            frmRepoHansaRenkei.show
            frmRepoHansaRenkei.WindowState = vbMaximized
        Case "mnu1_10_7_7"
            TOChild Layar.hWnd, frmRepoFunnelMRA.hWnd
            frmRepoFunnelMRA.show
            frmRepoFunnelMRA.WindowState = vbMaximized
            
        ' LAPORAN MONITORING
        Case "mnu1_10_8_1"
            frmRepoSupply.show
            frmRepoSupply.WindowState = vbMaximized
        Case "mnu1_10_8_2"
            frmRepoInvRetur.show
            frmRepoInvRetur.WindowState = vbMaximized
        Case "mnu1_10_8_3"
            frmRepoMonitoringMOU.show
            frmRepoMonitoringMOU.WindowState = vbMaximized
        Case "mnu1_10_8_4"
            frmRepoLockAR.show
            frmRepoLockAR.WindowState = vbMaximized
'        Case "mnu1_10_8_5"
'            frmRepoSupportTL.show
'            frmRepoSupportTL.WindowState = vbMaximized
        Case "mnu1_10_8_5"
            frmHistoryServiceAll.show
            frmHistoryServiceAll.WindowState = vbMaximized
            
        Case "mnu1_11"
            frmAbout.show
        
        ' Menu Antrian
        Case "mnu1_12_1"
            antrian = "Cetak"
            'frmAntrian.Show
            'frmAntrian2.Show
        Case "mnu1_12_2"
            antrian = "Panggil"
            'frmAntrian.Show
            'frmAntrian2.Top = Layar.ActiveBar1.Top
            'frmAntrian2.Left = 0 'Layar.imgCbg.Left + Layar.imgCbg.Width + 20
            'frmAntrian2.Show
            
        Case "mnu1_13_1"
            
            If Log.logRetention = 1 Then
                If Log.hansaRenkei = 1 Then
                    Dim fNewHR As New FormMainMRSRetentionHansaRenkei
                    fNewHR.Top = 0
                    fNewHR.Left = 0
                    fNewHR.show
                    fNewHR.WindowState = vbMaximized
                Else
                    Dim fNew As New FormMainMRSRetention
                    fNew.Top = 0
                    fNew.Left = 0
                    fNew.show
                    fNew.WindowState = vbMaximized
                End If
            Else
                If Log.hansaRenkei = 1 Then
                    Dim fHR As New FormMainMRSHansaRenkei
                    fHR.Top = 0
                    fHR.Left = 0
                    fHR.show
                    fHR.WindowState = vbMaximized
                Else
                    Dim f As New FormMainMRS
                    f.Top = 0
                    f.Left = 0
                    f.show
                    f.WindowState = vbMaximized
                End If
                
            End If
        Case "mnu1_13_2"
            If gNIK = "" Then
                MsgBox "Silahkan lengkapi data NIK anda", vbInformation
            Else
                Dim fRHR As New FormRemainderServiceHansaRenkei
                fRHR.WindowState = vbMaximized
                fRHR.show 'vbModal
                ''FormRemainderServiceHansaRenkei.WindowState = vbMaximized
            End If
        Case "mnu1_13_3"
            If Log.hansaRenkei = 1 Then
                If Trim(UCase(gTipeUser)) = "SA" Then
                    Dim fSAS As New frmInjectDataSAtoSales
                    fSAS.WindowState = vbMaximized
                    fSAS.show
                Else
                    MsgBox "Tipe SA yang diperuntukan menu ini.", vbInformation
                End If
            Else
                MsgBox "Project Hansa Renkei belum diaktifkan.", vbInformation
            End If
        Case "mnu1_13_4"
            If Log.hansaRenkeiBP = 1 Then
                If Trim(UCase(gTipeUser)) = "SA" Then
                    Dim fBP As New frmInjectDataBookingBP
                    fBP.WindowState = vbMaximized
                    fBP.show
                Else
                    MsgBox "Tipe SA yang diperuntukan menu ini.", vbInformation
                End If
            Else
                MsgBox "Project Hansa Renkei BP belum diaktifkan.", vbInformation
            End If
        Case "mnu1_13_5"
            If Log.hansaRenkeiBP = 1 Then
                Dim fRA As New frmRetentionAsuransi
                fRA.WindowState = vbMaximized
                fRA.show
            Else
                MsgBox "Project Hansa Renkei BP belum diaktifkan.", vbInformation
            End If
        Case "mnu1_13_6"
            If Log.hansaRenkeiBP = 1 Then
                Dim fRAH As New frmRetentionAsuransiHasil
                fRAH.WindowState = vbMaximized
                fRAH.show
            Else
                MsgBox "Project Hansa Renkei BP belum diaktifkan.", vbInformation
            End If
        Case "mnu1_13_7"
            If Log.p2Didx = 1 Then
                Dim fReminderMt As New FormRemainderServiceMToyota
                fReminderMt.WindowState = vbMaximized
                fReminderMt.show
            Else
                MsgBox "Project DIDX Phase 2 belum diaktifkan.", vbInformation
            End If
        ' Menu HO
        ' MASTER HO
        Case "mnu1_21_1"
            frmHOJobOrder.show
            frmHOJobOrder.WindowState = vbMaximized
        Case "mnu1_21_2"
            frmHOKelas.show
            frmHOKelas.WindowState = vbMaximized
        Case "mnu1_21_3"
            frmHOListMOU.show
            frmHOListMOU.WindowState = vbMaximized
        Case "mnu1_21_4"
            frmHOProgram.show
            frmHOProgram.WindowState = vbMaximized
        Case "mnu1_21_5"
            frmHOWarranty.show
            frmHOWarranty.WindowState = vbMaximized
        Case "mnu1_21_6"
            frmtDealer.show
            frmtDealer.WindowState = vbMaximized
'             frmHOPerbaikanNoRangka.Show
'             frmHOPerbaikanNoRangka.WindowState = vbMaximized
             
        Case "mnu1_21_7"
            frmHOPartPacket.show
            frmHOPartPacket.WindowState = vbMaximized
        Case "mnu1_21_8"
            frmHOFlateRate.show
            frmHOFlateRate.WindowState = vbMaximized
        Case "mnu1_21_9"
            frmtModel.show
            frmtModel.WindowState = vbMaximized
        Case "mnu1_21_10"
            frmtJPekerjaan.show
            frmtJPekerjaan.WindowState = vbMaximized
        Case "mnu1_21_11"
            frmHOEvent.show
            frmHOEvent.WindowState = vbMaximized
        Case "mnu1_21_12"
            frmHOPeriode.show
            frmHOPeriode.WindowState = vbMaximized
        Case "mnu1_21_13"
            '' 20181018DEV1 | DEV | DEVELOPMENT  | PENAMBAHAN FLAG UNTUK STANDARISASI SUBLET
            If Log.standarisasiSublet = 1 Then
                frmHOSupplierSublet.show
                frmHOSupplierSublet.WindowState = vbMaximized
            Else
                MsgBox "Flag standarisasi sublet belum diaktifkan", vbInformation
                Exit Sub
            End If
            
        Case "mnu1_21_14"
            '' 20181018DEV1 | DEV | DEVELOPMENT  | PENAMBAHAN FLAG UNTUK STANDARISASI SUBLET
            If Log.standarisasiSublet = 1 Then
                frmHOSublet.show
                frmHOSublet.WindowState = vbMaximized
            Else
                MsgBox "Flag standarisasi sublet belum diaktifkan", vbInformation
                Exit Sub
            End If
            
        Case "mnu1_21_15"
            '' 20181018DEV1 | DEV | DEVELOPMENT  | PENAMBAHAN FLAG UNTUK STANDARISASI SUBLET
            If Log.standarisasiSublet = 1 Then
                frmHOPromoSublet.show
                frmHOPromoSublet.WindowState = vbMaximized
            Else
                MsgBox "Flag standarisasi sublet belum diaktifkan", vbInformation
                Exit Sub
            End If
         
        Case "mnu1_21_16"
            frmHOSettingSCCNormal.show
            frmHOSettingSCCNormal.WindowState = vbMaximized
        Case "mnu1_21_17"
            frmHOSettingSCCKhusus.show
            frmHOSettingSCCKhusus.WindowState = vbMaximized
        Case "mnu1_272_1"
            If Log.rsA = 1 Then
                frmHOMasterProblemRSA.show
                frmHOMasterProblemRSA.WindowState = vbMaximized
            Else
                MsgBox "Project RSA belum diaktifkan.", vbInformation
                Exit Sub
            End If
        Case "mnu1_272_2"
            If Log.rsA = 1 Then
                frmHOMasterKomponenRSA.show
                frmHOMasterKomponenRSA.WindowState = vbMaximized
            Else
                MsgBox "Project RSA belum diaktifkan.", vbInformation
                Exit Sub
            End If
        Case "mnu1_272_3"
            If Log.rsA = 1 Then
                frmHOMasterEmailTL.show
                frmHOMasterEmailTL.WindowState = vbMaximized
            Else
                MsgBox "Project RSA belum diaktifkan.", vbInformation
                Exit Sub
            End If
        Case "mnu1_27_1"
            If Log.rsA = 1 Then
                frmHOSupportTL.show
                frmHOSupportTL.WindowState = vbMaximized
            Else
                MsgBox "Project RSA belum diaktifkan.", vbInformation
                Exit Sub
            End If
        Case "mnu1_28_1"
                frmHOInsentifTipeMp.show
                frmHOInsentifTipeMp.WindowState = vbMaximized
        Case "mnu1_28_2"
                frmHOInsentifMasterKPI.show
                frmHOInsentifMasterKPI.WindowState = vbMaximized
        Case "mnu1_28_3"
                frmHOInsentifPlottingKPIBobot.show
                frmHOInsentifPlottingKPIBobot.WindowState = vbMaximized
        Case "mnu1_28_4"
                frmHOInsentifKonversiKPI.show
                frmHOInsentifKonversiKPI.WindowState = vbMaximized
        Case "mnu1_28_5"
                frmHOInsentifBobotSparepart.show
                frmHOInsentifBobotSparepart.WindowState = vbMaximized
        Case "mnu1_28_6"
                frmHOInsentifBobotSublet.show
                frmHOInsentifBobotSublet.WindowState = vbMaximized
        Case "mnu1_28_7"
                frmHOInsentifDiskonAlokasi.show
                frmHOInsentifDiskonAlokasi.WindowState = vbMaximized
        Case "mnu1_28_8"
                frmHOInsentifKonversiKPIProses.show
                frmHOInsentifKonversiKPIProses.WindowState = vbMaximized
        Case "mnu1_28_9"
                frmHOInsentifSetResultKPICabang.show
                frmHOInsentifSetResultKPICabang.WindowState = vbMaximized
                
        Case "mnu1_21_18"
            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
                frmHOProgramFixPrice.show
                frmHOProgramFixPrice.WindowState = vbMaximized
            Else
                MsgBox "Project Multiple Program belum diaktifkan.", vbInformation
                Exit Sub
            End If
        Case "mnu1_21_19"
            If Log.freepart = 1 Then
'                frmHOProgramFOPNew.show
'                frmHOProgramFOPNew.WindowState = vbMaximized
                frmHOProgramFOPSummary.show
                frmHOProgramFOPSummary.WindowState = vbMaximized
            Else
                MsgBox "Project Free Part TAM belum diaktifkan.", vbInformation
                Exit Sub
            End If
        Case "mnu1_21_20"
            If Log.freepart = 1 Then
                frmHOSubletPacket.show
                frmHOSubletPacket.WindowState = vbMaximized
            Else
                MsgBox "Project Free Part TAM belum diaktifkan.", vbInformation
                Exit Sub
            End If
            
        ' LAPORAN HO
        Case "mnu1_22_1"
            frmRepoHOBulanan.show
            frmRepoHOBulanan.WindowState = vbMaximized
        Case "mnu1_22_5"
            frmRepoHOWar.show
            frmRepoHOWar.WindowState = vbMaximized
        Case "mnu1_22_6"
            frmRepoSBRateHO.show
            frmRepoSBRateHO.WindowState = vbMaximized
        Case "mnu1_22_7"
            frmRepoSupportTLHO.show
            frmRepoSupportTLHO.WindowState = vbMaximized
        Case "mnu1_22_8"
            If Log.standarisasiSublet = 1 Then
                TOChild Layar.hWnd, frmRepoSentralSublet.hWnd
                frmRepoSentralSublet.show
                frmRepoSentralSublet.WindowState = vbMaximized
                
            End If
            
        ' TARGET CABANG
        Case "mnu1_23_1"
            frmHOTargetUnit.show
            frmHOTargetUnit.WindowState = vbMaximized
        Case "mnu1_23_3"
            If Log.standarisasiSublet = 1 Then
                TOChild Layar.hWnd, frmRepoSentralSublet.hWnd
                frmRepoSentralSublet.show
                frmRepoSentralSublet.WindowState = vbMaximized
                
            End If
        Case "mnu1_23_4"
            If Log.standarisasiSublet = 1 Then
                TOChild Layar.hWnd, frmRepoSentralMstSublet.hWnd
                frmRepoSentralMstSublet.show
                frmRepoSentralMstSublet.WindowState = vbMaximized
                
            End If
        
        Case "mnu1_23_5"
            If Log.vinRetention = 1 Then
                If Log.vin2vinRev1 = 1 Then
                    TOChild Layar.hWnd, frmHOSBRententionRev1.hWnd
                    frmHOSBRententionRev1.show
                    frmHOSBRententionRev1.WindowState = vbMaximized
                Else
                    TOChild Layar.hWnd, frmHOSBRentention.hWnd
                    frmHOSBRentention.show
                    frmHOSBRentention.WindowState = vbMaximized
                End If
            Else
                MsgBox "Flag vin Retention belum diaktifkan", vbInformation
            End If
        Case "mnu1_23_6"
            If Log.vinRetention = 1 Then
                TOChild Layar.hWnd, frmHOMappingCustomer.hWnd
                frmHOMappingCustomer.show
                frmHOMappingCustomer.WindowState = vbMaximized
            Else
                MsgBox "Flag vin Retention belum diaktifkan", vbInformation
            End If
        ' SETTING CABANG
        Case "mnu1_24_1"
            frmtprofil.show
            frmtprofil.WindowState = vbMaximized
        Case "mnu1_24_2"
            frmSetupPetugas.show
            frmSetupPetugas.WindowState = vbMaximized
        Case "mnu1_24_3"
            frmSetupPassword.show
            frmSetupPassword.WindowState = vbMaximized
            
        ' PERBAIKAN DATA SERVICE
        
        Case "mnu1_25_1"
            ' Perbaikan no Rangka.
            frmHOPerbaikanNoRangka.show
            frmHOPerbaikanNoRangka.WindowState = vbMaximized
        Case "mnu1_25_2"
            ' Perbaikan KM Service
            frmHOPerbaikanKMService.show
            frmHOPerbaikanKMService.WindowState = vbMaximized
        Case "mnu1_25_3"
            ' Perbaikan Nama Mekanik
            frmHOPerbaikanNamaMekanik.show
            frmHOPerbaikanNamaMekanik.WindowState = vbMaximized
        Case "mnu1_25_4"
            ' Perbaikan Jenis Service
            frmHOPerbaikanJenisService.show
            frmHOPerbaikanJenisService.WindowState = vbMaximized
        Case "mnu1_25_5"
            ' perbaikan
            frmHOPerbaikanPanggilanService.show
            frmHOPerbaikanPanggilanService.WindowState = vbMaximized
        Case "mnu1_25_6"
            ' perbaikan
            frmHOPerbaikanWarranty.show
            frmHOPerbaikanWarranty.WindowState = vbMaximized
        Case "mnu1_25_7"
            ' perbaikan
            frmHOPerbaikanNoRangkaTWC.show
            frmHOPerbaikanNoRangkaTWC.WindowState = vbMaximized
                
        '------ MASTER WOS -----------
        Case "mnu1_26_1"
            frmHOTT1.show
            frmHOTT1.WindowState = vbMaximized
        Case "mnu1_26_2"
            frmHOTT2.show
            frmHOTT2.WindowState = vbMaximized
        Case "mnu1_26_3"
            frmHOTFlateRate.show
            frmHOTFlateRate.WindowState = vbMaximized
        Case "mnu1_26_4"
            frmHOTSublet.show
            frmHOTSublet.WindowState = vbMaximized
            
        Case "mnu1_7"
            If MsgBox("Anda yakin mau keluar dari program ini ?", vbQuestion + vbYesNo, "Close Program") = vbYes Then
                On Error Resume Next
                Kill Left(App.Path, 1) & ":\*.tmp"
                Call saveLogOutUser(gSA, ComptName)
                Call closeAllConnection
                End
            End If
            
        End Select
    End If
End Sub

Private Sub MDIForm_Load()
1         onTimer = True
          Dim vNama As String
          Dim vBand As String
          Dim vNmDrive As String
2         RemoveMenus Me, , , , , , True, True
3         Me.Timer1.Enabled = True
4         gTglTampung = gTglServer
5         pic4text.Top = StatusBar1.Top - pic4text.Height
6         If gLevel_Ptg <> 1 Then
              'hiden dulu
7             Set oRs = objO.FetchData("select * from t_menu WITH (NOLOCK) order by fn_urut")
8             If Not oRs.EOF Then
9                 oRs.MoveFirst
10                Do While Not oRs.EOF
11                    vNama = Trim(oRs("fs_nama"))
12                    titik = 0
13                    For i = 1 To Len(vNama)
14                        If Mid(vNama, i, 1) = "_" Then
15                            titik = i
16                        End If
17                    Next
18                    vBand = Trim(Left(vNama, titik - 1))
19                    vBand = "pop" + Right(vBand, (Len(vBand) - 3))
                     
20                    ActiveBar1.Bands(vBand).Tools(vNama).Visible = False
21                    oRs.MoveNext
22                Loop
23            End If
24            oRs.Close
              Dim IndexSpeed As Integer
25            IndexSpeed = 1
              'munculkan lagi petugas
26            Set oRs = objO.FetchData("select * from t_akses WITH (NOLOCK) where Left(akses,5)='" & (AddSpace(gSA, 5, 1, " ")) & "' order by fn_induk")
27            If Not oRs.BOF Then
28                oRs.MoveFirst
29                Do While Not oRs.EOF
30                    vNama = Mid((oRs("akses")), 6, Len((oRs("akses"))) - 5)
31                    titik = 0
32                    For i = 1 To Len(vNama)
33                        If Mid(vNama, i, 1) = "_" Then
34                            titik = i
35                        End If
36                    Next
37                    vBand = Trim(Left(vNama, titik - 1))
38                    vBand = "pop" + Right(vBand, (Len(vBand) - 3))

39                    If IndexSpeed <= 30 And Right(vNama, 1) = "s" Then
40                        vNama = Left(vNama, Len(vNama) - 1)

41                        ActiveBar1.Bands("Band").Tools("Speed" + Trim(str(IndexSpeed))).Name = _
                              ActiveBar1.Bands(vBand).Tools(vNama).Name

42                        ActiveBar1.Bands("Band").Tools(vNama).SetPicture _
                              0, _
                              ActiveBar1.Bands(vBand).Tools(vNama).GetPicture(0), _
                              &HC0C0C0

43                        ActiveBar1.Bands("Band").Tools(vNama).ToolTipText = _
                              ActiveBar1.Bands(vBand).Tools(vNama).ToolTipText

44                        ActiveBar1.Bands("Band").Tools(vNama).SubBand = _
                              ActiveBar1.Bands(vBand).Tools(vNama).SubBand

45                        IndexSpeed = IndexSpeed + 1
46                    End If
                    
47                    ActiveBar1.Bands(vBand).Tools(vNama).Visible = True
48                    oRs.MoveNext
49                Loop
50            End If
51            oRs.Close

              'munculkan lagi sama dengan
52            Set oRs = objO.FetchData("select * from t_akses WITH (NOLOCK) where Left(akses,5)='" & (AddSpace(gKd_Sama, 5, 1, " ")) & "' order by fn_induk")
53            If Not oRs.BOF Then
54                oRs.MoveFirst
55                Do While Not oRs.EOF
56                    vNama = Mid((oRs!akses), 6, Len((oRs!akses)) - 5)
57                    titik = 0
58                    For i = 1 To Len(vNama)
59                        If Mid(vNama, i, 1) = "_" Then
60                            titik = i
61                        End If
62                    Next
63                    vBand = Trim(Left(vNama, titik - 1))
64                    vBand = "pop" + Right(vBand, (Len(vBand) - 3))
                      
65                    If IndexSpeed <= 30 And Right(vNama, 1) = "s" Then
66                        vNama = Left(vNama, Len(vNama) - 1)
67                        ActiveBar1.Bands("Band").Tools("Speed" + Trim(str(IndexSpeed))).SetPicture _
                              0, _
                              ActiveBar1.Bands(vBand).Tools(vNama).GetPicture(0), _
                              &HC0C0C0

68                        ActiveBar1.Bands("Band").Tools("Speed" + Trim(str(IndexSpeed))).ToolTipText = _
                              ActiveBar1.Bands(vBand).Tools(vNama).ToolTipText

69                        ActiveBar1.Bands("Band").Tools("Speed" + Trim(str(IndexSpeed))).Name = _
                              ActiveBar1.Bands(vBand).Tools(vNama).Name

70                        ActiveBar1.Bands("Band").Tools(vNama).SubBand = _
                              ActiveBar1.Bands(vBand).Tools(vNama).SubBand

71                        IndexSpeed = IndexSpeed + 1
72                    Else
73                        If Right(vNama, 1) = "s" Then
74                            vNama = Left(vNama, Len(vNama) - 1)
75                        End If
76                    End If
                      
77                    ActiveBar1.Bands(vBand).Tools(vNama).Visible = True
78                    oRs.MoveNext
79                Loop
80            End If
81            oRs.Close
82            ActiveBar1.RecalcLayout
              
83            If gUstat = 0 Then
84                    Me.Picture = ImageList1.ListImages(1).Picture
85            Else
86                    Me.Picture = ImageList1.ListImages(1).Picture
87            End If
88        Else
89            Me.Picture = LoadPicture("")
90        End If

91        Me.Caption = gPerush + ", Nasmoco Integrated System (NIS) => Service System ( " & gVersi & " )"
92        vNmDrive = DriveSerialNumber("c:")
          
93        Call Timer1_Timer



94        If gUstat <> 1 Then
          'and kcb and adh and
95            If Trim(UCase(gTipeUser)) <> "SA" And Trim(UCase(gTipeUser)) <> "KABENG" And Trim(UCase(gTipeUser)) <> "KACAP" And Trim(UCase(gTipeUser)) <> "ADH" And Trim(UCase(gTipeUser)) <> "ADMIN" And Trim(UCase(gTipeUser)) <> "BILL" And Trim(UCase(gTipeUser)) <> "BILLING" And Trim(UCase(gTipeUser)) <> "MRS" Then
96                If Log.RPLHistori = 0 Then Call Do_Transfer_To_HO_Run 'If Log.RPLHistori = 1 Then Call Do_Transfer_To_HO_Run
97            End If
98        End If
End Sub

Private Sub isiStatusBar()
1         On Error Resume Next
2         If Log.RPLHistori = 0 Then
3             StatusBar1.Panels(1).Text = Log.LogHOConn & " IP : " & Trim(gServerDb) + " - HO : " & objO.FetchData_String("Select Count(*) from LogHO WITH (NOLOCK) Where LogHO = 0")
4         Else
5             If gTipeUser = "SA" And Log.newCetak = 1 Then
6                 StatusBar1.Panels(1).Text = "IP : " & Trim(gServerDb) '+ " - Belum Invoice : " & objO.FetchData_String("Select Count(*) from jobdta WITH (NOLOCK) Where deliv <> 'btl' and noinv = '' and front = '" & gSA & "'")
7             Else
8                 StatusBar1.Panels(1).Text = "IP : " & Trim(gServerDb)
9             End If
10        End If
11        StatusBar1.Panels(2).Text = Format(gTglTampung, "Long Date")
12        StatusBar1.Panels(3).Text = Format(gTglTampung, "HH:NN")
          'StatusBar1.Panels(5).Text = "  Unit Entry = " + Format(isiUnit, "###,0") + " "
13        StatusBar1.Panels(11).Text = " Login: " & gSA
14        If Log.OTD = 1 Then
15          StatusBar1.Panels(9).Text = "Monitoring Unit OTD"
16        End If
17        Select Case gTipeUser
              Case "SA"
18                nUnit = isiCustFU(gSA)
19                StatusBar1.Panels(10).Text = "Cust. Follow Up : " + Format(nUnit, "###,0") + ""
20                If Log.newCetak = 1 Then
21                    StatusBar1.Panels(8).Text = "Today Wait. Inv. : " + Format(support.countUnitWaitInvoiceBySA(1, gSA), "###,0") + " "
22                End If
23            Case "FOREMAN"
24                StatusBar1.Panels(6).Text = "Today Unit Waiting : " + Format(isiUnitWait, "###,0") + " "
25                StatusBar1.Panels(8).Text = "Today On.Job.Stop.: " + Format(isiUnitProgStop, "###,0") + " "
26            Case "BILLING"
27                statusKoneksi = cariPendingTrans
28                StatusBar1.Panels(8).Text = "Today Wait. Inv. : " + Format(isiUnitWaitInv, "###,0") + " "
29            Case "KASIR"
30                StatusBar1.Panels(8).Text = "Today Wait. Clos.: " + Format(isiUnitWaitCls, "###,0") + " "
31            Case "MRS"
      '            'penambahan booking online HO
      '            Set rsHlp = objO.FetchDataHO("select count(*) from bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
      '            Dim bOnline As Integer
      '            bOnline = 0
      '            If rsHlp.RecordCount > 0 Then bOnline = rsHlp.Fields(0).value
      '            StatusBar1.Panels(6).Text = "BookingOnline: " & bOnline
32        End Select
          'penambahan booking online HO
33        If Log.RPLHistori = 1 Then
34            Set rsHlp = objO.FetchData("select count(*) from [" & dbReplikasiHO & "]..bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
35        Else
              '' ==============================================================
              '' 20180523  revisi penambahan pengecekan flag koneksi HO
              '' ==============================================================
36            If Log.LogHOConn = 1 Then
37                Set rsHlp = objO.FetchDataHO("select count(*) from bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
38            Else
39                Set rsHlp = objO.FetchData("Select 0 as nilai")
40            End If
              
41        End If
          
          Dim bOnline As Integer
42        bOnline = 0
43        If rsHlp.RecordCount > 0 Then bOnline = rsHlp.Fields(0).value
44        StatusBar1.Panels(6).Text = "BookingOnline: " & bOnline
45        StatusBar1.Panels(7).Text = "Lock.AR : " & objAR.FetchDataLockAR_ & ", Overdue H-3 : " & objAR.FetchDataCust_OnAR_Overdue_(3)
End Sub

Private Function cariPendingTrans() As String
          Dim rsTrans As ADODB.Recordset
1        On Error GoTo cariPendingTrans_Error

2         cariPendingTrans = ""
3         Set rsTrans = objO.FetchData("SELECT COUNT(*) AS JML FROM TAMPHISTORY WITH (NOLOCK)")
4         If Not rsTrans.EOF Then
5             cariPendingTrans = "ada " & CDbl(rsTrans("JML")) & " data yang belum ke- transfer ke HO ......  Klik disini jika koneksi ke HO berhasil."
6         End If

7        On Error GoTo 0
8        Exit Function

cariPendingTrans_Error:

9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cariPendingTrans of Form Layar")
End Function

Private Function isiCustFU(sa As String) As Integer
1        On Error GoTo isiCustFU_Error

2         isiCustFU = objO.FetchData_Double("SELECT COUNT(DISTINCT a.NO_WO) AS JUMLAH FROM JOBDTA A WITH (NOLOCK) " _
                                     & "INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO " _
                                     & "WHERE TGLFU <= CONVERT(VarChar(15), GETDATE(), 1) AND RPR_COND <> '' AND DELIV <> 'BTL' AND FRONT = '" & Trim(sa) & "' AND YEAR(TGLFU) > 1900 " _
                                     & "AND A.NO_WO NOT IN (SELECT DISTINCT NO_WO FROM DTLFU WHERE NO_CHASSIS = A.NO_CHASSIS AND NO_WO = A.NO_WO AND (SAMBUNG = 'Y' OR COUNTER = 3))  AND A.FS_LOG_DEL = ''")

3        On Error GoTo 0
4        Exit Function

isiCustFU_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiCustFU of Form Layar")
End Function

Private Function isiUnit() As Integer
1        On Error GoTo isiUnit_Error

2         isiUnit = objO.FetchData_Double("SELECT count(no_wo) as jumlah FROM jobdta WITH (NOLOCK) WHERE date = CONVERT(varchar(10),GETDATE(),1) AND DELIV <> 'BTL'  AND FS_LOG_DEL = ''")

3        On Error GoTo 0
4        Exit Function

isiUnit_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnit of Form Layar")
End Function

Private Function isiUnitInv() As Integer
1        On Error GoTo isiUnitInv_Error

2         isiUnitInv = support.countUnitInvoice

3        On Error GoTo 0
4        Exit Function

isiUnitInv_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitInv of Form Layar")
End Function

Private Function isiUnitWaitInvAll() As Integer
          ' FUNGSI UNTUK MENCARI JUMLAH UNIT ORDER YANG MENUNGGU UNTUK DI INVOICE
1        On Error GoTo isiUnitWaitInvAll_Error

2             isiUnitWaitInvAll = support.countUnitWaitInvoice(0)

3        On Error GoTo 0
4        Exit Function

isiUnitWaitInvAll_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitWaitInvAll of Form Layar")
End Function

Private Function isiUnitWaitInv() As Integer
          ' FUNGSI UNTUK MENCARI JUMLAH UNIT ORDER YANG MENUNGGU UNTUK DI INVOICE PER HARI AKTIF
1        On Error GoTo isiUnitWaitInv_Error

2             isiUnitWaitInv = support.countUnitWaitInvoice(1)

3        On Error GoTo 0
4        Exit Function

isiUnitWaitInv_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitWaitInv of Form Layar")
End Function

Private Function isiUnitWaitClsAll() As Integer
          ' FUNGSI UNTUK MENCARI JUMLAH UNIT ORDER YANG MENUNGGU UNTUK DI registrasi
1        On Error GoTo isiUnitWaitClsAll_Error

2             isiUnitWaitClsAll = support.countUnitWaitClosing(0)

3        On Error GoTo 0
4        Exit Function

isiUnitWaitClsAll_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitWaitClsAll of Form Layar")
End Function

Private Function isiUnitWaitCls() As Integer
          ' FUNGSI UNTUK MENCARI JUMLAH UNIT ORDER YANG MENUNGGU UNTUK DI registrasi
1        On Error GoTo isiUnitWaitCls_Error

2             isiUnitWaitCls = support.countUnitWaitClosing(1)

3        On Error GoTo 0
4        Exit Function

isiUnitWaitCls_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitWaitCls of Form Layar")
End Function

Private Function isiUnitProg() As Integer
          ' FUNGSI UNTUK MENCARI JUMLAH UNIT ORDER YG SEDANG DIKERJAKAN
1        On Error GoTo isiUnitProg_Error

2             isiUnitProg = support.countUnitOnProcess

3        On Error GoTo 0
4        Exit Function

isiUnitProg_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitProg of Form Layar")
End Function

Private Function isiUnitWait() As Integer
          ' FUNGSI UNTUK MENCARI JUMLAH UNIT ORDER YG MENUNGGU UNTUK DIKERJAKAN
1        On Error GoTo isiUnitWait_Error

2             isiUnitWait = support.countUnitWaitProcess

3        On Error GoTo 0
4        Exit Function

isiUnitWait_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitWait of Form Layar")
End Function

Private Function isiUnitProgStop() As Integer
          ' FUNGSI UNTUK MENCARI JUMLAH UNIT ORDER YG SEDANG STOPPED
1        On Error GoTo isiUnitProgStop_Error

2             isiUnitProgStop = support.countUnitOnStopped(0)

3        On Error GoTo 0
4        Exit Function

isiUnitProgStop_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitProgStop of Form Layar")
End Function

Private Function isiUnitProgStopAll() As Integer
          ' FUNGSI UNTUK MENCARI JUMLAH UNIT ORDER YG STOPPED SMP DENGAN HARI INI
1        On Error GoTo isiUnitProgStopAll_Error

2             isiUnitProgStopAll = support.countUnitOnStopped(1)

3        On Error GoTo 0
4        Exit Function

isiUnitProgStopAll_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitProgStopAll of Form Layar")
End Function

Private Function isiUnitBook() As Integer
      Dim rsSearch As ADODB.Recordset
          ' FUNGSI UNTUK MENCARI JUMLAH UNIT BOOKING PER HARI AKTIF
           
1        On Error GoTo isiUnitBook_Error

2         Set rsSearch = objO.FetchData("SELECT * FROM Book WITH (NOLOCK) WHERE date = CONVERT(varchar(10),GETDATE(),1) AND deliv <> 'BTL'  AND FS_LOG_DEL = ''")
3         If Not rsSearch.EOF Then
4             isiUnitBook = CDbl(rsSearch.RecordCount)
5         End If

6        On Error GoTo 0
7        Exit Function

isiUnitBook_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiUnitBook of Form Layar")
           
End Function

Private Function isiTomBook() As Integer
1        On Error GoTo isiTomBook_Error

2         isiTomBook = objO.FetchData_Count("SELECT NO_Polisi FROM Book WITH (NOLOCK) WHERE date = CONVERT(varchar(10),GETDATE(),1) AND deliv <> 'BTL' AND FS_LOG_DEL = ''")

3        On Error GoTo 0
4        Exit Function

isiTomBook_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiTomBook of Form Layar")
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Anda yakin mau keluar dari program ini ?", vbQuestion + vbYesNo, "Close Program") = vbYes Then
        On Error Resume Next
        Kill Left(App.Path, 1) & ":\*.tmp"
        Call saveLogOutUser(gSA, ComptName)
        Call closeAllConnection
        End
    End If
    Cancel = 1
End Sub

Private Sub cekStatusProgress()
      Dim pesan As String
1        On Error GoTo cekStatusProgress_Error

2         pesan = ""
3         Set rsHlp = objO.FetchData("SELECT LEFT(KETERANGAN, CHARINDEX('|', KETERANGAN)-1) AS NOPOL, * FROM STALL WITH (NOLOCK) WHERE FLAG = 'S' AND TANGGAL = CONVERT(varchar(10),GETDATE(),1) ORDER BY CAST(SUBSTRING(NOSTALL,7,2) AS INTEGER)")
4         If Not rsHlp.EOF Then
5             Do While Not rsHlp.EOF
6                 pesan = pesan & " " & rsHlp("NOSTALL") & " untuk kendaraan No. Polisi : " & rsHlp("NOPOL") & " belum di pause. " & vbCr
7                 rsHlp.MoveNext
8             Loop
9             MsgBox pesan, vbInformation, "STALL BELUM DIPAUSE"
10        End If
11        rsHlp.Close

12       On Error GoTo 0
13       Exit Sub

cekStatusProgress_Error:

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekStatusProgress of Form Layar")
End Sub

Private Sub cekFileExport()
      Dim rsHlp2 As ADODB.Recordset


          ' HO EXPORT FILE UPDATE KE CABANG

1        On Error GoTo cekFileExport_Error

2     If Log.LogHOConn = 1 Then
3         If ConHO = True Then
4             Set rsHlp = objO.FetchDataHO("SELECT * FROM TEMPKIRIM WITH (NOLOCK)")
5             If Not rsHlp.EOF Then
6                 Do While Not rsHlp.EOF
7                     Set rsHlp2 = objO.FetchDataHO("SELECT * FROM T_DEALER WITH (NOLOCK) WHERE INTERNAL = 1 AND VPNCABANG <> '' and FS_LOG_DEL = '' ORDER BY FS_KD_DEALER")
8                     If Not rsHlp2.EOF Then
9                         Do While Not rsHlp2.EOF
10                            If Trim(rsHlp2("DBASECBG")) <> "" Then
                                  
11                                If Trim(rsHlp2("NEWENKRIP")) = "1" Then
12                                    If Get_ConnectionAcces(rsHlp2("DBASECBG"), Dekripsi(rsHlp2("USERCBG"), KataKunci(3)), Dekripsi(rsHlp2("PASSCBG"), KataKunci(5)), rsHlp2("VPNCABANG")) Then
13                                        Call objO.InsertDataTempKirim("INSERT INTO TEMPKIRIM SELECT DISTINCT '" & Trim(rsHlp("TIPE")) & "', '" & Trim(rsHlp("KODE")) & "', '" & Trim(rsHlp("STATUS")) & "' WHERE NOT EXISTS (SELECT * FROM TEMPKIRIM WITH (NOLOCK) WHERE TIPE = '" & Trim(rsHlp("TIPE")) & "' AND KODE = '" & Trim(rsHlp("KODE")) & "')", "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Dekripsi(rsHlp2("USERCBG"), KataKunci(3)) & ";password=" & Dekripsi(rsHlp2("PASSCBG"), KataKunci(5)) & ";Initial Catalog=" & rsHlp2("DBASECBG") & ";server=" & rsHlp2("VPNCABANG"))
14                                    Else
15                                        Call objO.InsertDataHO("INSERT INTO TEMPCBGGAGAL SELECT DISTINCT '" & rsHlp2("FS_KD_DEALER") & "','" & rsHlp("TIPE") & "', '" & rsHlp("KODE") & "', '" & rsHlp("STATUS") & "' WHERE NOT EXISTS (SELECT * FROM TEMPCBGGAGAL WITH (NOLOCK) WHERE FS_KD_DEALER = '" & rsHlp2("FS_KD_DEALER") & "' AND TIPE = '" & rsHlp("TIPE") & "' AND KODE = '" & rsHlp("KODE") & "')")
16                                    End If
17                                Else
18                                    If Get_ConnectionAcces(rsHlp2("DBASECBG"), Trim(rsHlp2("USERCBG")), Trim(rsHlp2("PASSCBG")), rsHlp2("VPNCABANG")) Then
19                                        Call objO.InsertDataTempKirim("INSERT INTO TEMPKIRIM SELECT DISTINCT '" & Trim(rsHlp("TIPE")) & "', '" & Trim(rsHlp("KODE")) & "', '" & Trim(rsHlp("STATUS")) & "' WHERE NOT EXISTS (SELECT * FROM TEMPKIRIM WITH (NOLOCK) WHERE TIPE = '" & Trim(rsHlp("TIPE")) & "' AND KODE = '" & Trim(rsHlp("KODE")) & "')", "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Trim(rsHlp2("USERCBG")) & ";password=" & Trim(rsHlp2("PASSCBG")) & ";Initial Catalog=" & rsHlp2("DBASECBG") & ";server=" & rsHlp2("VPNCABANG"))
20                                    Else
21                                        Call objO.InsertDataHO("INSERT INTO TEMPCBGGAGAL SELECT DISTINCT '" & rsHlp2("FS_KD_DEALER") & "','" & rsHlp("TIPE") & "', '" & rsHlp("KODE") & "', '" & rsHlp("STATUS") & "' WHERE NOT EXISTS (SELECT * FROM TEMPCBGGAGAL WITH (NOLOCK) WHERE FS_KD_DEALER = '" & rsHlp2("FS_KD_DEALER") & "' AND TIPE = '" & rsHlp("TIPE") & "' AND KODE = '" & rsHlp("KODE") & "')")
22                                    End If
23                                End If
24                            End If
25                            rsHlp2.MoveNext
26                        Loop
27                    End If
28                    rsHlp.MoveNext
29                Loop
30                Call objO.DeleteDataHO("DELETE TEMPKIRIM")
31            End If
32        End If
33    End If

34       On Error GoTo 0
35       Exit Sub

cekFileExport_Error:

36        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekFileExport of Form Layar")
End Sub

Private Sub syncHistory()
      Dim rsHlp2 As ADODB.Recordset
      Dim rsDealer As ADODB.Recordset
      Dim rsHO As ADODB.Recordset
      Dim memoSublet As String
      Dim memoPart As String
      Dim memoAddJob As String

1        On Error GoTo syncHistory_Error

2         Set rsHlp2 = objO.FetchDataHO("SELECT * FROM T_DEALER WITH (NOLOCK) WHERE INTERNAL = 1 AND FS_KD_DEALER = 'NR 14' AND VPNCABANG <> '' and FS_LOG_DEL = '' ORDER BY FS_KD_DEALER")
3         If Not rsHlp2.EOF Then
4             Do While Not rsHlp2.EOF
5                 If Trim(rsHlp2("DBASECBG")) <> "" Then
6                     If Get_Connection(rsHlp2("DBASECBG"), rsHlp2("USERCBG"), rsHlp2("PASSCBG"), rsHlp2("VPNCABANG")) Then
7                         Set rsDealer = objO.FetchData("SELECT A.CUSTOMER,A.CUST_ADRES,A.KOTA,A.NO_POLISI,A.NO_ENGINE,A.NO_CHASSIS,A.FRONT,A.MECH,A.KM,A.DATE,A.PART_G,A.PART_NG, " _
                                                     & "A.OIL_G,A.OIL_NG,A.MTRL,A.LABOR,A.BODY,A.SBLT,A.MET,A.TOTAL,A.JOB,A.REP1,A.REP2,A.REP3,A.REP4,A.REP5,A.REP6,A.REP7, " _
                                                     & "A.MEKANIK1,A.MEKANIK2,A.MEKANIK3,A.MEKANIK4,A.MEKANIK5,A.MEKANIK6,A.MEKANIK7,A.MEKANIK8,A.KLPKERJA1,A.KLPKERJA2,A.KLPKERJA3, " _
                                                     & "A.KLPKERJA4,A.KLPKERJA5,A.KLPKERJA6,A.KLPKERJA7,A.KLPKERJA8,B.RTE1,B.RTE2,B.RTE3,B.RTE4,B.RTE5,B.RTE6,B.RTE7,B.RTE8,A.NO_WO,A.NOINV, " _
                                                     & "A.NOFAK,A.STATUS,A.PAID,A.DISC,A.DISCT,A.REMARK1,A.FN_RETURN,A.FS_RC,A.FS_KD_DEALER,A.FS_NM_DEALER,B.TIME,B.TIME_M, " _
                                                     & "B.TIME_S , B.TIME_A, TIME_AT, TIME_SA, TIPEWO, a.KODETIPE, a.SBE, a.FS_KD_PLG, a.FS_KD_WIL, a.THS, a.MOU, a.PDS, ALASAN_DEL, B.INVDATE " _
                                                     & "FROM HISTORY A WITH (NOLOCK) INNER JOIN T_INVOICE B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NOINV = B.NOINV " _
                                                     & "INNER JOIN JOBDTA C WITH (NOLOCK) ON A.NO_WO = C.NO_WO AND A.NOINV = C.NOINV WHERE MONTH(B.INVDATE) = 6 AND YEAR(B.INVDATE) = 2010 AND A.NOINV <> '' AND A.FS_LOG_DEL = '' ORDER BY A.NO_WO")
8                         If Not rsDealer.EOF Then
9                             n = rsDealer.RecordCount
10                            Do While Not rsDealer.EOF

11                                n = n - 1
12                                Set rsHO = objO.FetchDataHO("SELECT * FROM HISTORYnew WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsDealer("NO_WO")) & "' AND NOINV = '" & Trim(rsDealer("NOINV")) & "'")
13                                If rsHO.EOF Then
                                      ' MEMOSUBLET
14                                    memoSublet = ""
15                                    Set rsHlp = objO.FetchData("SELECT * FROM HIST_SUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsDealer("NO_WO")) & "'")
16                                    If Not rsHlp.EOF Then
17                                        Do While Not rsHlp.EOF
18                                            memoSublet = memoSublet + (rsHlp("NOPO") & "^" & rsHlp("NO_WO") & "^" & rsHlp("NOINV") & "^" & rsHlp("NO_CHASSIS") & "^" & rsHlp("NOFAK") & "^" & rsHlp("KODEITEM") & "^" & rsHlp("NAMABARANG") & "^" & rsHlp("KODESATUAN") & "^" & rsHlp("QTY") & "^" & rsHlp("HARGA") & "^" & rsHlp("NONOTA") & "^" & rsHlp("KETERANGAN") & "^" & rsHlp("FLAG") & "^" & rsHlp("SUPPLIER") & "^" & rsHlp("TAGIH") & "|")
19                                            rsHlp.MoveNext
20                                        Loop
21                                    End If

                                      ' MEMOPART
22                                    memoPart = ""
23                                    Set rsHlp = objO.FetchData("SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsDealer("NO_WO")) & "' and FS_LOG_DEL = ''")
24                                    If Not rsHlp.EOF Then
25                                        Do While Not rsHlp.EOF
26                                            memoPart = memoPart + (rsHlp("NO_WO") & "^" & rsHlp("KODEITEM") & "^" & rsHlp("KETERANGAN") & "^" & rsHlp("QTY") & "^" & rsHlp("HARGA") & "^" & rsHlp("OLI") & "^" & rsHlp("NOINV") & "^" & rsHlp("FN_EXT_DISKON")) & "|"
27                                            rsHlp.MoveNext
28                                        Loop
29                                    End If

                                      ' MEMOADDJOB
30                                    memoAddJob = ""
31                                    Set rsHlp = objO.FetchData("SELECT * FROM HIST_JOBADD WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsDealer("NO_WO")) & "'")
32                                    If Not rsHlp.EOF Then
33                                        Do While Not rsHlp.EOF
34                                            memoAddJob = memoAddJob + (rsHlp("NO_WO") & "^" & rsHlp("NO_CHASSIS") & "^" & Format(rsHlp("DATE"), "MM/DD/YYYY") & "^" & rsHlp("KLPKERJA") & "^" & rsHlp("REP") & "^" & CDbl(rsHlp("RTE")) & "^" & rsHlp("ACTRTE") & "^" & rsHlp("MEKANIK") & "^" & rsHlp("GRP") & "^" & rsHlp("OFFN") & "^" & rsHlp("DISC") & "^" & rsHlp("TIME") & "|")
35                                            rsHlp.MoveNext
36                                        Loop
37                                    End If

                                      ' INSERT HISTORY
38                                    Call objO.InsertDataHO("INSERT INTO HISTORY VALUES ('" & Trim(rsDealer("CUSTOMER")) & "', '" & Trim(rsDealer("CUST_ADRES")) & "', '" & Trim(rsDealer("KOTA")) & "', '" & Trim(rsDealer("NO_POLISI")) & "', '" & Trim(rsDealer("NO_ENGINE")) & "', '" & Trim(rsDealer("NO_CHASSIS")) & "', " _
                                                     & "'" & Trim(rsDealer("FRONT")) & "', '" & Trim(rsDealer("MECH")) & "', '" & Trim(rsDealer("KM")) & "', '" & Format(rsDealer("DATE"), "MM/DD/YYYY") & "', " & Trim(rsDealer("PART_G")) & ", " & Trim(rsDealer("PART_NG")) & ", " & Trim(rsDealer("OIL_G")) & ", " & Trim(rsDealer("OIL_NG")) & " , " _
                                                     & "" & Trim(rsDealer("MTRL")) & ", " & Trim(rsDealer("LABOR")) & ", " & Trim(rsDealer("BODY")) & ", " & Trim(rsDealer("SBLT")) & ", " & Trim(rsDealer("MET")) & ", " & Trim(rsDealer("TOTAL")) & ", '" & Trim(rsDealer("JOB")) & "', '" & Trim(rsDealer("REP1")) & "', " _
                                                     & "'" & Trim(rsDealer("REP2")) & "', '" & Trim(rsDealer("REP3")) & "', '" & Trim(rsDealer("REP4")) & "', '" & Trim(rsDealer("REP5")) & "', '" & Trim(rsDealer("REP6")) & "', '" & Trim(rsDealer("REP7")) & "', '" & Trim(rsDealer("MEKANIK1")) & "', '" & Trim(rsDealer("MEKANIK2")) & "', " _
                                                     & "'" & Trim(rsDealer("MEKANIK3")) & "', '" & Trim(rsDealer("MEKANIK4")) & "', '" & Trim(rsDealer("MEKANIK5")) & "', '" & Trim(rsDealer("MEKANIK6")) & "', '" & Trim(rsDealer("MEKANIK7")) & "', '" & Trim(rsDealer("MEKANIK8")) & "', '" & Trim(rsDealer("KLPKERJA1")) & "', '" & Trim(rsDealer("KLPKERJA2")) & "', " _
                                                     & "'" & Trim(rsDealer("KLPKERJA3")) & "', '" & Trim(rsDealer("KLPKERJA4")) & "', '" & Trim(rsDealer("KLPKERJA5")) & "', '" & Trim(rsDealer("KLPKERJA6")) & "', '" & Trim(rsDealer("KLPKERJA7")) & "', '" & Trim(rsDealer("KLPKERJA8")) & "', '" & CDbl(Replace(rsDealer("RTE1"), ",", ".")) & "', '" & CDbl(Replace(rsDealer("RTE2"), ",", ".")) & "', " _
                                                     & "'" & CDbl(Replace(rsDealer("RTE3"), ",", ".")) & "', '" & CDbl(Replace(rsDealer("RTE4"), ",", ".")) & "', '" & CDbl(Replace(rsDealer("RTE5"), ",", ".")) & "', '" & CDbl(Replace(rsDealer("RTE6"), ",", ".")) & "', '" & CDbl(Replace(rsDealer("RTE7"), ",", ".")) & "', '" & CDbl(Replace(rsDealer("RTE8"), ",", ".")) & "', '" & Trim(rsDealer("NO_WO")) & "', '" & Trim(rsDealer("NOINV")) & "', " _
                                                     & "'" & Trim(rsDealer("NOFAK")) & "', '" & Trim(rsDealer("STATUS")) & "', '" & Trim(rsDealer("PAID")) & "', '" & CDbl(Replace(rsDealer("DISC"), ",", ".")) & "', '" & CDbl(Replace(rsDealer("DISCT"), ",", ".")) & "', '" & Trim(rsDealer("REMARK1")) & "', '" & CDbl(rsDealer("FN_RETURN")) & "', '" & Trim(rsDealer("FS_RC")) & "', " _
                                                     & "'" & Trim(rsDealer("FS_KD_DEALER")) & "', '" & Trim(rsDealer("FS_NM_DEALER")) & "', '" & Trim(rsDealer("TIME")) & "', '" & Trim(rsDealer("TIME_M")) & "', '" & Trim(rsDealer("TIME_S")) & "', '" & Trim(rsDealer("TIME_A")) & "', '" & Trim(rsDealer("TIME_AT")) & "', '" & Trim(rsDealer("TIME_SA")) & "', " _
                                                     & "'" & Trim(rsDealer("TIPEWO")) & "', '" & Trim(rsDealer("KODETIPE")) & "', '" & Trim(rsDealer("SBE")) & "', '" & Trim(rsDealer("FS_KD_PLG")) & "', '" & Trim(rsDealer("FS_KD_WIL")) & "', '" & Trim(rsDealer("THS")) & "', '" & Trim(rsDealer("MOU")) & "', '" & CDbl(rsDealer("PDS")) & "', " _
                                                     & "'" & Trim(memoPart) & "', '" & Trim(memoSublet) & "', '', '', '', '" & Trim(memoAddJob) & "','" & Format(rsDealer("INVDATE"), "MM/DD/YYYY") & "')")
39                                Else
40                                    Call objO.UpdateDataHO("UPDATE HISTORY SET INVDATE = '" & Format(rsDealer("INVDATE"), "MM/DD/YYYY") & "' WHERE NO_WO = '" & Trim(rsDealer("NO_WO")) & "' AND NOINV = '" & Trim(rsDealer("NOINV")) & "'")
41                                End If
42                                rsDealer.MoveNext
43                            Loop
44                        End If
45                    End If
46                End If
47                rsHlp2.MoveNext
48            Loop
49        End If
50        rsHlp2.Close

51       On Error GoTo 0
52       Exit Sub

syncHistory_Error:

53        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure syncHistory of Form Layar")
End Sub

Private Sub cekFileImport()
      Dim rsHlp2 As ADODB.Recordset
      Dim rsHlp3 As ADODB.Recordset
      Dim rsHlp4 As ADODB.Recordset
          
          ' CABANG CEK FILE HASIL IMPORT HO
1        On Error GoTo cekFileImport_Error

2         Set rsHlp = objO.FetchData("SELECT * FROM TEMPKIRIM WITH (NOLOCK) ORDER BY STATUS")
          
3         If Not rsHlp.EOF Then
4             Do While Not rsHlp.EOF
5                 If Log.LogHOConn = 1 Then
6                 If ConHO = True Then
7                     Select Case Trim(rsHlp("TIPE"))
                          Case "MODEL"
8                             If Trim(rsHlp("STATUS")) = "S" Then
                                   
9                                 Set rsHlp2 = objO.FetchDataHO("SELECT * FROM MODEL WITH (NOLOCK) WHERE MDL = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
10                                If Not rsHlp2.EOF Then
                                       
11                                    Set rsHlp3 = objO.FetchData("SELECT * FROM MODEL WITH (NOLOCK) WHERE MDL = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
12                                    If Not rsHlp3.EOF Then
13                                        Call objO.UpdateData("UPDATE MODEL SET MODEL = '" & rsHlp2("MODEL") & "', [DESC] = '" & rsHlp2("DESC") & "', FS_GRP_MDL = '" & rsHlp2("FS_GRP_MDL") & "', NO_CHASSIS = '" & rsHlp2("NO_CHASSIS") & "', ISICABANG = 0 WHERE MDL = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
14                                    Else
15                                        Call objO.InsertData("INSERT INTO MODEL VALUES ('" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("DESC") & "', '" & rsHlp2("FS_GRP_MDL") & "', '" & rsHlp2("NO_CHASSIS") & "','','','0','0')")
16                                    End If
17                                    rsHlp3.Close
18                                End If
19                                rsHlp2.Close
20                            ElseIf Trim(rsHlp("STATUS")) = "D" Then
21                                Call objO.DeleteData("DELETE MODEL WHERE MDL = '" & rsHlp("KODE") & "'")
22                            End If
23                        Case "M_SUBLET"
24                            If Trim(rsHlp("STATUS")) = "S" Then
                                  
25                                Set rsHlp2 = objO.FetchDataHO("SELECT * FROM MASTER_SUBLET WITH (NOLOCK) WHERE ID = '" & Trim(rsHlp!kode) & "'")
26                                If Not rsHlp2.EOF Then
                                       
27                                    Set rsHlp3 = objO.FetchData("SELECT * FROM MASTER_SUBLET WITH (NOLOCK) WHERE id = '" & rsHlp("KODE") & "'")
28                                    If Not rsHlp3.EOF Then
29                                        Call objO.UpdateData("UPDATE MASTER_SUBLET SET [TYPE] = '" & rsHlp2!Type & "',[DESC] = '" & rsHlp2!desc & "',HARGABELI = " & Replace(rsHlp2!hargaBeli, ",", ".") & " ,HARGAJUAL = " & Replace(rsHlp2!hargaJual, ",", ".") & " ,UPDATEBY = '" & rsHlp2!updateby & "',TGLUPDATE = GETDATE() WHERE ID = '" & rsHlp!kode & "'")
30                                    Else
31                                        Call objO.InsertData("INSERT INTO [MASTER_SUBLET]([TYPE],[DESC],[HARGABELI],[HARGAJUAL],[CREATEBY],[TGLCREATE],[UPDATEBY],[TGLUPDATE],[ISHAPUS]) VALUES ('" & rsHlp2!Type & "','" & rsHlp2!desc & "'," & rsHlp2!hargaBeli & "," & rsHlp2!hargaJual & ",'" & gSA & "',getdate(),'','','0')")
32                                    End If
33                                    rsHlp3.Close
34                                End If
35                                rsHlp2.Close
36                            End If
37                        Case "M_T1"
38                            If Trim(rsHlp("STATUS")) = "S" Then
39                                Set rsHlp2 = objO.FetchDataHO("SELECT * FROM MASTER_T1 WITH (NOLOCK) WHERE ID = '" & Trim(rsHlp!kode) & "' AND ISHAPUS = '0'")
40                                If Not rsHlp2.EOF Then
                                       
41                                    Set rsHlp3 = objO.FetchData("SELECT * FROM MASTER_T1 WITH (NOLOCK) WHERE id = '" & rsHlp("KODE") & "' AND ISHAPUS = '0'")
42                                    If Not rsHlp3.EOF Then
43                                        Call objO.UpdateData("UPDATE MASTER_T1 SET [TERM] = '" & rsHlp2!TERM & "',[DESC] = '" & rsHlp2!desc & "',TGLUPDATE = GETDATE() WHERE ID = '" & rsHlp!kode & "'")
44                                    Else
45                                        Call objO.InsertData("INSERT INTO [MASTER_T1] VALUES ('" & rsHlp2!id & "','" & rsHlp2!TERM & "'," & rsHlp2!desc & ",GETDATE(),'','0')")
46                                    End If
47                                    rsHlp3.Close
48                                End If
49                                rsHlp2.Close
50                            ElseIf Trim(rsHlp!Status) = "D" Then
51                                Set rsHlp2 = objO.FetchDataHO("SELECT * FROM MASTER_T1 WITH (NOLOCK) WHERE ID = '" & Trim(rsHlp!kode) & "' AND ISHAPUS = '1'")
52                                If Not rsHlp2.EOF Then
53                                    Call objO.UpdateData("UPDATE MASTER_T1 SET ISHAPUS = '1',TGLUPDATE = GETDATE() WHERE ID = '" & rsHlp!kode & "'")
54                                End If
55                            End If
56                        Case "M_T2"
57                            If Trim(rsHlp("STATUS")) = "S" Then
58                                Set rsHlp2 = objO.FetchDataHO("SELECT * FROM MASTER_T2 WITH (NOLOCK) WHERE ID = '" & Trim(rsHlp!kode) & "' AND ISHAPUS = '0'")
59                                If Not rsHlp2.EOF Then
                                       
60                                    Set rsHlp3 = objO.FetchData("SELECT * FROM MASTER_T2 WITH (NOLOCK) WHERE id = '" & rsHlp("KODE") & "' AND ISHAPUS = '0'")
61                                    If Not rsHlp3.EOF Then
62                                        Call objO.UpdateData("UPDATE MASTER_T2 SET [TERM] = '" & rsHlp2!TERM & "',[DESC1] = '" & rsHlp2!desc1 & "',[DESC2] = '" & rsHlp2!desc2 & "',TGLUPDATE = GETDATE() WHERE ID = '" & rsHlp!kode & "'")
63                                    Else
64                                        Call objO.InsertData("INSERT INTO [MASTER_T2] VALUES ('" & rsHlp2!id & "','" & rsHlp2!TERM & "','" & rsHlp2!desc1 & "','" & rsHlp2!desc2 & "',GETDATE(),'','0')")
65                                    End If
66                                    rsHlp3.Close
67                                End If
68                                rsHlp2.Close
69                            ElseIf Trim(rsHlp!Status) = "D" Then
70                                Set rsHlp2 = objO.FetchDataHO("SELECT * FROM MASTER_T2 WITH (NOLOCK) WHERE ID = '" & Trim(rsHlp!kode) & "' AND ISHAPUS = '1'")
71                                If Not rsHlp2.EOF Then
72                                    Call objO.UpdateData("UPDATE MASTER_T2 SET ISHAPUS = '1',TGLUPDATE = GETDATE() WHERE ID = '" & rsHlp!kode & "'")
73                                End If
74                            End If
75                        Case "M_FRATE"
76                            If Trim(rsHlp("STATUS")) = "S" Then
                                  Dim OPENO, VDS, Tahun As String
77                                OPENO = Left(rsHlp!kode, InStr(rsHlp!kode, "|") - 1)
78                                VDS = Mid(rsHlp!kode, InStr(rsHlp!kode, "|") + 1, InStrRev(rsHlp!kode, "|"))
79                                VDS = Left(VDS, InStr(rsHlp!kode, "|") - 1)
80                                Tahun = Mid(rsHlp!kode, InStrRev(rsHlp!kode, "|") + 1, Len(rsHlp!kode))
                                  
81                                Set rsHlp2 = objO.FetchDataHO("SELECT * FROM MASTER_FLATERATE WITH (NOLOCK) WHERE openo = '" & OPENO & "' AND VDS = '" & VDS & "' AND YEAR = '" & Tahun & "'")
82                                If Not rsHlp2.EOF Then
83                                    Set rsHlp3 = objO.FetchData("SELECT * FROM MASTER_FLATERATE WITH (NOLOCK) WHERE openo = '" & OPENO & "' AND VDS = '" & VDS & "' AND YEAR = '" & Tahun & "'")
84                                    If Not rsHlp3.EOF Then
85                                        Call objO.UpdateData("UPDATE [MASTER_FLATERATE] SET [OPENO] = '" & OPENO & "',[VDS] = '" & VDS & "',[YEAR] = '" & Tahun & "',[HOUR] = '" & rsHlp2!hour & "',[CLM_TYPE] = '" & rsHlp2!clm_type & "',[WMI] = '" & rsHlp2!wmi & "',[ID] = '" & rsHlp2!id & "',[FRNC] = '" & rsHlp2!frnc & "',[LANGUAGE] = '" & rsHlp2!Language & "',[DESC_A] = '" & rsHlp2!desc_a & "',[DESC_B] = '" & rsHlp2!desc_b & "',[rate] = " & Replace(rsHlp2!Rate, ",", ".") & " Where openo = '" & OPENO & "' AND VDS = '" & VDS & "' AND YEAR = '" & Tahun & "'")
86                                    Else
87                                        Call objO.InsertData("INSERT INTO [MASTER_FLATERATE]([OPENO],[VDS],[YEAR],[HOUR],[CLM_TYPE],[WMI],[ID],[FRNC],[LANGUAGE],[DESC_A],[DESC_B],[TGLENTRY],[rate]) VALUES ('" & Trim(rsHlp2!OPENO) & "','" & Trim(rsHlp2!VDS) & "','" & Trim(rsHlp2!year) & "','" & Trim(rsHlp2!hour) & "','" & Trim(rsHlp2!clm_type) & "','" & Trim(rsHlp2!wmi) & "','" & Trim(rsHlp2!id) & "','" & Trim(rsHlp2!frnc) & "','" & Trim(rsHlp2!Language) & "','" & Trim(rsHlp2!desc_a) & "','" & Trim(rsHlp2!desc_b) & "',getdate()," & Replace(rsHlp2!Rate, ",", ".") & ")")
88                                    End If
89                                    rsHlp3.Close
90                                End If
91                                rsHlp2.Close
92                            End If
93                        Case "PARTPACK"
      '                        If Trim(rsHlp("STATUS")) = "S" Then
                                   
94                                Set rsHlp2 = objO.FetchDataHO("SELECT * FROM HO_PARTPACK WITH (NOLOCK) WHERE FS_ITEM = '" & rsHlp("KODE") & "'")
95                                If Not rsHlp2.EOF Then
96                                    Do While Not rsHlp2.EOF
                                           
97                                        Set rsHlp3 = objO.FetchData("SELECT * FROM HO_PARTPACK WITH (NOLOCK) WHERE FS_ITEM = '" & rsHlp2("FS_ITEM") & "' AND MDL = '" & rsHlp2("MDL") & "' AND ITEM_PART = '" & rsHlp2("ITEM_PART") & "' AND FS_LOG_DEL = '" & Trim(rsHlp2("FS_LOG_DEL")) & "'")
98                                        If Not rsHlp3.EOF Then
99                                            Call objO.UpdateData("UPDATE HO_PARTPACK SET JOB = '" & rsHlp2("JOB") & "', MODEL = '" & rsHlp2("MODEL") & "', ITEM_PART = '" & rsHlp2("ITEM_PART") & "', NAMA_PART = '" & rsHlp2("NAMA_PART") & "', QTY = '" & CDbl(rsHlp2("QTY")) & "', KETERANGAN = '" & rsHlp2("KETERANGAN") & "', FS_LOG_DEL = '" & Trim(rsHlp2("FS_LOG_DEL")) & "', FD_TGL_DEL = '" & Format(rsHlp2("FD_TGL_DEL"), "MM/DD/YYYY HH:NN:SS") & "' WHERE FS_ITEM = '" & rsHlp2("FS_ITEM") & "' AND MDL = '" & rsHlp2("MDL") & "' AND FS_LOG_DEL = '" & Trim(rsHlp2("FS_LOG_DEL")) & "'")
100                                       Else
101                                           Call objO.InsertData("INSERT INTO HO_PARTPACK VALUES ('" & rsHlp2("JOB") & "','" & rsHlp2("FS_ITEM") & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("ITEM_PART") & "','" & rsHlp2("NAMA_PART") & "', '" & CDbl(rsHlp2("QTY")) & "', '" & rsHlp2("KETERANGAN") & "','" & Trim(rsHlp2("FS_LOG_DEL")) & "','" & Format(rsHlp2("FD_TGL_DEL"), "MM/DD/YYYY HH:NN:SS") & "')")
102                                       End If
103                                       rsHlp3.Close
104                                       rsHlp2.MoveNext
105                                   Loop
106                               End If
107                               rsHlp2.Close
108                       Case "GANTIPART"
109                           Select Case Trim(rsHlp("STATUS"))
                                  Case "S"
                                       
110                                   Set rsHlp2 = objO.FetchDataHO("SELECT * FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
111                                   If Not rsHlp2.EOF Then
112                                       Do While Not rsHlp2.EOF
                                               
113                                           Set rsHlp3 = objO.FetchData("SELECT * FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & rsHlp2("ITEM_PART") & "' AND FS_ITEM = '" & rsHlp2("FS_ITEM") & "' AND MDL = '" & rsHlp2("MDL") & "' AND FS_LOG_DEL = ''")
114                                           If Not rsHlp3.EOF Then
115                                               Call objO.UpdateData("UPDATE HO_PARTPACK SET JOB = '" & rsHlp2("JOB") & "', MODEL = '" & rsHlp2("MODEL") & "', NAMA_PART = '" & rsHlp2("NAMA_PART") & "', QTY = '" & CDbl(rsHlp2("QTY")) & "', KETERANGAN = '" & rsHlp2("KETERANGAN") & "', FS_LOG_DEL = '" & rsHlp2("FS_LOG_DEL") & "', FD_TGL_DEL = '" & rsHlp2("FD_TGL_DEL") & "' WHERE ITEM_PART = '" & rsHlp2("ITEM_PART") & "' AND FS_ITEM = '" & rsHlp2("FS_ITEM") & "' AND MDL = '" & rsHlp2("MDL") & "' AND FS_LOG_DEL = ''")
116                                           Else
117                                               Call objO.InsertData("INSERT INTO HO_PARTPACK VALUES ('" & rsHlp2("JOB") & "', '" & rsHlp2("FS_ITEM") & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("ITEM_PART") & "', '" & rsHlp2("NAMA_PART") & "', '" & CDbl(rsHlp2("QTY")) & "', '" & rsHlp2("KETERANGAN") & "', '" & rsHlp2("FS_LOG_DEL") & "', '" & rsHlp2("FD_TGL_DEL") & "')")
118                                           End If
119                                           rsHlp3.Close
120                                           rsHlp2.MoveNext
121                                       Loop
122                                   End If
123                               Case "D"
124                                   Call objO.DeleteData("DELETE HO_PARTPACK WHERE ITEM_PART = '" & rsHlp("KODE") & "'")
125                           End Select
126                       Case "FLATE"
127                           If Trim(rsHlp("STATUS")) = "S" Then
128                               Set rsHlp4 = objO.FetchData("Select * from Model WITH (NOLOCK) where left(MDL," & Len(Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|") - 1))) & ") = '" & Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|") - 1)) & "'")
129                               If rsHlp4.RecordCount > 0 Then
130                                   rsHlp4.MoveFirst
131                                   Do Until rsHlp4.EOF
132                                       Set rsHlp2 = objO.FetchDataHO("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & Replace(rsHlp("KODE"), Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|"))), "") & "' and MDL = '" & Trim(rsHlp4("MDL")) & "' and FS_LOG_DEL = ''")
133                                       If rsHlp2.RecordCount > 0 Then
134                                           Set rsHlp3 = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & Replace(rsHlp("KODE"), Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|"))), "") & "' and MDL = '" & Trim(rsHlp4("MDL")) & "' and FS_LOG_DEL = ''")
135                                           If rsHlp3.RecordCount > 0 Then
136                                               Call objO.UpdateData("UPDATE HO_FLATE SET FS_GRP_MDL = '" & rsHlp2("FS_GRP_MDL") & "', MODEL = '" & rsHlp2("MODEL") & "', JOB = '" & rsHlp2("JOB") & "', REMARK = '" & rsHlp2("REMARK") & "', RATE = '" & Replace(CDbl(rsHlp2("RATE")), ",", ".") & "', SBE = '" & rsHlp2("SBE") & "'" _
                                                                       & "KETERANGAN = '" & rsHlp2("KETERANGAN") & "', TIPEBP = '" & rsHlp2("TIPEBP") & "', HARGA = '" & CDbl(rsHlp2("HARGA")) & "', FS_LOG_DEL = '" & Trim(rsHlp2("FS_LOG_DEL")) & "', FD_TGL_DEL = '" & Format(rsHlp2("FD_TGL_DEL"), "MM/DD/YYYY HH:NN:SS") & "' WHERE FS_ITEM = '" & Replace(rsHlp("KODE"), Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|"))), "") & "' and MDL = '" & Trim(rsHlp4("MDL")) & "' and FS_LOG_DEL = ''")
137                                               Call objO.UpdateData("UPDATE ACTRATE SET MODEL = '" & rsHlp2("MODEL") & "', JOB = '" & rsHlp2("JOB") & "', REMARK = '" & rsHlp2("REMARK") & "', RATE = '" & Replace((CDbl(rsHlp2("RATE")) * 0.8), ",", ".") & "', SBE = '" & rsHlp2("SBE") & "'" _
                                                                       & "KETERANGAN = '" & rsHlp2("KETERANGAN") & "', FS_LOG_DEL = '" & Trim(rsHlp2("FS_LOG_DEL")) & "', FD_TGL_DEL = '" & Format(rsHlp2("FD_TGL_DEL"), "MM/DD/YYYY HH:NN:SS") & "', Other = 0 WHERE FS_ITEM = '" & Replace(rsHlp("KODE"), Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|"))), "") & "' and MDL = '" & Trim(rsHlp4("MDL")) & "' and FS_LOG_DEL = ''")
138                                           Else
139                                               Call objO.InsertData("INSERT INTO HO_FLATE VALUES ('" & rsHlp2("FS_GRP_MDL") & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("JOB") & "','" & rsHlp2("REMARK") & "', '" & rsHlp2("FS_ITEM") & "', '" & Replace(CDbl(rsHlp2("RATE")), ",", ".") & "', '" & rsHlp2("KETERANGAN") & "', '" & rsHlp2("TIPEBP") & "', '" & CDbl(rsHlp2("HARGA")) & "', '" & Trim(rsHlp2("FS_LOG_DEL")) & "', '" & rsHlp2("FD_TGL_DEL") & "', '" & rsHlp2("SBE") & "')")
140                                               Call objO.InsertData("INSERT INTO ACTRATE VALUES ( '" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("JOB") & "','" & rsHlp2("REMARK") & "', '" & rsHlp2("FS_ITEM") & "', '','','" & Replace((CDbl(rsHlp2("RATE")) * 0.8), ",", ".") & "', '" & rsHlp2("KETERANGAN") & "', 0, '" & rsHlp2("SBE") & "','" & Trim(rsHlp2("FS_LOG_DEL")) & "', '" & rsHlp2("FD_TGL_DEL") & "')")
141                                           End If
142                                           rsHlp4.MoveNext
143                                       End If
144                                   Loop
145                               End If
146                           ElseIf Trim(rsHlp("STATUS")) = "D" Then
147                               Set rsHlp4 = objO.FetchData("Select * from Model WITH (NOLOCK) where left(MDL," & Len(Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|") - 1))) & ") = '" & Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|") - 1)) & "'")
148                               If rsHlp4.RecordCount > 0 Then
149                                   rsHlp4.MoveFirst
150                                   Do Until rsHlp4.EOF
151                                       Call objO.UpdateData("UPDATE ho_flate SET FD_TGL_DEL = getdate(), FS_LOG_DEL = 'HO' WHERE mdl like '" & Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|") - 1)) & "-%' AND fs_item = '" & Replace(rsHlp("KODE"), Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|"))), "") & "' and FS_LOG_DEL = ''")
152                                       Call objO.UpdateData("UPDATE ACTRATE SET FD_TGL_DEL = getdate(), FS_LOG_DEL = 'HO' WHERE mdl like '" & Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|") - 1)) & "-%' AND fs_item = '" & Replace(rsHlp("KODE"), Left(rsHlp("KODE"), (InStr(1, rsHlp("KODE"), "|"))), "") & "' and FS_LOG_DEL = ''")
153                                       rsHlp4.MoveNext
154                                   Loop
155                               End If
156                           End If
157                       Case "KELAS"
158                           If Trim(rsHlp("STATUS")) = "S" Then
                                   
159                               Set rsHlp2 = objO.FetchDataHO("SELECT * FROM CARCLASS WITH (NOLOCK) WHERE CLASSCODE = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
160                               If Not rsHlp2.EOF Then
161                                   Call objO.DeleteData("DELETE CARCLASS WHERE CLASSCODE = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
162                                   Call objO.InsertData("INSERT INTO CARCLASS VALUES ('" & rsHlp2("CLASSCODE") & "', '" & rsHlp2("CLASSNAME") & "', '" & CDbl(rsHlp2("PRICE")) & "', '" & rsHlp2("NOTE") & "','" & CDbl(rsHlp2("WAR")) & "', '','','" & CDbl(rsHlp2("PRICETAM")) & "','" & CDbl(rsHlp2("DiscNRM")) & "','" & CDbl(rsHlp2("DiscNON")) & "')")
      '                                Set rsHlp3 = objO.FetchData("SELECT * FROM CARCLASS WHERE CLASSCODE = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
      '                                If Not rsHlp3.EOF Then
      '                                    Call objO.UpdateData("UPDATE CARCLASS SET CLASSNAME = '" & rsHlp2("CLASSNAME") & "', PRICE = '" & CDbl(rsHlp2("PRICE")) & "', NOTE = '" & rsHlp2("NOTE") & "', WAR = '" & CDbl(rsHlp2("WAR")) & "', PRICETAM = '" & CDbl(rsHlp2("PRICETAM")) & "' WHERE CLASSCODE = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
      '                                Else
      '                                    Call objO.InsertData("INSERT INTO CARCLASS VALUES ('" & rsHlp2("CLASSCODE") & "', '" & rsHlp2("CLASSNAME") & "', '" & CDbl(rsHlp2("PRICE")) & "', '" & rsHlp2("NOTE") & "','" & CDbl(rsHlp2("WAR")) & "', '','','" & CDbl(rsHlp2("PRICETAM")) & "')")
      '                                End If
      '                                rsHlp3.Close
163                               End If
164                               rsHlp2.Close
165                           ElseIf Trim(rsHlp("STATUS")) = "D" Then
166                               Call objO.DeleteData("DELETE CARCLASS WHERE CLASSCODE = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
                                  'Call objO.UpdateData("UPDATE CARCLASS SET FD_TGL_DEL = getdate(), FS_LOG_DEL = 'HO' WHERE CLASSCODE = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
167                           End If
168                       Case "DEALER"
169                           If Trim(rsHlp("STATUS")) = "S" Then
                                   
170                               Set rsHlp2 = objO.FetchDataHO("SELECT * FROM T_DEALER WITH (NOLOCK) WHERE FS_KD_DEALER = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
171                               If Not rsHlp2.EOF Then
                                       
172                                   Set rsHlp3 = objO.FetchData("SELECT * FROM T_DEALER WITH (NOLOCK) WHERE FS_KD_DEALER = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
173                                   If Not rsHlp3.EOF Then
174                                       Call objO.UpdateData("UPDATE T_DEALER SET FS_NM_DEALER = '" & rsHlp2("FS_NM_DEALER") & "', FS_ALAMAT = '" & rsHlp2("FS_ALAMAT") & "', FS_KOTA = '" & rsHlp2("FS_KOTA") & "', " _
                                                       & "FS_NPWP = '" & rsHlp2("FS_NPWP") & "', FS_PKP = '" & rsHlp2("FS_PKP") & "', KODECABANG = '" & rsHlp2("KODECABANG") & "', VPNCABANG = '" & rsHlp2("VPNCABANG") & "', USERCBG = '" & rsHlp2("USERCBG") & "', " _
                                                       & "PASSCBG = '" & rsHlp2("PASSCBG") & "', INTERNAL = '" & CDbl(rsHlp2("INTERNAL")) & "', KDPLG = '" & Trim(rsHlp2("KDPLG")) & "', FS_NM_TAGIH = '" & Trim(rsHlp2("FS_NM_TAGIH")) & "' WHERE FS_KD_DEALER = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
175                                   Else
                                          'Call objO.InsertData("INSERT INTO T_DEALER VALUES ('" & rsHlp2("FS_KD_DEALER") & "', '" & rsHlp2("FS_NM_DEALER") & "', '" & rsHlp2("FS_ALAMAT") & "', '" & rsHlp2("FS_KOTA") & "', '" & rsHlp2("FS_NPWP") & "', '" & rsHlp2("FS_PKP") & "', " _
                                                       & "'" & rsHlp2("KODECABANG") & "', '" & rsHlp2("VPNCABANG") & "', '" & rsHlp2("USERCBG") & "', '" & rsHlp2("PASSCBG") & "','" & CDbl(rsHlp2("INTERNAL")) & "', '" & Trim(rsHlp2("KDPLG")) & "','','','" & Trim(rsHlp2("FS_NM_TAGIH")) & "')")
176                                       Call objO.InsertData("INSERT INTO T_DEALER (FS_KD_DEALER, FS_NM_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, FS_PKP, KODECABANG, VPNCABANG, userCbg, passCbg, INTERNAL, KDPLG, FS_LOG_DEL, FD_TGL_DEL, FS_NM_TAGIH, HARGANRM) VALUES ('" & rsHlp2("FS_KD_DEALER") & "', '" & rsHlp2("FS_NM_DEALER") & "', '" & rsHlp2("FS_ALAMAT") & "', '" & rsHlp2("FS_KOTA") & "', '" & rsHlp2("FS_NPWP") & "', '" & rsHlp2("FS_PKP") & "', " _
                                                       & "'" & rsHlp2("KODECABANG") & "', '" & rsHlp2("VPNCABANG") & "', '" & rsHlp2("USERCBG") & "', '" & rsHlp2("PASSCBG") & "','" & CDbl(rsHlp2("INTERNAL")) & "', '" & Trim(rsHlp2("KDPLG")) & "','','','" & Trim(rsHlp2("FS_NM_TAGIH")) & "','" & Trim(rsHlp2("harganrm")) & "')")

177                                   End If
178                                   rsHlp3.Close
179                               End If
180                               rsHlp2.Close
181                           ElseIf Trim(rsHlp("STATUS")) = "D" Then
182                               Call objO.DeleteData("DELETE T_DEALER WHERE FS_KD_DEALER = '" & rsHlp("KODE") & "'")
183                           End If
184                       Case "WARRANTY"
185                           If Trim(rsHlp("STATUS")) = "S" Then
186                               Set rsHlp2 = objO.FetchDataHO("SELECT * FROM HO_WARRANTY A WITH (NOLOCK) LEFT JOIN HO_WARDTL B WITH (NOLOCK) ON A.ID_WARRANTY = B.ID_WARRANTY WHERE A.ID_WARRANTY = '" & rsHlp("KODE") & "' AND A.FS_LOG_DEL = ''")
187                               If Not rsHlp2.EOF Then
188                                   Set rsHlp3 = objO.FetchData("SELECT * FROM HO_WARRANTY WITH (NOLOCK) WHERE ID_WARRANTY = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
189                                   If Not rsHlp3.EOF Then
190                                       Call objO.UpdateData("UPDATE HO_WARRANTY SET WARRANTY = '" & rsHlp2("WARRANTY") & "', KETERANGAN = '" & rsHlp2("KETERANGAN") & "', TGLMULAI = '" & Format(rsHlp2("TGLMULAI"), "MM/DD/YYYY") & "', " _
                                                       & "TGLSELESAI = '" & Format(rsHlp2("TGLSELESAI"), "MM/DD/YYYY") & "', NOTI = '" & rsHlp2("NOTI") & "', TGLTI = '" & Format(rsHlp2("TGLTI"), "MM/DD/YYYY") & "' WHERE ID_WARRANTY = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
191                                       Call objO.DeleteData("DELETE HO_WARDTL WHERE ID_WARRANTY = '" & rsHlp("KODE") & "'")
192                                       If Trim(Nul(rsHlp2("MDL"))) <> "" Then
193                                           Do While Not rsHlp2.EOF
194                                               Call objO.InsertData("INSERT INTO HO_WARDTL VALUES ('" & rsHlp2("ID_WARRANTY") & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("NO_RANGKA1") & "', '" & rsHlp2("NO_RANGKA2") & "')")
195                                               rsHlp2.MoveNext
196                                           Loop
197                                       End If
198                                       Set rsHlp4 = objO.FetchDataHO("SELECT * FROM HO_WARJOBPART WITH (NOLOCK) WHERE ID_WARRANTY = '" & rsHlp("KODE") & "'")
199                                       If Not rsHlp4.EOF Then
200                                           Call objO.DeleteData("DELETE HO_WARJOBPART WHERE ID_WARRANTY = '" & rsHlp("KODE") & "'")
201                                           Do While Not rsHlp4.EOF
202                                               Call objO.InsertData("INSERT INTO HO_WARJOBPART VALUES ('" & rsHlp4("ID_WARRANTY") & "', '" & rsHlp4("MDL") & "', '" & rsHlp4("MODEL") & "', '" & rsHlp4("JOB") & "', '" & rsHlp4("FS_ITEM") & "', " _
                                                               & "'" & rsHlp4("JOBKET") & "', '" & Replace(rsHlp4("RATE"), ",", ".") & "', '" & rsHlp4("ITEMPART") & "', '" & rsHlp4("NAMAPART") & "', '" & rsHlp4("SATPART") & "', '" & rsHlp4("QTY") & "', " _
                                                               & "'" & rsHlp4("FLAG") & "', '" & rsHlp4("KETERANGAN") & "')")
203                                               rsHlp4.MoveNext
204                                           Loop
205                                       End If
206                                       rsHlp4.Close
207                                   Else
208                                       Call objO.DeleteData("DELETE HO_WARRANTY WHERE ID_WARRANTY = '" & rsHlp2("ID_WARRANTY") & "' and FS_LOG_DEL <> ''")
209                                       Call objO.InsertData("INSERT INTO HO_WARRANTY VALUES ('" & rsHlp2("ID_WARRANTY") & "', '" & rsHlp2("WARRANTY") & "', '" & rsHlp2("KETERANGAN") & "', '" & Format(rsHlp2("TGLMULAI"), "MM/DD/YYYY") & "', '" & Format(rsHlp2("TGLSELESAI"), "MM/DD/YYYY") & "', " _
                                                       & "'" & rsHlp2("NOTI") & "', '" & Format(rsHlp2("TGLTI"), "MM/DD/YYYY") & "', '','')")
210                                       If Trim(Nul(rsHlp2("MDL"))) <> "" Then
211                                           Do While Not rsHlp2.EOF
212                                               Call objO.InsertData("INSERT INTO HO_WARDTL VALUES ('" & rsHlp2("ID_WARRANTY") & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("NO_RANGKA1") & "', '" & rsHlp2("NO_RANGKA2") & "')")
213                                               rsHlp2.MoveNext
214                                           Loop
215                                       End If
216                                       Set rsHlp4 = objO.FetchDataHO("SELECT * FROM HO_WARJOBPART WITH (NOLOCK) WHERE ID_WARRANTY = '" & rsHlp("KODE") & "'")
217                                       If Not rsHlp4.EOF Then
218                                           Call objO.DeleteData("DELETE HO_WARJOBPART WHERE ID_WARRANTY = '" & rsHlp("KODE") & "'")
219                                           Do While Not rsHlp4.EOF
220                                               Call objO.InsertData("INSERT INTO HO_WARJOBPART VALUES ('" & rsHlp4("ID_WARRANTY") & "', '" & rsHlp4("MDL") & "', '" & rsHlp4("MODEL") & "', '" & rsHlp4("JOB") & "', '" & rsHlp4("FS_ITEM") & "', " _
                                                               & "'" & rsHlp4("JOBKET") & "', '" & Replace(rsHlp4("RATE"), ",", ".") & "', '" & rsHlp4("ITEMPART") & "', '" & rsHlp4("NAMAPART") & "', '" & rsHlp4("SATPART") & "', '" & rsHlp4("QTY") & "', " _
                                                               & "'" & rsHlp4("FLAG") & "', '" & rsHlp4("KETERANGAN") & "')")
221                                               rsHlp4.MoveNext
222                                           Loop
223                                       End If
224                                       rsHlp4.Close
225                                   End If
226                                   rsHlp3.Close
227                               End If
228                               rsHlp2.Close
229                           ElseIf Trim(rsHlp("STATUS")) = "D" Then
230                               Call objO.DeleteData("DELETE HO_WARRANTY WHERE ID_WARRANTY = '" & rsHlp("KODE") & "'")
231                               Call objO.DeleteData("DELETE HO_WARDTL WHERE ID_WARRANTY = '" & rsHlp("KODE") & "'")
232                               Call objO.DeleteData("DELETE HO_WARJOBPART WHERE ID_WARRANTY = '" & rsHlp("KODE") & "'")
233                           End If
234                       Case "PROGRAM"
235                           If Trim(rsHlp("STATUS")) = "S" Then
                                   
236                               Set rsHlp2 = objO.FetchDataHO("SELECT A.*, GRPMDL,MDL, [DESC] FROM HO_PROGRAM A WITH (NOLOCK) LEFT JOIN HO_PROMDL B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.ID_PROGRAM = '" & rsHlp("KODE") & "' and A.FS_LOG_DEL = ''")
237                               If Not rsHlp2.EOF Then
                                       
238                                   Set rsHlp3 = objO.FetchData("SELECT * FROM HO_PROGRAM WITH (NOLOCK) WHERE ID_PROGRAM = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
239                                   If Not rsHlp3.EOF Then
240                                       Call objO.UpdateData("UPDATE HO_PROGRAM SET PROGRAM = '" & rsHlp2("PROGRAM") & "', KETERANGAN = '" & rsHlp2("KETERANGAN") & "', TGLMULAI = '" & Format(rsHlp2("TGLMULAI"), "MM/DD/YYYY") & "', " _
                                                       & "TGLSELESAI = '" & Format(rsHlp2("TGLSELESAI"), "MM/DD/YYYY") & "', DISCJASA = '" & rsHlp2("DISCJASA") & "', FLAGPART = '" & rsHlp2("FLAGPART") & "', KODEPART = '" & rsHlp2("KODEPART") & "', " _
                                                       & "NAMAPART = '" & rsHlp2("NAMAPART") & "', DISCPART = '" & rsHlp2("DISCPART") & "', DEALERDAFT = '" & rsHlp2("DEALERDAFT") & "', NAMADEALER = '" & rsHlp2("NAMADEALER") & "', " _
                                                       & "FLAGDEALER = '" & rsHlp2("FLAGDEALER") & "', DOUBLEINV = '" & rsHlp2("DOUBLEINV") & "', FLAGMODEL = '" & rsHlp2("FLAGMODEL") & "' WHERE ID_PROGRAM = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
241                                       Call objO.DeleteData("DELETE HO_PROMDL WHERE ID_PROGRAM = '" & rsHlp("KODE") & "'")
242                                       If Trim(Nul(rsHlp2("MDL"))) <> "" Then
243                                           Do While Not rsHlp2.EOF
244                                               Call objO.InsertData("INSERT INTO HO_PROMDL VALUES ('" & rsHlp2("ID_PROGRAM") & "', '" & rsHlp2("GRPMDL") & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("DESC") & "')")
245                                               rsHlp2.MoveNext
246                                           Loop
247                                       End If
248                                       Set rsHlp4 = objO.FetchDataHO("SELECT * FROM HO_PROCBG WITH (NOLOCK) WHERE ID_PROGRAM = '" & rsHlp("KODE") & "'")
249                                       If Not rsHlp4.EOF Then
250                                           Call objO.DeleteData("DELETE HO_PROCBG WHERE ID_PROGRAM = '" & rsHlp("KODE") & "'")
251                                           Do While Not rsHlp4.EOF
252                                               Call objO.InsertData("INSERT INTO HO_PROCBG VALUES ('" & rsHlp4("ID_PROGRAM") & "', '" & rsHlp4("KODEDEALER") & "', '" & rsHlp4("NAMADEALER") & "')")
253                                               rsHlp4.MoveNext
254                                           Loop
255                                       End If
256                                       rsHlp4.Close
257                                   Else
258                                       Call objO.InsertData("INSERT INTO HO_PROGRAM VALUES ('" & rsHlp2("ID_PROGRAM") & "', '" & rsHlp2("PROGRAM") & "', '" & rsHlp2("KETERANGAN") & "', '" & Format(rsHlp2("TGLMULAI"), "MM/DD/YYYY") & "', '" & Format(rsHlp2("TGLSELESAI"), "MM/DD/YYYY") & "', " _
                                                       & "'" & rsHlp2("DISCJASA") & "', '" & rsHlp2("FLAGPART") & "', '" & rsHlp2("KODEPART") & "', '" & rsHlp2("NAMAPART") & "', '" & rsHlp2("DISCPART") & "','" & rsHlp2("DEALERDAFT") & "','" & rsHlp2("NAMADEALER") & "','" & rsHlp2("FLAGDEALER") & "', '" & rsHlp2("DOUBLEINV") & "','" & rsHlp2("FLAGMODEL") & "','','')")
259                                       If Trim(Nul(rsHlp2("MDL"))) <> "" Then
260                                           Do While Not rsHlp2.EOF
261                                               Call objO.InsertData("INSERT INTO HO_PROMDL VALUES ('" & rsHlp2("ID_PROGRAM") & "', '" & rsHlp2("GRPMDL") & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("DESC") & "')")
262                                               rsHlp2.MoveNext
263                                           Loop
264                                       End If
265                                       Set rsHlp4 = objO.FetchDataHO("SELECT * FROM HO_PROCBG WITH (NOLOCK) WHERE ID_PROGRAM = '" & rsHlp("KODE") & "'")
266                                       If Not rsHlp4.EOF Then
267                                           Call objO.DeleteData("DELETE HO_PROCBG WHERE ID_PROGRAM = '" & rsHlp("KODE") & "'")
268                                           Do While Not rsHlp4.EOF
269                                               Call objO.InsertData("INSERT INTO HO_PROCBG VALUES ('" & rsHlp4("ID_PROGRAM") & "', '" & rsHlp4("KODEDEALER") & "', '" & rsHlp4("NAMADEALER") & "')")
270                                               rsHlp4.MoveNext
271                                           Loop
272                                       End If
273                                   End If
274                               End If
275                           ElseIf Trim(rsHlp("STATUS")) = "D" Then
276                               Call objO.DeleteData("DELETE HO_PROGRAM WHERE ID_PROGRAM = '" & rsHlp("KODE") & "'")
277                               Call objO.DeleteData("DELETE HO_PROMDL WHERE ID_PROGRAM = '" & rsHlp("KODE") & "'")
278                           End If
279                       Case "MOU"
280                           If Trim(rsHlp("STATUS")) = "S" Then
                                   
281                               Set rsHlp2 = objO.FetchDataHO("SELECT * FROM MOUHDR A WITH (NOLOCK) LEFT JOIN MOUDTL B WITH (NOLOCK) ON A.NO_MOU = B.NO_MOU WHERE A.NO_MOU = '" & rsHlp("KODE") & "' AND A.FS_LOG_DEL = ''")
282                               If Not rsHlp2.EOF Then
                                       
283                                   Set rsHlp3 = objO.FetchData("SELECT * FROM MOUHDR WITH (NOLOCK) WHERE NO_MOU = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
284                                   If Not rsHlp3.EOF Then
285                                       objO.UpdateData ("UPDATE MOUHDR SET PERUSAHAAN = '" & rsHlp2("PERUSAHAAN") & "', PIC = '" & rsHlp2("PIC") & "', ALAMAT = '" & rsHlp2("ALAMAT") & "', KOTA = '" & rsHlp2("KOTA") & "', NPWP = '" & rsHlp2("NPWP") & "', GSM = '" & rsHlp2("GSM") & "', " _
                                                       & "CDMA = '" & rsHlp2("CDMA") & "', PSTN = '" & rsHlp2("PSTN") & "', FAX = '" & rsHlp2("FAX") & "', TGLMULAI = '" & Format(rsHlp2("TGLMULAI"), "MM/DD/YYYY") & "', TGLSELESAI = '" & Format(rsHlp2("TGLSELESAI"), "MM/DD/YYYY") & "', [TOP] = '" & Trim(rsHlp2("TOP")) & "', " _
                                                       & "DISCJASA = '" & rsHlp2("DISCJASA") & "', DISCPART = '" & rsHlp2("DISCPART") & "', DEALERDAFT = '" & rsHlp2("DEALERDAFT") & "', NAMADEALER = '" & rsHlp2("NAMADEALER") & "', FLAGDEALER = '" & rsHlp2("FLAGDEALER") & "', " _
                                                       & "LIMITAR = " & CDbl(rsHlp2("LIMITAR")) & ", HUTANG = " & CDbl(rsHlp2("HUTANG")) & ", BAYAR = " & CDbl(rsHlp2("BAYAR")) & ", USERLOG = '" & rsHlp2("USERLOG") & "', INPUTDATE = '" & Format(rsHlp2("INPUTDATE"), "MM/DD/YYYY") & "' WHERE NO_MOU = '" & rsHlp("KODE") & "' AND FS_LOG_DEL = ''")
286                                       objO.DeleteData ("DELETE MOUDTL WHERE NO_MOU = '" & rsHlp("KODE") & "'")
                                      
287                                       Do While Not rsHlp2.EOF
288                                           If Trim(Nul(rsHlp2("MDL"))) <> "" Then
289                                               objO.InsertData ("INSERT INTO MOUDTL VALUES ('" & rsHlp2("NO_MOU") & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("NO_POLISI") & "', '" & rsHlp2("NO_RANGKA") & "')")
290                                           End If
291                                           rsHlp2.MoveNext
292                                       Loop
293                                       Set rsHlp4 = objO.FetchDataHO("SELECT * FROM MOUCBG WITH (NOLOCK) WHERE NO_MOU = '" & rsHlp("KODE") & "'")
294                                       If Not rsHlp4.EOF Then
295                                           Call objO.DeleteData("DELETE MOUCBG WHERE NO_MOU = '" & rsHlp("KODE") & "'")
296                                           Do While Not rsHlp4.EOF
297                                               Call objO.InsertData("INSERT INTO MOUCBG VALUES ('" & rsHlp4("NO_MOU") & "', '" & rsHlp4("KODEDEALER") & "', '" & rsHlp4("NAMADEALER") & "', '" & rsHlp4("DISCJASA") & "', '" & rsHlp4("DISCPART") & "')")
298                                               rsHlp4.MoveNext
299                                           Loop
300                                       End If
301                                   Else
302                                       Call objO.DeleteData("DELETE MOUHDR WHERE NO_MOU = '" & rsHlp("KODE") & "' and FS_LOG_DEL <> ''")
303                                       Call objO.InsertData("INSERT INTO MOUHDR VALUES ('" & Trim(rsHlp2.Fields(0).value) & "', '" & rsHlp2("PERUSAHAAN") & "', '" & rsHlp2("PIC") & "', '" & rsHlp2("ALAMAT") & "', '" & rsHlp2("KOTA") & "', '" & rsHlp2("NPWP") & "', " _
                                                       & "'" & rsHlp2("GSM") & "', '" & rsHlp2("CDMA") & "', '" & rsHlp2("PSTN") & "', '" & rsHlp2("FAX") & "','" & Format(rsHlp2("TGLMULAI"), "MM/DD/YYYY") & "', '" & Format(rsHlp2("TGLSELESAI"), "MM/DD/YYYY") & "', " _
                                                       & "'" & rsHlp2("TOP") & "', '" & rsHlp2("DISCJASA") & "', '" & rsHlp2("DISCPART") & "', '" & rsHlp2("DEALERDAFT") & "', '" & rsHlp2("NAMADEALER") & "', '" & rsHlp2("FLAGDEALER") & "', " _
                                                       & "" & CDbl(rsHlp2("LIMITAR")) & ", " & CDbl(rsHlp2("HUTANG")) & ", " & CDbl(rsHlp2("BAYAR")) & ", '" & rsHlp2("USERLOG") & "', '" & Format(rsHlp2("INPUTDATE"), "MM/DD/YYYY") & "','','')")
304                                       If Trim(Nul(rsHlp2("MDL"))) <> "" Then
305                                           Do While Not rsHlp2.EOF
306                                               Call objO.InsertData("INSERT INTO MOUDTL VALUES ('" & Trim(rsHlp2.Fields(0).value) & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("NO_POLISI") & "', '" & rsHlp2("NO_RANGKA") & "')")
307                                               rsHlp2.MoveNext
308                                           Loop
309                                       End If
310                                       Set rsHlp4 = objO.FetchDataHO("SELECT * FROM MOUCBG WITH (NOLOCK) WHERE NO_MOU = '" & rsHlp("KODE") & "'")
311                                       If Not rsHlp4.EOF Then
312                                           Call objO.DeleteData("DELETE MOUCBG WHERE NO_MOU = '" & rsHlp("KODE") & "'")
313                                           Do While Not rsHlp4.EOF
314                                               Call objO.InsertData("INSERT INTO MOUCBG VALUES ('" & rsHlp4("NO_MOU") & "', '" & rsHlp4("KODEDEALER") & "', '" & rsHlp4("NAMADEALER") & "', '" & rsHlp4("DISCJASA") & "', '" & rsHlp4("DISCPART") & "')")
315                                               rsHlp4.MoveNext
316                                           Loop
317                                       End If
318                                   End If
319                               End If
320                           ElseIf Trim(rsHlp("STATUS")) = "D" Then
321                               Call objO.DeleteData("DELETE MOUHDR WHERE NO_MOU = '" & rsHlp("KODE") & "'")
322                               Call objO.DeleteData("DELETE MOUDTL WHERE NO_MOU = '" & rsHlp("KODE") & "'")
323                               Call objO.DeleteData("DELETE MOUCBG WHERE NO_MOU = '" & rsHlp("KODE") & "'")
324                           End If
325                       Case "EVENT"
326                           If Trim(rsHlp("STATUS")) = "S" Then
                                   
327                               Set rsHlp2 = objO.FetchDataHO("SELECT * FROM HO_EVENT WITH (NOLOCK) WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "' AND FS_LOG_DEL = ''")
328                               If Not rsHlp2.EOF Then
                                       
329                                   Set rsHlp3 = objO.FetchData("SELECT * FROM HO_EVENT WITH (NOLOCK) WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "' AND FS_LOG_DEL = ''")
330                                   If Not rsHlp3.EOF Then
331                                       Call objO.UpdateData("UPDATE HO_EVENT SET NamaEvent = '" & Trim(rsHlp2("namaEvent")) & "',tipeOrder = '" & Trim(rsHlp2("tipeOrder")) & "',thVehicle1 = '" & CDbl(rsHlp2("thVehicle1")) & "',thVehicle2 = '" & CDbl(rsHlp2("thVehicle2")) & "',KM1 = '" & CDbl(rsHlp2("KM1")) & "',KM2 = '" & CDbl(rsHlp2("KM2")) & "',lastService1 = '" & CDbl(rsHlp2("lastService1")) & "',lastService2 = '" & CDbl(rsHlp2("lastService2")) & "',birthDay1 = '" & Format(rsHlp2("birthDay1"), "MM/DD/YYYY") & "',birthDay2 = '" & Format(rsHlp2("birthDay2"), "MM/DD/YYYY") & "',hobby = '" & Trim(rsHlp2("hobby")) & "', " _
                                                       & "DelivDay1 = '" & Format(rsHlp2("DelivDay1"), "MM/DD/YYYY") & "',DelivDay2 = '" & Format(rsHlp2("DelivDay2"), "MM/DD/YYYY") & "',modelVehichle = '" & Trim(rsHlp2("modelVehichle")) & "',dealerBerlaku = '" & Trim(rsHlp2("dealerBerlaku")) & "',berlaku1 = '" & Format(rsHlp2("berlaku1"), "MM/DD/YYYY") & "',berlaku2 = '" & Format(rsHlp2("berlaku2"), "MM/DD/YYYY") & "', textEvent = '" & Trim(rsHlp2("textEvent")) & "',aktif = '" & CDbl(rsHlp2("aktif")) & "',KodeProgram = '" & Trim(rsHlp2("KodeProgram")) & "',fs_Log_Add = '" & Trim(rsHlp2("fs_Log_Add")) & "', " _
                                                       & "fs_Tgl_Add = '" & Format(rsHlp2("fs_Tgl_Add"), "MM/DD/YYYY") & "',fs_Log_DeAktifasi = '" & Trim(rsHlp2("fs_Log_DeAktifasi")) & "',fs_Tgl_DeAktifasi = '" & Format(rsHlp2("fs_Tgl_DeAktifasi"), "MM/DD/YYYY") & "',fs_Log_Del = '" & Trim(rsHlp2("fs_Log_Del")) & "',fs_Tgl_Del = '" & Format(rsHlp2("fs_Tgl_Del"), "MM/DD/YYYY") & "', diReminder = '" & Trim(rsHlp2("direminder")) & "' WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "' AND FS_LOG_DEL = ''")
332                                   Else
333                                       Call objO.InsertData("INSERT INTO HO_EVENT VALUES ('" & Trim(rsHlp2("noEvent")) & "','" & Trim(rsHlp2("namaEvent")) & "','" & Trim(rsHlp2("tipeOrder")) & "','" & CDbl(rsHlp2("thVehicle1")) & "','" & CDbl(rsHlp2("thVehicle2")) & "','" & CDbl(rsHlp2("KM1")) & "','" & CDbl(rsHlp2("KM2")) & "','" & CDbl(rsHlp2("lastService1")) & "','" & CDbl(rsHlp2("lastService2")) & "','" & Format(rsHlp2("birthDay1"), "MM/DD/YYYY") & "','" & Format(rsHlp2("birthDay2"), "MM/DD/YYYY") & "','" & Trim(rsHlp2("hobby")) & "','" & Format(rsHlp2("DelivDay1"), "MM/DD/YYYY") & "','" & Format(rsHlp2("DelivDay2"), "MM/DD/YYYY") & "','" & Trim(rsHlp2("modelVehichle")) & "','" & Trim(rsHlp2("dealerBerlaku")) & "','" & Format(rsHlp2("berlaku1"), "MM/DD/YYYY") & "','" & Format(rsHlp2("berlaku2"), "MM/DD/YYYY") & "', " _
                                                       & "'" & Trim(rsHlp2("textEvent")) & "','" & CDbl(rsHlp2("aktif")) & "','" & Trim(rsHlp2("KodeProgram")) & "','" & Trim(rsHlp2("fs_Log_Add")) & "','" & Format(rsHlp2("fs_Tgl_Add"), "MM/DD/YYYY") & "','" & Trim(rsHlp2("fs_Log_DeAktifasi")) & "','" & Format(rsHlp2("fs_Tgl_DeAktifasi"), "MM/DD/YYYY") & "','" & Trim(rsHlp2("fs_Log_Del")) & "','" & Format(rsHlp2("fs_Tgl_Del"), "MM/DD/YYYY") & "', '" & Trim(rsHlp2("diReminder")) & "')")
334                                   End If
                                      
335                                   Set rsHlp4 = objO.FetchDataHO("SELECT * FROM HO_EVTDEALER WITH (NOLOCK) WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "'")
336                                   If Not rsHlp4.EOF Then
337                                       Do While Not rsHlp4.EOF
                                               
338                                           Set rsHlp3 = objO.FetchData("SELECT * FROM HO_EVTDEALER WITH (NOLOCK) WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "' And KDDEALER = '" & Trim(rsHlp4("kdDealer")) & "'")
339                                           If Not rsHlp3.EOF Then
340                                               Call objO.UpdateData("UPDATE HO_EVTDEALER SET kdDealer = '" & Trim(rsHlp4("kdDealer")) & "',namaDealer = '" & Trim(rsHlp4("NAMADealer")) & "' WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "'")
341                                           Else
342                                               Call objO.InsertData("INSERT INTO HO_EVTDEALER VALUES ('" & Trim(rsHlp4("noEvent")) & "','" & Trim(rsHlp4("kdDealer")) & "','" & Trim(rsHlp4("namaDealer")) & "')")
343                                           End If
344                                           rsHlp4.MoveNext
345                                       Loop
346                                   End If
                                      
347                                   Set rsHlp4 = objO.FetchDataHO("SELECT * FROM HO_EVTMODEL WITH (NOLOCK) WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "'")
348                                   If Not rsHlp4.EOF Then
349                                       Do While Not rsHlp4.EOF
                                               
350                                           Set rsHlp3 = objO.FetchData("SELECT * FROM HO_EVTMODEL WITH (NOLOCK) WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "' AND MDL = '" & Trim(rsHlp4("MDL")) & "'")
351                                           If Not rsHlp3.EOF Then
352                                               Call objO.UpdateData("UPDATE HO_EVTMODEL SET GRPMDL = '" & Trim(rsHlp4("GRPMDL")) & "',MDL = '" & Trim(rsHlp4("MDL")) & "',[DESC] = '" & Trim(rsHlp4("DESC")) & "' WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "'")
353                                           Else
354                                               Call objO.InsertData("INSERT INTO HO_EVTMODEL VALUES ('" & Trim(rsHlp("KODE")) & "','" & Trim(rsHlp4("GRPMDL")) & "','" & Trim(rsHlp4("MDL")) & "','" & Trim(rsHlp4("DESC")) & "')")
355                                           End If
356                                           rsHlp4.MoveNext
357                                       Loop
358                                   End If
                                      
359                               End If
360                           ElseIf Trim(rsHlp("STATUS")) = "D" Then
361                               Call objO.DeleteData("DELETE HO_EVENT WHERE NOEVENT = '" & Trim(rsHlp("KODE")) & "'")
362                           End If
363                       Case "PERIODE"
364                           Set rsHlp2 = objO.FetchDataHO("SELECT [week],[minLs],[maxLS] FROM master_setting WITH (NOLOCK)")
365                           If Not rsHlp2.EOF Then
366                               rsHlp2.MoveFirst
367                               Do While Not rsHlp2.EOF
368                                   Set rsHlp3 = objO.FetchData("SELECT * FROM master_setting WITH (NOLOCK) WHERE [week] = '" & NulSID(rsHlp2!Week) & "'")
369                                   If Not rsHlp3.EOF Then
370                                       Call objO.UpdateData("UPDATE master_setting SET minLs = '" & NulSID(rsHlp2!MINLS) & "',maxLS = '" & NulSID(rsHlp2!MAXLS) & "' WHERE [week] = '" & NulSID(rsHlp2!Week) & "'")
371                                   Else
372                                       Call objO.InsertData("INSERT INTO master_setting([week],[minLs],[maxLS]) VALUES ('" & NulSID(rsHlp2!Week) & "','" & NulSID(rsHlp2!MINLS) & "','" & NulSID(rsHlp2!MAXLS) & "')")
373                                   End If
374                                   rsHlp2.MoveNext
375                               Loop
376                           End If
377                   End Select
378                   Call objO.DeleteData("DELETE TEMPKIRIM WHERE TIPE = '" & Trim(rsHlp("TIPE")) & "' AND KODE = '" & Trim(rsHlp("KODE")) & "'")
379               End If
380               End If
381               rsHlp.MoveNext
382           Loop
383       End If
           
           
384       Set rsHlp = Nothing
385       Set rsHlp2 = Nothing
386       Set rsHlp3 = Nothing
387       Set rsHlp4 = Nothing

388      On Error GoTo 0
389      Exit Sub

cekFileImport_Error:

390       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekFileImport of Form Layar")
End Sub

Private Sub cekReminder()
      Dim rsLama As ADODB.Recordset
1        On Error GoTo cekReminder_Error

2         Set rsHlp = objO.FetchData("SELECT DISTINCT NORANGKA FROM M_ESTIMASI A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NORANGKA = B.NO_CHASSIS " _
                                  & "WHERE DATEDIFF(DAY, DATEADD(day, 15, CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), GETDATE()) = 0 " _
                                  & "AND ISNULL(JT_TGL,'') <> '' ORDER BY NORANGKA")
                                  
3         If Not rsHlp.EOF Then
4             rsHlp.MoveFirst
5             Do While Not rsHlp.EOF
6                 Set rsLama = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS TGL_RM, *, HEAVYUSE FROM M_ESTIMASI A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NORANGKA = B.NO_CHASSIS WHERE NORANGKA = '" & Trim(rsHlp("NORANGKA")) & "' WHERE ISNULL(JT_TGL,'') <> '' ORDER BY SBENEXT DESC")
7                 If Not rsLama.EOF Then
                      'objO.UpdateData ("EXEC NextEstimasi '" & Trim(rsLama("NORANGKA")) & "', '" & Format(Trim(rsLama("TGL_RM")), "MM/DD/YYYY") & "', '" & CDbl(rsLama("KMNEXT")) & "', '" & CDbl(rsLama("SBENEXT")) & "', '" & IIf(rsLama("HEAVYUSE") = 1, 5000, 10000) & "', '', '','" & Trim(rsLama("NAMACUST")) & "', '" & Trim(rsLama("ALAMATCUST")) & "','" & Trim(rsLama("NOPOLISI")) & "', '" & Trim(rsLama("PSTN")) & "', '" & Trim(rsLama("GSM")) & "', '" & Trim(rsLama("CDMA")) & "'")
8                     objO.UpdateData ("EXEC " & Trim(hTipePenghitunganDealer) & " '" & Trim(rsLama("NORANGKA")) & "', '" & Format(Trim(rsLama("TGL_RM")), "MM/DD/YYYY") & "', '" & CDbl(rsLama("KMNEXT")) & "', '" & CDbl(rsLama("SBENEXT")) & "', '" & IIf(rsLama("HEAVYUSE") = 1, 5000, 10000) & "', '', '','" & Trim(rsLama("NAMACUST")) & "', '" & Trim(rsLama("ALAMATCUST")) & "','" & Trim(rsLama("NOPOLISI")) & "', '" & Trim(rsLama("PSTN")) & "', '" & Trim(rsLama("GSM")) & "', '" & Trim(rsLama("CDMA")) & "'")
9                 End If
10                rsHlp.MoveNext
11            Loop
12        End If

13       On Error GoTo 0
14       Exit Sub

cekReminder_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekReminder of Form Layar")
           
End Sub

Private Sub MDIForm_Resize()
 If Me.WindowState <> 1 Then Me.WindowState = 2
 objA.SetLeft = Me.Width - 3280
 objA.SetTop = 270
 fHeight = Me.StatusBar1.Top
End Sub

Private Sub pic4text_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel <> "" Then
        Call cekTampungan
        pic4text.Panels(1).Text = statusKoneksi
    End If
End Sub


Private Sub Timer1_Timer()

      '    Exit Sub
1        On Error GoTo Timer1_Timer_Error

2         gTglTampung = gTglServer
3         gTanggal = TglDDMMYYYY(Format(gTglTampung, "DD-MM-YYYY"))
          
4         StatusBar1.Panels(2).Text = " " + Format(gTglTampung, "Long Date") + " "
5         StatusBar1.Panels(3).Text = " " + Format(gTglTampung, "HH:NN") + " "
6         If onTimer = False Then Exit Sub
          
7         If gUstat <> 1 Then
              'StatusBar1.Panels(1).text = Log.LogHOConn & " IP : " & Trim(gServerDb) + " - HO : " & objO.FetchData_String("Select Count(*) from LogHO WITH (NOLOCK) Where LogHO = 0")
8             If Log.RPLHistori = 0 Then
9                 StatusBar1.Panels(1).Text = Log.LogHOConn & " IP : " & Trim(gServerDb) + " - HO : " & objO.FetchData_String("Select Count(*) from LogHO WITH (NOLOCK) Where LogHO = 0")
10            Else
11                If gTipeUser = "SA" And Log.newCetak = 1 Then
12                    StatusBar1.Panels(1).Text = "IP : " & Trim(gServerDb) '+ " - Belum Invoice : " & objO.FetchData_String("Select Count(*) from jobdta WITH (NOLOCK) Where deliv <> 'btl' and noinv = '' and front = '" & gSA & "'")
13                Else
14                    StatusBar1.Panels(1).Text = "IP : " & Trim(gServerDb)
15                End If
                  
16            End If
17        Else
18            StatusBar1.Panels(1).Text = " IP : " & Trim(gServerDb)
19        End If
20        StatusBar1.Panels(5).Text = "U.E : " + Format(isiUnit, "###,0") + " "
21        StatusBar1.Panels(11).Text = " Login : " & gSA
22        StatusBar1.Panels(7).Text = "Lock.AR : " & objAR.FetchDataLockAR_ & ", Overdue H-3 : " & objAR.FetchDataCust_OnAR_Overdue_(3)
          
23        If nC10 < IIf(gTipeUser = "FOREMAN", 12, IIf(gTipeUser = "SA", 120, 120)) And nC10 <> 0 Then
24            nC10 = nC10 + 1
25        Else
              Dim rsSearch As ADODB.Recordset
              
26                Set rsHlp = objO.FetchData_Profil("SELECT * FROM PROFIL WITH (NOLOCK)")
27                If Not rsHlp.EOF Then
28                    gDbVersi = Trim(rsHlp("VERSION"))
29                    gDbVersiLama = Trim(rsHlp("VERSIONLAMA"))
30                    gDbTanggal = Trim(rsHlp("TGLAKTIF"))
31                    If CInt(rsHlp("HRSAMA")) = 1 Then
32                        gHrSama = True
33                    Else
34                        gHrSama = False
35                    End If
36                End If
                   
                  
37                If DateDiff("d", gTglTampung, gDbTanggal) > 0 Then
38                    If Trim(gDbVersiLama) <> Trim(gDbVersi) Then
39                        If Trim(gDbVersiLama) = Trim(gVersi) Then
40                            MsgBox "Download NIS Service Terbaru : " & Trim(gDbVersi) & vbCr & vbCr & "Versi Lama : " & Trim(gDbVersiLama) & " akan di-LOCK pada Tanggal " & Format(gDbTanggal, "DD MMMM YYYY") & vbCr & vbCr & vbCr & "KURANG " & DateDiff("d", isnow, gDbTanggal) & " Hari.", vbCritical, "Update Versi Lama Anda (Versi Anda : " & gDbVersiLama & ")"
41                            allowIn = True
42                        ElseIf Trim(gDbVersi) = Trim(gVersi) Then
43                            allowIn = True
44                        End If
45                    Else
46                        allowIn = True
47                    End If
48                ElseIf DateDiff("d", gTglTampung, gDbTanggal) <= 0 Then
49                    If Trim(gDbVersi) = Trim(gVersi) Then
50                        allowIn = True
51                    End If
52                End If
53            If gUstat = 0 Then
54                Call isiStatusBar
                  Dim tanggal1 As String
                  Dim tanggal2 As String
55                If Month(isnow) = 1 Then
56                   tanggal1 = Left(Right(year(isnow) - 1, 2) & "12", 3)
57                Else
58                   tanggal1 = Left(Right(year(isnow), 2) & Format(Month(isnow) - 1, "0#"), 3)
59                End If
                  
60                Select Case Month(isnow)
                     Case 1
61                      tanggal2 = Left(Right(year(isnow) - 1, 2) & "11", 3)
62                   Case 2
63                      tanggal2 = Left(Right(year(isnow) - 1, 2) & "12", 3)
64                   Case Else
65                      tanggal2 = Left(Right(year(isnow), 2) & Format(Month(isnow) - 2, "0#"), 3)
66                End Select
                  
                  'Call objO.InsertData("update Stall set flag = 'N' where Keterangan not like '%|160%' and Tanggal > '04/01/2016' and flag = ''")
                  'Call objO.InsertData("update Stall set flag = 'N' where Keterangan not like '%|" & Left(Right(year(isnow), 2) & Format(Month(isnow), "0#"), 3) & "%' and Tanggal > '04/01/2016' and flag = ''")
                  ''2016-10-03 || penambahan untuk bulan kemarin dan 2 bulan sebelum
67                Call objO.InsertData("update Stall set flag = 'N' where (Keterangan not like '%|" & Left(Right(year(isnow), 2) & Format(Month(isnow), "0#"), 3) & "%' and Keterangan not like '%|" & tanggal1 & "%' and Keterangan not like '%|" & tanggal2 & "%') and Tanggal > '04/01/2016' and flag = ''")
68                Call objO.InsertData("DELETE TempProgressCuci WHERE NoOrder = ''")
                  
69                If Trim(gTipeUser) = "SA" Or Trim(gTipeUser) = "FOREMAN" Then
70                ElseIf Trim(gTipeUser) = "SA" Then
                       
      '63                    Set rsHlp = objO.FetchData("SELECT * FROM T_STALL WITH (NOLOCK) WHERE HARI = '" & IIf(Weekday(gTglTampung) = 1, 6, Weekday(gTglTampung) - 2) & WeekdayName(IIf(Weekday(gTglTampung) = 1, 8, Weekday(gTglTampung)) - 1) & "'")
      '64                    If Not rsHlp.EOF Then
      '65                        tJamSelesai = IIf(Trim(rsHlp("WAKTUSELESAI")) = "24:00", "00:00", Trim(rsHlp("WAKTUSELESAI")))
      '66                    End If
                       
71                ElseIf Trim(gTipeUser) = "MRS" Then
72                    Call adaNoShow
73                ElseIf Trim(gTipeUser) = "BILLING" Then
                      ' koneksi HO
74                    Me.pic4text.Visible = True
75                    If Len(Trim(pic4text.Panels(1).Text)) > 4 Then
76                        pic4text.Panels(1).Text = Right(pic4text.Panels(1).Text, Len(pic4text.Panels(1).Text) - 7)
77                    ElseIf Len(pic4text.Panels(1).Text) <= 4 Then
78                        pic4text.Panels(1).Text = "                                                                                                                                                                                                                                                                                                                      " + statusKoneksi
79                    End If
80                End If
81            Else
82                If gUstat <> 1 Then
                      'StatusBar1.Panels(1).text = Log.LogHOConn & " IP : " & Trim(gServerDb) + " - HO : " & objO.FetchData_String("Select Count(*) from LogHO WITH (NOLOCK) Where LogHO = 0")
83                    If Log.RPLHistori = 0 Then
84                        StatusBar1.Panels(1).Text = Log.LogHOConn & " IP : " & Trim(gServerDb) + " - HO : " & objO.FetchData_String("Select Count(*) from LogHO WITH (NOLOCK) Where LogHO = 0")
85                    Else
86                        StatusBar1.Panels(1).Text = "IP : " & Trim(gServerDb)
87                    End If
88                Else
89                    StatusBar1.Panels(1).Text = " IP : " & Trim(gServerDb)
90                End If
91                StatusBar1.Panels(2).Text = Format(gTglTampung, "Long Date")
92                StatusBar1.Panels(11).Text = " Login: " & gSA
93            End If
              'If fgCabang Then
              ''20180221
94            If fgCabang And Log.RPLHistori = 0 Then
95                Call cekFileImport
96            End If
              
97            If gHo And Log.RPLHistori = 0 Then
98                Call cekFileExport
99            End If
100           nC10 = 1
101       End If
102       StatusBar1.Panels(3).Text = Format(gTglTampung, "HH:NN")
103       If Len(Trim(pic4text.Panels(1).Text)) > 10 Then
104           pic4text.Panels(1).Text = Right(pic4text.Panels(1).Text, Len(pic4text.Panels(1).Text) - 7)
105       ElseIf Len(pic4text.Panels(1).Text) <= 10 Then
106           pic4text.Panels(1).Text = "                                                                                                                                                                                                                                                                                                                      " + statusKoneksi
107       End If
          
108       If bookAktif Then
109           Call kedipWarranty(frmBooking)
110           frmBooking.Text(19) = Format(gTglTampung, "HH:NN")
111       End If
112       If orderAktif Then
113           Call kedipWarranty(frmOrder)
              Dim nWKT As String
114           If frmOrder.cmdPilih.Caption = "No. Rangka" Then
115               frmOrder.Text(66) = Format(gTglTampung, "HH:NN")
116           End If
117           If frmOrder.Text(86) <> ">>:>>" And frmOrder.Text(65) <> ">>->>->>>>" Then
118               Call frmOrder.hitungEstimasiPenyerahan
                  
      '            nWKT = hitungWaktu(IIf(frmOrder.chEM.value = Checked, 1, CDbl(frmOrder.text(114))), frmOrder.text(86), frmOrder.text(65))
      '            frmOrder.text(63) = Left(nWKT, InStr(nWKT, "|") - 1)
      '            frmOrder.text(62) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
119           End If
120           If frmOrder.frmEvent.Visible Then
121               If frmOrder.frmEvent.BackColor = &HC0FFFF Then
122                   frmOrder.frmEvent.BackColor = &H80C0FF
123                   frmOrder.sgnEvent.ForeColor = vbBlue
124               Else
125                   frmOrder.frmEvent.BackColor = &HC0FFFF
126                   frmOrder.sgnEvent.ForeColor = vbRed
127               End If
128           End If
129       End If
          

130      On Error GoTo 0
131      Exit Sub

Timer1_Timer_Error:

132       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Timer1_Timer of Form Layar")
End Sub

Function panggilLayar()
    With viewLayar
        .ColumnHeaders.Add , , "Aktif", 200
        .ColumnHeaders.Add , , "No", 400
        .ColumnHeaders.Add , , "FormAktif", 2300
        
        Dim lv As ListItem
        If createLayar Then
            Do While Not rsLayar.EOF
                Set lv = .ListItems.Add(, , 1)
                lv.SubItems(1) = rsLayar("AKTIF")
                lv.SubItems(2) = rsLayar("NO")
                lv.SubItems(3) = rsLayar("LAYARAKTIF")
                rsLayar.MoveNext
            Loop
            picLayar.Visible = True
            picLayar.Top = 1000
            picLayar.Left = 1440
        End If
    End With
End Function

Private Function createLayar() As Boolean
1        On Error GoTo createLayar_Error

2         createLayar = False
3         If rsLayar.State Then rsLayar.Close
4         Set rsLayar = Nothing
5         With rsLayar
              'If .State = adStateClosed Then
6                 .Fields.Append "AKTIF", adVarChar, 1
7                 .Fields.Append "NO", adNumeric
8                 .Fields.Append "LAYARAKTIF", adVarChar, 50
9                 .Open
              'End If
              
10            If .State <> adStateClosed Then
11                createLayar = True
12            End If
13        End With

14       On Error GoTo 0
15       Exit Function

createLayar_Error:

16        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createLayar of Form Layar")
End Function

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
    If Panel.Index = 1 Then
        If Log.RPLHistori = 0 Then
            Call Do_Transfer_To_HO
        Else
            'MsgBox "Data telah terreplikasi", vbInformation
            Exit Sub
        End If
    End If
    If Panel.Index = 7 Then Call ViewDataAR
    If Panel.Index = 6 Then
        If gTipeUser = "MRS" Then
            frmBookingMenuOnline.show
            frmBookingMenuOnline.WindowState = vbMaximized
        End If
    End If
    
    If Panel.Index = 9 And Log.OTD = 1 Then
      If gTipeUser = "SA" Then
         frmMonitoringOTD.monitoringSA = ""
         frmMonitoringOTD.show
      ElseIf gTipeUser <> "Billing" Then
           frmMonitoringOTDRekap.show
           
      End If
    End If
    
End Sub

Private Sub Do_Transfer_To_HO()
1        On Error GoTo Do_Transfer_To_HO_Error

2         If Right(StatusBar1.Panels(1).Text, 6) = "HO : 0" Then Exit Sub
3         If MsgBox("Transfer Data ke HO ?", vbQuestion + vbYesNo) = vbYes Then
4             If Get_ConnectionProfil_Bengkel99 = False Then
5                 MsgBox "Ada Kesalahan Setting ke HO," & Chr(13) & "Cek VPN dan konektivitas ke HO" & Chr(13) & "(Lakukan Routing IP)", vbCritical + vbOKOnly
6             Else
7                 Call ProsesTransfer_HO
8                 Call objO.UpdateData("Update T_Profil Set KoneksiHO = 'OK'")
9                 Log.LogHOConn = 1
10            End If
11            StatusBar1.Panels(1).Text = Log.LogHOConn & " IP : " & Trim(gServerDb) + " - HO : " & objO.FetchData_String("Select Count(*) from LogHO WITH (NOLOCK) Where LogHO = 0")
12        End If

13       On Error GoTo 0
14       Exit Sub

Do_Transfer_To_HO_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Do_Transfer_To_HO of Form Layar")
End Sub

Private Sub Do_Transfer_To_HO_Run()
1        On Error GoTo Do_Transfer_To_HO_Run_Error

2         If Right(StatusBar1.Panels(1).Text, 6) = "HO : 0" Then Exit Sub
3         If Left(StatusBar1.Panels(1).Text, 1) = "0" Then Exit Sub
4         If Get_ConnectionProfil_Bengkel99 = True Then
5             Call ProsesTransfer_HO
6             Call objO.UpdateData("Update T_Profil Set KoneksiHO = 'OK'")
7             Log.LogHOConn = 1
8         End If
9         StatusBar1.Panels(1).Text = Log.LogHOConn & " IP : " & Trim(gServerDb) + " - HO : " & objO.FetchData_String("Select Count(*) from LogHO WITH (NOLOCK) Where LogHO = 0")

10       On Error GoTo 0
11       Exit Sub

Do_Transfer_To_HO_Run_Error:

12        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Do_Transfer_To_HO_Run of Form Layar")
End Sub

Private Sub ProsesTransfer_HO()
          Dim rsT As ADODB.Recordset
          Dim isBegin As Boolean
1         isBegin = False
2        On Error GoTo ProsesTransfer_HO_Error
3         Set rsT = objO.FetchData("Select Top 1 * FROM LogHO WITH (NOLOCK) where LogHO.LogHO = 0")
4         If rsT.RecordCount = 0 Then Exit Sub
          Dim nkode As String
5         nkode = Trim(rsT.Fields(1).value)
         
6         Set rsT = objO.FetchData("Select * FROM LogHO WITH (NOLOCK) where LogHO.LogHO = 0 And LogHO.Kode = '" & nkode & "' Order By Tipe ASC")
      '               Set rsT = objO.FetchData("Select * FROM LogHO WITH (NOLOCK) where LogHO.LogHO = 0 And LogHO.tipe = 'DETAILWAC' Order By Tipe ASC")
7         rsT.MoveFirst
8         Do Until rsT.EOF
9       If Trim(rsT.Fields(0).value) = "BATCH" Then
            '==== TRANSFER KE HO =========
            'Call objHT.OpenTransactionHO
            
            'Call objHT.CloseTransactionHO
      '            Call objO.InsertData("insert Into OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKomp & ";Uid=" & Log.LogHOUser & ";PWD=" & Log.LogHOPass & "', [bengkel99].dbo.T_InvoiceNew) " & _
      '                                 "Select CUSTOMER, CUST_ADRES, KOTA, TELPON, NO_POLISI, NO_ENGINE, NO_CHASSIS, MDL, MODEL, FMDL, COLOR, FS_KD_DEALER, FS_NM_DEALER, DDATE, CODE, NPWP, PPN, DISC, DISCT, FRONT, " & _
      '                                 "[GROUP], MECH, KM, DATE, TIME, TIME_M, TIME_S, TIME_A, PART_G, PART_NG, MTRL, OIL_G, OIL_NG, MET, SBLT, LABOR, BODY, TOTAL, JOB, REQ1, REQ2, REQ3, REQ4, REQ5, REP1, REP2, REP3, " & _
      '                                 "REP4, REP5, REP6, REP7, REP8, SLIP, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, COND1, COND2, COND3, COND4, COND5, COND6, COND7, INVDATE, NO_WO, NOINV, STATUS, PAID, NOFAK, " & _
      '                                 "REMARK1, SUGGEST, DELIV, SPK, RPR_COND, MEKANIK1, MEKANIK2, MEKANIK3, MEKANIK4, MEKANIK5, MEKANIK6, MEKANIK7, MEKANIK8, GRP1, GRP2, GRP3, GRP4, GRP5, GRP6, GRP7, GRP8, ACTRATE1, " & _
      '                                 "ACTRATE2, ACTRATE3, ACTRATE4, ACTRATE5, ACTRATE6, ACTRATE7, ACTRATE8, OFF1, OFF2, OFF3, OFF4, OFF5, OFF6, OFF7, OFF8, KLPKERJA1, KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, KLPKERJA6, " & _
      '                                 "KLPKERJA7, KLPKERJA8, FS_KREDIT, FN_PAJAK, FS_TH_MOBIL , FN_RETURN, FS_CATATAN, FS_LOG_ADD, FS_LOG_DEL, FN_JML_ORDER, FN_JML_RATE, ALASAN_INV,FS_RC , FS_KD_PLG, FS_KD_WIL, FD_TGL_PART, " & _
      '                                 "FS_NO_RETUR, FD_RETUR, FS_KD_RETUR, TPSVC, THS, pds, MOU, KODETIPE, CMPL1, CMPL2, CMPL3, CMPL4, CMPL5, INVTIME, DISC1, DISC2, DISC3, DISC4, DISC5, DISC6, DISC7, DISC8, NoFaktur, VALID1, " & _
      '                                 "VALID2, VALID3, VALID4, VALID5, VALID6, VALID7, VALID8, FREEC1, FREEC2, FREEC3, FREEC4, FREEC5, FREEC6, FREEC7, FREEC8, noFP, NOMB, TGLKASIR, NOFP2, KASIRADD, CASHADD, TGLCASH,'" & Log.LogBengkel & "' " & _
      '                                 "from T_Invoice WITH (NOLOCK) Where No_WO = '" & Trim(rsT.Fields(1).value) & "'")
      '
            
10          Set rsHlp = objO.FetchData("Select CUSTOMER, CUST_ADRES, KOTA, TELPON, NO_POLISI, NO_ENGINE, NO_CHASSIS, MDL, MODEL, FMDL, COLOR, FS_KD_DEALER, FS_NM_DEALER, DDATE, CODE, NPWP, PPN, DISC, DISCT, FRONT, " & _
                                 "[GROUP], MECH, KM, DATE, TIME, TIME_M, TIME_S, TIME_A, PART_G, PART_NG, MTRL, OIL_G, OIL_NG, MET, SBLT, LABOR, BODY, TOTAL, JOB, REQ1, REQ2, REQ3, REQ4, REQ5, REP1, REP2, REP3, " & _
                                 "REP4, REP5, REP6, REP7, REP8, SLIP, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, COND1, COND2, COND3, COND4, COND5, COND6, COND7, INVDATE, NO_WO, NOINV, STATUS, PAID, NOFAK, " & _
                                 "REMARK1, SUGGEST, DELIV, SPK, RPR_COND, MEKANIK1, MEKANIK2, MEKANIK3, MEKANIK4, MEKANIK5, MEKANIK6, MEKANIK7, MEKANIK8, GRP1, GRP2, GRP3, GRP4, GRP5, GRP6, GRP7, GRP8, ACTRATE1, " & _
                                 "ACTRATE2, ACTRATE3, ACTRATE4, ACTRATE5, ACTRATE6, ACTRATE7, ACTRATE8, OFF1, OFF2, OFF3, OFF4, OFF5, OFF6, OFF7, OFF8, KLPKERJA1, KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, KLPKERJA6, " & _
                                 "KLPKERJA7, KLPKERJA8, FS_KREDIT, FN_PAJAK, FS_TH_MOBIL , FN_RETURN, FS_CATATAN, FS_LOG_ADD, FS_LOG_DEL, FN_JML_ORDER, FN_JML_RATE, ALASAN_INV,FS_RC , FS_KD_PLG, FS_KD_WIL, FD_TGL_PART, " & _
                                 "FS_NO_RETUR, FD_RETUR, FS_KD_RETUR, TPSVC, THS, pds, MOU, KODETIPE, CMPL1, CMPL2, CMPL3, CMPL4, CMPL5, INVTIME, DISC1, DISC2, DISC3, DISC4, DISC5, DISC6, DISC7, DISC8, NoFaktur, VALID1, " & _
                                 "VALID2, VALID3, VALID4, VALID5, VALID6, VALID7, VALID8, FREEC1, FREEC2, FREEC3, FREEC4, FREEC5, FREEC6, FREEC7, FREEC8, noFP, NOMB, TGLKASIR, NOFP2, KASIRADD, CASHADD, TGLCASH,'" & Log.LogBengkel & "' " & _
                                 "from T_Invoice WITH (NOLOCK) Where No_WO = '" & Trim(rsT.Fields(1).value) & "'")
11          If Not rsHlp.EOF Then
12              rsHlp.MoveFirst
13              Call Transfer_INV1(Trim(rsT.Fields(1).value))
14              Call ClearsSql
15              Do While Not rsHlp.EOF
16                  gsSql1(0) = "CUSTOMER"
17                  gsSql2(0) = Nul(rsHlp!customer)
18                  gsSql1(1) = "CUST_ADRES"
19                  gsSql2(1) = Nul(rsHlp!cust_adres)
20                  gsSql1(2) = "KOTA"
21                  gsSql2(2) = Nul(rsHlp!Kota)
22                  gsSql1(3) = "TELPON"
23                  gsSql2(3) = Nul(rsHlp!Telpon)
24                  gsSql1(4) = "NO_POLISI"
25                  gsSql2(4) = Nul(rsHlp!NO_POLISI)
                    
26                  gsSql1(5) = "NO_ENGINE"
27                  gsSql2(5) = Nul(rsHlp!No_Engine)
28                  gsSql1(6) = "NO_CHASSIS"
29                  gsSql2(6) = Nul(rsHlp!NO_CHASSIS)
30                  gsSql1(7) = "MDL"
31                  gsSql2(7) = Nul(rsHlp!MDL)
32                  gsSql1(8) = "MODEL"
33                  gsSql2(8) = Nul(rsHlp!model)
34                  gsSql1(9) = "FMDL"
35                  gsSql2(9) = Nul(rsHlp!Fmdl)
36                  gsSql1(10) = "COLOR"
37                  gsSql2(10) = Nul(rsHlp!Color)
                    
38                  gsSql1(11) = "FS_KD_DEALER"
39                  gsSql2(11) = Nul(rsHlp!FS_KD_DEALER)
40                  gsSql1(12) = "FS_NM_DEALER"
41                  gsSql2(12) = Nul(rsHlp!fs_nm_dealer)
42                  gsSql1(13) = "DDATE"
43                  gsSql2(13) = Format(Nul(rsHlp!dDate), "MM/DD/YYYY")
44                  gsSql1(14) = "CODE"
45                  gsSql2(14) = Nul(rsHlp!CODE)
46                  gsSql1(15) = "NPWP"
47                  gsSql2(15) = Nul(rsHlp!npwp)
                    
48                  gsSql1(16) = "PPN"
49                  gsSql2(16) = Nul(rsHlp!PPn)
50                  gsSql1(17) = "DISC"
51                  gsSql2(17) = Nul(rsHlp!disc)
52                  gsSql1(18) = "DISCT"
53                  gsSql2(18) = Nul(rsHlp!DISCT)
54                  gsSql1(19) = "FRONT"
55                  gsSql2(19) = Nul(rsHlp!FRONT)
56                  gsSql1(20) = "[GROUP]"
57                  gsSql2(20) = Nul(rsHlp!Group)
                    
58                  gsSql1(21) = "MECH"
59                  gsSql2(21) = Nul(rsHlp!Mech)
60                  gsSql1(22) = "KM"
61                  gsSql2(22) = Nul(rsHlp!km)
62                  gsSql1(23) = "DATE"
63                  gsSql2(23) = Format(Nul(rsHlp!Date), "MM/DD/YYYY")
64                  gsSql1(24) = "TIME"
65                  gsSql2(24) = Nul(rsHlp!Time)
66                  gsSql1(25) = "TIME_M"
67                  gsSql2(25) = Nul(rsHlp!TIME_M)
                    
68                  gsSql1(26) = "TIME_S"
69                  gsSql2(26) = Nul(rsHlp!TIME_S)
70                  gsSql1(27) = "TIME_A"
71                  gsSql2(27) = Nul(rsHlp!TIME_A)
72                  gsSql1(28) = "PART_G"
73                  gsSql2(28) = Replace(Nul(rsHlp!PART_G), ",", ".")
74                  gsSql1(29) = "PART_NG"
75                  gsSql2(29) = Replace(Nul(rsHlp!PART_NG), ",", ".")
76                  gsSql1(30) = "MTRL"
77                  gsSql2(30) = Replace(Nul(rsHlp!MTRL), ",", ".")
                    
78                  gsSql1(31) = "OIL_G"
79                  gsSql2(31) = Replace(Nul(rsHlp!OIL_G), ",", ".")
80                  gsSql1(32) = "OIL_NG"
81                  gsSql2(32) = Replace(Nul(rsHlp!OIL_NG), ",", ".")
82                  gsSql1(33) = "MET"
83                  gsSql2(33) = Replace(Nul(rsHlp!met), ",", ".")
84                  gsSql1(34) = "SBLT"
85                  gsSql2(34) = Replace(Nul(rsHlp!sblt), ",", ".")
86                  gsSql1(35) = "LABOR"
87                  gsSql2(35) = Replace(Nul(rsHlp!labor), ",", ".")
                    
88                  gsSql1(36) = "BODY"
89                  gsSql2(36) = Replace(Nul(rsHlp!BODY), ",", ".")
90                  gsSql1(37) = "TOTAL"
91                  gsSql2(37) = Replace(Nul(rsHlp!total), ",", ".")
92                  gsSql1(38) = "JOB"
93                  gsSql2(38) = Nul(rsHlp!job)
94                  gsSql1(39) = "REQ1"
95                  gsSql2(39) = Nul(rsHlp!REQ1)
96                  gsSql1(40) = "REQ2"
97                  gsSql2(40) = Nul(rsHlp!REQ2)
                    
98                  gsSql1(41) = "REQ3"
99                  gsSql2(41) = Nul(rsHlp!REQ3)
100                 gsSql1(42) = "REQ4"
101                 gsSql2(42) = Nul(rsHlp!REQ4)
102                 gsSql1(43) = "REQ5"
103                 gsSql2(43) = Nul(rsHlp!REQ5)
104                 gsSql1(44) = "REP1"
105                 gsSql2(44) = Nul(rsHlp!rep1)
106                 gsSql1(45) = "REP2"
107                 gsSql2(45) = Nul(rsHlp!REP2)
                    
108                 gsSql1(46) = "REP3"
109                 gsSql2(46) = Nul(rsHlp!REP3)
110                 gsSql1(47) = "REP4"
111                 gsSql2(47) = Nul(rsHlp!REP4)
112                 gsSql1(48) = "REP5"
113                 gsSql2(48) = Nul(rsHlp!REP5)
114                 gsSql1(49) = "REP6"
115                 gsSql2(49) = Nul(rsHlp!REP6)
116                 gsSql1(50) = "REP7"
117                 gsSql2(50) = Nul(rsHlp!REP7)
                    
118                 gsSql1(51) = "REP8"
119                 gsSql2(51) = Nul(rsHlp!REP8)
120                 gsSql1(52) = "SLIP"
121                 gsSql2(52) = Nul(rsHlp!Slip)
122                 gsSql1(53) = "RTE1"
123                 gsSql2(53) = Replace(Nul(rsHlp!RTE1), ",", ".")
124                 gsSql1(54) = "RTE2"
125                 gsSql2(54) = Replace(Nul(rsHlp!RTE2), ",", ".")
126                 gsSql1(55) = "RTE3"
127                 gsSql2(55) = Replace(Nul(rsHlp!RTE3), ",", ".")
                    
128                 gsSql1(56) = "RTE4"
129                 gsSql2(56) = Replace(Nul(rsHlp!RTE4), ",", ".")
130                 gsSql1(57) = "RTE5"
131                 gsSql2(57) = Replace(Nul(rsHlp!RTE5), ",", ".")
132                 gsSql1(58) = "RTE6"
133                 gsSql2(58) = Replace(Nul(rsHlp!RTE6), ",", ".")
134                 gsSql1(59) = "RTE7"
135                 gsSql2(59) = Replace(Nul(rsHlp!RTE7), ",", ".")
136                 gsSql1(60) = "RTE8"
137                 gsSql2(60) = Replace(Nul(rsHlp!RTE8), ",", ".")
                    
138                 gsSql1(61) = "COND1"
139                 gsSql2(61) = Nul(rsHlp!COND1)
140                 gsSql1(62) = "COND2"
141                 gsSql2(62) = Nul(rsHlp!COND2)
142                 gsSql1(63) = "COND3"
143                 gsSql2(63) = Nul(rsHlp!COND3)
144                 gsSql1(64) = "COND4"
145                 gsSql2(64) = Nul(rsHlp!COND4)
146                 gsSql1(65) = "COND5"
147                 gsSql2(65) = Nul(rsHlp!COND5)
                    
148                 gsSql1(66) = "COND6"
149                 gsSql2(66) = Nul(rsHlp!COND6)
150                 gsSql1(67) = "COND7"
151                 gsSql2(67) = Nul(rsHlp!COND7)
152                 gsSql1(68) = "INVDATE"
153                 gsSql2(68) = Format(Nul(rsHlp!InvDate), "MM/DD/YYYY")
154                 gsSql1(69) = "NO_WO"
155                 gsSql2(69) = Nul(rsHlp!no_wo)
156                 gsSql1(70) = "NOINV"
157                 gsSql2(70) = Nul(rsHlp!noInv)
                    
158                 gsSql1(71) = "STATUS"
159                 gsSql2(71) = Nul(rsHlp!Status)
160                 gsSql1(72) = "PAID"
161                 gsSql2(72) = Nul(rsHlp!paid)
162                 gsSql1(73) = "NOFAK"
163                 gsSql2(73) = Nul(rsHlp!NoFak)
164                 gsSql1(74) = "REMARK1"
165                 gsSql2(74) = Nul(rsHlp!REMARK1)
166                 gsSql1(75) = "SUGGEST"
167                 gsSql2(75) = Nul(rsHlp!SUGGEST)
                    
168                 gsSql1(76) = "DELIV"
169                 gsSql2(76) = Nul(rsHlp!DELIV)
170                 gsSql1(77) = "SPK"
171                 gsSql2(77) = Nul(rsHlp!SPK)
172                 gsSql1(78) = "RPR_COND"
173                 gsSql2(78) = Nul(rsHlp!RPR_COND)
174                 gsSql1(79) = "MEKANIK1"
175                 gsSql2(79) = Nul(rsHlp!mekanik1)
176                 gsSql1(80) = "MEKANIK2"
177                 gsSql2(80) = Nul(rsHlp!mekanik2)
                    
178                 gsSql1(81) = "MEKANIK3"
179                 gsSql2(81) = Nul(rsHlp!MEKANIK3)
180                 gsSql1(82) = "MEKANIK4"
181                 gsSql2(82) = Nul(rsHlp!MEKANIK4)
182                 gsSql1(83) = "MEKANIK5"
183                 gsSql2(83) = Nul(rsHlp!MEKANIK5)
184                 gsSql1(84) = "MEKANIK6"
185                 gsSql2(84) = Nul(rsHlp!MEKANIK6)
186                 gsSql1(85) = "MEKANIK7"
187                 gsSql2(85) = Nul(rsHlp!MEKANIK7)
                    
188                 gsSql1(86) = "MEKANIK8"
189                 gsSql2(86) = Nul(rsHlp!MEKANIK8)
190                 gsSql1(87) = "GRP1"
191                 gsSql2(87) = Nul(rsHlp!GRP1)
192                 gsSql1(88) = "GRP2"
193                 gsSql2(88) = Nul(rsHlp!GRP2)
194                 gsSql1(89) = "GRP3"
195                 gsSql2(89) = Nul(rsHlp!GRP3)
196                 gsSql1(90) = "GRP4"
197                 gsSql2(90) = Nul(rsHlp!GRP4)
                    
198                 gsSql1(91) = "GRP5"
199                 gsSql2(91) = Nul(rsHlp!GRP5)
200                 gsSql1(92) = "GRP6"
201                 gsSql2(92) = Nul(rsHlp!GRP6)
202                 gsSql1(93) = "GRP7"
203                 gsSql2(93) = Nul(rsHlp!GRP7)
204                 gsSql1(94) = "GRP8"
205                 gsSql2(94) = Nul(rsHlp!GRP8)
206                 gsSql1(95) = "ACTRATE1"
207                 gsSql2(95) = Replace(Nul(rsHlp!ACTRATE1), ",", ".")
                    
208                 gsSql1(96) = "ACTRATE2"
209                 gsSql2(96) = Replace(Nul(rsHlp!ACTRATE2), ",", ".")
210                 gsSql1(97) = "ACTRATE3"
211                 gsSql2(97) = Replace(Nul(rsHlp!ACTRATE3), ",", ".")
212                 gsSql1(98) = "ACTRATE4"
213                 gsSql2(98) = Replace(Nul(rsHlp!ACTRATE4), ",", ".")
214                 gsSql1(99) = "ACTRATE5"
215                 gsSql2(99) = Replace(Nul(rsHlp!ACTRATE5), ",", ".")
216                 gsSql1(100) = "ACTRATE6"
217                 gsSql2(100) = Replace(Nul(rsHlp!ACTRATE6), ",", ".")
                    
218                 gsSql1(101) = "ACTRATE7"
219                 gsSql2(101) = Replace(Nul(rsHlp!ACTRATE7), ",", ".")
220                 gsSql1(102) = "ACTRATE8"
221                 gsSql2(102) = Replace(Nul(rsHlp!ACTRATE8), ",", ".")
222                 gsSql1(103) = "OFF1"
223                 gsSql2(103) = RepRp(CDbl(Nul(rsHlp!Off1)))
224                 gsSql1(104) = "OFF2"
225                 gsSql2(104) = RepRp(CDbl(Nul(rsHlp!Off2)))
226                 gsSql1(105) = "OFF3"
227                 gsSql2(105) = RepRp(CDbl(Nul(rsHlp!Off3)))
                    
228                 gsSql1(106) = "OFF4"
229                 gsSql2(106) = RepRp(CDbl(Nul(rsHlp!Off4)))
230                 gsSql1(107) = "OFF5"
231                 gsSql2(107) = RepRp(CDbl(Nul(rsHlp!Off5)))
232                 gsSql1(108) = "OFF6"
233                 gsSql2(108) = RepRp(CDbl(Nul(rsHlp!Off6)))
234                 gsSql1(109) = "OFF7"
235                 gsSql2(109) = RepRp(CDbl(Nul(rsHlp!Off7)))
236                 gsSql1(110) = "OFF8"
237                 gsSql2(110) = RepRp(CDbl(Nul(rsHlp!Off8)))
                    
238                 gsSql1(111) = "KLPKERJA1"
239                 gsSql2(111) = Nul(rsHlp!KLPKERJA1)
240                 gsSql1(112) = "KLPKERJA2"
241                 gsSql2(112) = Nul(rsHlp!KLPKERJA2)
242                 gsSql1(113) = "KLPKERJA3"
243                 gsSql2(113) = Nul(rsHlp!KLPKERJA3)
244                 gsSql1(114) = "KLPKERJA4"
245                 gsSql2(114) = Nul(rsHlp!KLPKERJA4)
246                 gsSql1(115) = "KLPKERJA5"
247                 gsSql2(115) = Nul(rsHlp!KLPKERJA5)
                    
248                 gsSql1(116) = "KLPKERJA6"
249                 gsSql2(116) = Nul(rsHlp!KLPKERJA6)
250                 gsSql1(117) = "KLPKERJA7"
251                 gsSql2(117) = Nul(rsHlp!KLPKERJA7)
252                 gsSql1(118) = "KLPKERJA8"
253                 gsSql2(118) = Nul(rsHlp!KLPKERJA8)
254                 gsSql1(119) = "FS_KREDIT"
255                 gsSql2(119) = Nul(rsHlp!FS_KREDIT)
256                 gsSql1(120) = "FN_PAJAK"
257                 gsSql2(120) = Nul(rsHlp!fn_pajak)
                    
258                 gsSql1(121) = "FS_TH_MOBIL"
259                 gsSql2(121) = Nul(rsHlp!FS_TH_MOBIL)
260                 gsSql1(122) = "FN_RETURN"
261                 gsSql2(122) = Nul(rsHlp!FN_RETURN)
262                 gsSql1(123) = "FS_CATATAN"
263                 gsSql2(123) = Nul(rsHlp!FS_CATATAN)
264                 gsSql1(124) = "FS_LOG_ADD"
265                 gsSql2(124) = Nul(rsHlp!FS_LOG_ADD)
266                 gsSql1(125) = "FS_LOG_DEL"
267                 gsSql2(125) = Nul(rsHlp!FS_LOG_DEL)
                    
268                 gsSql1(126) = "FN_JML_ORDER"
269                 gsSql2(126) = Nul(rsHlp!FN_JML_ORDER)
270                 gsSql1(127) = "FN_JML_RATE"
271                 gsSql2(127) = Nul(rsHlp!FN_JML_RATE)
272                 gsSql1(128) = "ALASAN_INV"
273                 gsSql2(128) = Nul(rsHlp!ALASAN_INV)
274                 gsSql1(129) = "FS_RC"
275                 gsSql2(129) = Nul(rsHlp!FS_RC)
276                 gsSql1(130) = "FS_KD_PLG"
277                 gsSql2(130) = Nul(rsHlp!FS_KD_PLG)
                    
278                 gsSql1(131) = "FS_KD_WIL"
279                 gsSql2(131) = Nul(rsHlp!FS_KD_WIL)
280                 gsSql1(132) = "FD_TGL_PART"
281                 gsSql2(132) = Format(Nul(rsHlp!FD_TGL_PART), "MM/DD/YYYY")
282                 gsSql1(133) = "FS_NO_RETUR"
283                 gsSql2(133) = Nul(rsHlp!FS_NO_RETUR)
284                 gsSql1(134) = "FD_RETUR"
285                 gsSql2(134) = Format(Nul(rsHlp!FD_RETUR), "MM/DD/YYYY")
286                 gsSql1(135) = "FS_KD_RETUR"
287                 gsSql2(135) = Nul(rsHlp!FS_KD_RETUR)
                    
288                 gsSql1(136) = "TPSVC"
289                 gsSql2(136) = Nul(rsHlp!TPSVC)
290                 gsSql1(137) = "THS"
291                 gsSql2(137) = Nul(rsHlp!THS)
292                 gsSql1(138) = "pds"
293                 gsSql2(138) = Nul(rsHlp!pds)
294                 gsSql1(139) = "MOU"
295                 gsSql2(139) = Nul(rsHlp!MOU)
296                 gsSql1(140) = "KODETIPE"
297                 gsSql2(140) = Nul(rsHlp!KODETIPE)
                    
298                 gsSql1(141) = "CMPL1"
299                 gsSql2(141) = Nul(rsHlp!CMPL1)
300                 gsSql1(142) = "CMPL2"
301                 gsSql2(142) = Nul(rsHlp!CMPL2)
302                 gsSql1(143) = "CMPL3"
303                 gsSql2(143) = Nul(rsHlp!CMPL3)
304                 gsSql1(144) = "CMPL4"
305                 gsSql2(144) = Nul(rsHlp!CMPL4)
306                 gsSql1(145) = "CMPL5"
307                 gsSql2(145) = Nul(rsHlp!CMPL5)
                    
308                 gsSql1(146) = "INVTIME"
309                 gsSql2(146) = Nul(rsHlp!INVTIME)
310                 gsSql1(147) = "DISC1"
311                 gsSql2(147) = Replace(Nul(rsHlp!DISC1), ",", ".")
312                 gsSql1(148) = "DISC2"
313                 gsSql2(148) = Replace(Nul(rsHlp!DISC2), ",", ".")
314                 gsSql1(149) = "DISC3"
315                 gsSql2(149) = Replace(Nul(rsHlp!DISC3), ",", ".")
316                 gsSql1(150) = "DISC4"
317                 gsSql2(150) = Replace(Nul(rsHlp!DISC4), ",", ".")
                    
318                 gsSql1(151) = "DISC5"
319                 gsSql2(151) = Replace(Nul(rsHlp!DISC5), ",", ".")
320                 gsSql1(152) = "DISC6"
321                 gsSql2(152) = Replace(Nul(rsHlp!DISC6), ",", ".")
322                 gsSql1(153) = "DISC7"
323                 gsSql2(153) = Replace(Nul(rsHlp!DISC7), ",", ".")
324                 gsSql1(154) = "DISC8"
325                 gsSql2(154) = Replace(Nul(rsHlp!DISC8), ",", ".")
326                 gsSql1(155) = "NoFaktur"
327                 gsSql2(155) = Nul(rsHlp!NoFaktur)
                    
328                 gsSql1(156) = "VALID1"
329                 gsSql2(156) = Nul(rsHlp!VALID1)
330                 gsSql1(157) = "VALID2"
331                 gsSql2(157) = Nul(rsHlp!VALID2)
332                 gsSql1(158) = "VALID3"
333                 gsSql2(158) = Nul(rsHlp!VALID3)
334                 gsSql1(159) = "VALID4"
335                 gsSql2(159) = Nul(rsHlp!VALID4)
336                 gsSql1(160) = "VALID5"
337                 gsSql2(160) = Nul(rsHlp!VALID5)
                    
338                 gsSql1(161) = "VALID6"
339                 gsSql2(161) = Nul(rsHlp!VALID6)
340                 gsSql1(162) = "VALID7"
341                 gsSql2(162) = Nul(rsHlp!VALID7)
342                 gsSql1(163) = "VALID8"
343                 gsSql2(163) = Nul(rsHlp!VALID8)
344                 gsSql1(164) = "FREEC1"
345                 gsSql2(164) = Nul(rsHlp!FREEC1)
346                 gsSql1(165) = "FREEC2"
347                 gsSql2(165) = Nul(rsHlp!FREEC2)
                    
348                 gsSql1(166) = "FREEC3"
349                 gsSql2(166) = Nul(rsHlp!FREEC3)
350                 gsSql1(167) = "FREEC4"
351                 gsSql2(167) = Nul(rsHlp!FREEC4)
352                 gsSql1(168) = "FREEC5"
353                 gsSql2(168) = Nul(rsHlp!FREEC5)
354                 gsSql1(169) = "FREEC6"
355                 gsSql2(169) = Nul(rsHlp!FREEC6)
356                 gsSql1(170) = "FREEC7"
357                 gsSql2(170) = Nul(rsHlp!FREEC7)
                    
358                 gsSql1(171) = "FREEC8"
359                 gsSql2(171) = Nul(rsHlp!FREEC8)
360                 gsSql1(172) = "noFP"
361                 gsSql2(172) = Nul(rsHlp!noFP)
362                 gsSql1(173) = "NOMB"
363                 gsSql2(173) = Nul(rsHlp!NOMB)
364                 gsSql1(174) = "TGLKASIR"
365                 gsSql2(174) = Format(Nul(rsHlp!TGLKASIR), "MM/DD/YYYY")
366                 gsSql1(175) = "NOFP2"
367                 gsSql2(175) = Nul(rsHlp!NOFP2)
                    
368                 gsSql1(176) = "KASIRADD"
369                 gsSql2(176) = Nul(rsHlp!KASIRADD)
370                 gsSql1(177) = "CASHADD"
371                 gsSql2(177) = Nul(rsHlp!CASHADD)
372                 gsSql1(178) = "TGLCASH"
373                 gsSql2(178) = Format(Nul(rsHlp!TGLCASH), "MM/DD/YYYY")
374                 gsSql1(179) = "CABANG"
375                 gsSql2(179) = Log.LogBengkel

376                 sSQL = GensSql2Layar("T_INVOICENEW", False)
377                 Call objO.InsertDataHO(sSQL)
                    
378                 rsHlp.MoveNext
379             Loop
380             Call objO.Update_LOG(Trim(rsT.Fields(1).value), (Trim(rsT.Fields(2).value)), Trim(rsT.Fields(0).value))
381         Else
                '---============= JIKA NO WO BELUM DIINVOICE =========---------
382             Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA WHERE NO_WO = '" & Trim(rsT.Fields(1).value) & "'")
383             If Not rsHlp.EOF Then
384                 rsHlp.MoveFirst
385                 Set rsHlp2 = objO.FetchDataHO("SELECT * FROM JOBDTA WHERE NO_WO = '" & Trim(rsT.Fields(1).value) & "'")
386                 If Not rsHlp2.EOF Then
387                     Call objO.DeleteDataHO("DELETE FROM JOBDTA WHERE NO_WO = '" & Trim(rsT.Fields(1).value) & "' ")
388                 End If
389                 Do While Not rsHlp.EOF
390                     sSQL = "INSERT INTO[JOBDTA]([CUSTOMER],[CUST_ADRES],[KOTA],[TELPON],[NO_POLISI],[NO_ENGINE],[NO_CHASSIS],[MDL],[MODEL],[FMDL],[COLOR],[FS_KD_DEALER],[FS_NM_DEALER],[DDATE]"
391                     sSQL = sSQL & ",[DOTAM],[CODE],[NPWP],[PPN],[DISC],[DISCT],[CRDTCD],[FRONT],[GROUP],[MECH],[KM],[DATE],[TIME],[TIME_M],[TIME_S],[TIME_A],[TIME_AT],[TIME_SA],[PART],[MTRL],[OIL]"
392                     sSQL = sSQL & ",[MET],[SBLT],[LABOR],[BODY],[TOTAL],[JOB],[REQ1],[REQ2],[REQ3],[REQ4],[REQ5],[JOBCD],[CR],[REP1],[REP2],[REP3],[REP4],[REP5],[REP6],[REP7],[REP8],[SLIP],[RTE1]"
393                     sSQL = sSQL & ",[RTE2],[RTE3],[RTE4],[RTE5],[RTE6],[RTE7],[RTE8],[COND1],[COND2],[COND3],[COND4],[COND5],[COND6],[COND7],[INVDATE],[NO_WO],[TPSVC],[NOINV],[STATUS],[PAID],[NOFAK]"
394                     sSQL = sSQL & ",[REMARK1],[SUGGEST],[CATAT],[DELIV],[SPK],[RPR_COND],[KD_K],[MEKANIK1],[MEKANIK2],[MEKANIK3],[MEKANIK4],[MEKANIK5],[MEKANIK6],[MEKANIK7],[MEKANIK8],[GRP1],[GRP2]"
395                     sSQL = sSQL & ",[GRP3],[GRP4],[GRP5],[GRP6],[GRP7],[GRP8],[ACTRATE1],[ACTRATE2],[ACTRATE3],[ACTRATE4],[ACTRATE5],[ACTRATE6],[ACTRATE7],[ACTRATE8],[OFF1],[OFF2],[OFF3],[OFF4],[OFF5]"
396                     sSQL = sSQL & ",[OFF6],[OFF7],[OFF8],[KLPKERJA1],[KLPKERJA2],[KLPKERJA3],[KLPKERJA4],[KLPKERJA5],[KLPKERJA6],[KLPKERJA7],[KLPKERJA8],[ALASAN],[FS_STATUS_INV],[FS_TH_MOBIL],[FN_RETURN]"
397                     sSQL = sSQL & ",[FS_LOG_ADD],[FS_LOG_DEL],[FD_TGL_DEL],[FS_RC],[FS_NM_HOBI],[FS_KD_PLG],[FS_KD_WIL],[KODETIPE],[PDS],[THS],[SBE]"
398                     sSQL = sSQL & ",[ANCUST],[ALASAN_GATE],[MOU],[TELPANCUST],[NOFK],[TGLFK],[DBASEFK],[NOFKRETUR],[TGLFKRETUR],[DBASEFKRETUR])"
399                     sSQL = sSQL & " VALUES"
400                     sSQL = sSQL & "('" & Trim(rsHlp!customer) & "','" & Trim(rsHlp!cust_adres) & "','" & Trim(rsHlp!Kota) & "'"
401                     sSQL = sSQL & ",'" & Trim(rsHlp!Telpon) & "'"
402                     sSQL = sSQL & ",'" & Trim(rsHlp!NO_POLISI) & "'"
403                     sSQL = sSQL & ",'" & Trim(rsHlp!No_Engine) & "'"
404                     sSQL = sSQL & ",'" & Trim(rsHlp!NO_CHASSIS) & "'"
405                     sSQL = sSQL & ",'" & Trim(rsHlp!MDL) & "'"
406                     sSQL = sSQL & ",'" & Trim(rsHlp!model) & "'"
407                     sSQL = sSQL & ",'" & Trim(rsHlp!Fmdl) & "'"
408                     sSQL = sSQL & ",'" & Trim(rsHlp!Color) & "'"
409                     sSQL = sSQL & ",'" & Trim(rsHlp!FS_KD_DEALER) & "'"
410                     sSQL = sSQL & ",'" & Trim(rsHlp!fs_nm_dealer) & "'"
411                     sSQL = sSQL & ",'" & Trim(rsHlp!dDate) & "'"
412                     sSQL = sSQL & ",'" & Trim(rsHlp!DOTAM) & "'"
413                     sSQL = sSQL & ",'" & Trim(rsHlp!CODE) & "'"
414                     sSQL = sSQL & ",'" & Trim(rsHlp!npwp) & "'"
415                     sSQL = sSQL & "," & Replace(Trim(rsHlp!PPn), ",", ".")
416                     sSQL = sSQL & "," & Replace(Trim(rsHlp!disc), ",", ".")
417                     sSQL = sSQL & "," & Replace(Trim(rsHlp!DISCT), ",", ".")
418                     sSQL = sSQL & "," & Replace(Trim(rsHlp!CRDTCD), ",", ".")
419                     sSQL = sSQL & ",'" & Trim(rsHlp!FRONT) & "'"
420                     sSQL = sSQL & ",'" & Trim(rsHlp!Group) & "'"
421                     sSQL = sSQL & ",'" & Trim(rsHlp!Mech) & "'"
422                     sSQL = sSQL & ",'" & Trim(rsHlp!km) & "'"
423                     sSQL = sSQL & ",'" & Format(Trim(rsHlp!Date), "MM/DD/YYYY") & "'"
424                     sSQL = sSQL & ",'" & Trim(rsHlp!Time) & "'"
425                     sSQL = sSQL & ",'" & Trim(rsHlp!TIME_M) & "'"
426                     sSQL = sSQL & ",'" & Trim(rsHlp!TIME_S) & "'"
427                     sSQL = sSQL & ",'" & Trim(rsHlp!TIME_A) & "'"
428                     sSQL = sSQL & ",'" & Trim(rsHlp!TIME_AT) & "'"
429                     sSQL = sSQL & ",'" & Trim(rsHlp!TIME_SA) & "'"
430                     sSQL = sSQL & "," & Replace(Trim(rsHlp!part), ",", ".")
431                     sSQL = sSQL & "," & Replace(Trim(rsHlp!MTRL), ",", ".")
432                     sSQL = sSQL & "," & Replace(Trim(rsHlp!OIL), ",", ".")
433                     sSQL = sSQL & "," & Replace(Trim(rsHlp!met), ",", ".")
434                     sSQL = sSQL & "," & Replace(Trim(rsHlp!sblt), ",", ".")
435                     sSQL = sSQL & "," & Replace(Trim(rsHlp!labor), ",", ".")
436                     sSQL = sSQL & "," & Replace(Trim(rsHlp!BODY), ",", ".")
437                     sSQL = sSQL & "," & Replace(Trim(rsHlp!total), ",", ".")
438                     sSQL = sSQL & ",'" & Trim(rsHlp!job) & "'"
439                     sSQL = sSQL & ",'" & Trim(rsHlp!REQ1) & "'"
440                     sSQL = sSQL & ",'" & Trim(rsHlp!REQ2) & "'"
441                     sSQL = sSQL & ",'" & Trim(rsHlp!REQ3) & "'"
442                     sSQL = sSQL & ",'" & Trim(rsHlp!REQ4) & "'"
443                     sSQL = sSQL & ",'" & Trim(rsHlp!REQ5) & "'"
444                     sSQL = sSQL & ",'" & Trim(rsHlp!JobCD) & "'"
445                     sSQL = sSQL & ",'" & Trim(rsHlp!CR) & "'"
446                     sSQL = sSQL & ",'" & Trim(rsHlp!rep1) & "'"
447                     sSQL = sSQL & ",'" & Trim(rsHlp!REP2) & "'"
448                     sSQL = sSQL & ",'" & Trim(rsHlp!REP3) & "'"
449                     sSQL = sSQL & ",'" & Trim(rsHlp!REP4) & "'"
450                     sSQL = sSQL & ",'" & Trim(rsHlp!REP5) & "'"
451                     sSQL = sSQL & ",'" & Trim(rsHlp!REP6) & "'"
452                     sSQL = sSQL & ",'" & Trim(rsHlp!REP7) & "'"
453                     sSQL = sSQL & ",'" & Trim(rsHlp!REP8) & "'"
454                     sSQL = sSQL & ",'" & Trim(rsHlp!Slip) & "'"
455                     sSQL = sSQL & "," & Replace(Trim(rsHlp!RTE1), ",", ".")
456                     sSQL = sSQL & "," & Replace(Trim(rsHlp!RTE2), ",", ".")
457                     sSQL = sSQL & "," & Replace(Trim(rsHlp!RTE3), ",", ".")
458                     sSQL = sSQL & "," & Replace(Trim(rsHlp!RTE4), ",", ".")
459                     sSQL = sSQL & "," & Replace(Trim(rsHlp!RTE5), ",", ".")
460                     sSQL = sSQL & "," & Replace(Trim(rsHlp!RTE6), ",", ".")
461                     sSQL = sSQL & "," & Replace(Trim(rsHlp!RTE7), ",", ".")
462                     sSQL = sSQL & "," & Replace(Trim(rsHlp!RTE8), ",", ".")
463                     sSQL = sSQL & ",'" & Trim(rsHlp!COND1) & "'"
464                     sSQL = sSQL & ",'" & Trim(rsHlp!COND2) & "'"
465                     sSQL = sSQL & ",'" & Trim(rsHlp!COND3) & "'"
466                     sSQL = sSQL & ",'" & Trim(rsHlp!COND4) & "'"
467                     sSQL = sSQL & ",'" & Trim(rsHlp!COND5) & "'"
468                     sSQL = sSQL & ",'" & Trim(rsHlp!COND6) & "'"
469                     sSQL = sSQL & ",'" & Trim(rsHlp!COND7) & "'"
470                     sSQL = sSQL & ",'" & Format(Trim(rsHlp!InvDate), "mm/dd/yyyy") & "'"
471                     sSQL = sSQL & ",'" & Trim(rsHlp!no_wo) & "'"
472                     sSQL = sSQL & ",'" & Trim(rsHlp!TPSVC) & "'"
473                     sSQL = sSQL & ",'" & Trim(rsHlp!noInv) & "'"
474                     sSQL = sSQL & ",'" & Trim(rsHlp!Status) & "'"
475                     sSQL = sSQL & ",'" & Trim(rsHlp!paid) & "'"
476                     sSQL = sSQL & ",'" & Trim(rsHlp!NoFak) & "'"
477                     sSQL = sSQL & ",'" & Trim(rsHlp!REMARK1) & "'"
478                     sSQL = sSQL & ",'" & Trim(rsHlp!SUGGEST) & "'"
479                     sSQL = sSQL & ",'" & Trim(rsHlp!CATAT) & "'"
480                     sSQL = sSQL & ",'" & Trim(rsHlp!DELIV) & "'"
481                     sSQL = sSQL & ",'" & Trim(rsHlp!SPK) & "'"
482                     sSQL = sSQL & ",'" & Trim(rsHlp!RPR_COND) & "'"
483                     sSQL = sSQL & ",'" & Trim(rsHlp!KD_K) & "'"
484                     sSQL = sSQL & ",'" & Trim(rsHlp!mekanik1) & "'"
485                     sSQL = sSQL & ",'" & Trim(rsHlp!mekanik2) & "'"
486                     sSQL = sSQL & ",'" & Trim(rsHlp!MEKANIK3) & "'"
487                     sSQL = sSQL & ",'" & Trim(rsHlp!MEKANIK4) & "'"
488                     sSQL = sSQL & ",'" & Trim(rsHlp!MEKANIK5) & "'"
489                     sSQL = sSQL & ",'" & Trim(rsHlp!MEKANIK6) & "'"
490                     sSQL = sSQL & ",'" & Trim(rsHlp!MEKANIK7) & "'"
491                     sSQL = sSQL & ",'" & Trim(rsHlp!MEKANIK8) & "'"
492                     sSQL = sSQL & ",'" & Trim(rsHlp!GRP1) & "'"
493                     sSQL = sSQL & ",'" & Trim(rsHlp!GRP2) & "'"
494                     sSQL = sSQL & ",'" & Trim(rsHlp!GRP3) & "'"
495                     sSQL = sSQL & ",'" & Trim(rsHlp!GRP4) & "'"
496                     sSQL = sSQL & ",'" & Trim(rsHlp!GRP5) & "'"
497                     sSQL = sSQL & ",'" & Trim(rsHlp!GRP6) & "'"
498                     sSQL = sSQL & ",'" & Trim(rsHlp!GRP7) & "'"
499                     sSQL = sSQL & ",'" & Trim(rsHlp!GRP8) & "'"
500                     sSQL = sSQL & "," & Replace(Trim(rsHlp!ACTRATE1), ",", ".")
501                     sSQL = sSQL & "," & Replace(Trim(rsHlp!ACTRATE2), ",", ".")
502                     sSQL = sSQL & "," & Replace(Trim(rsHlp!ACTRATE3), ",", ".")
503                     sSQL = sSQL & "," & Replace(Trim(rsHlp!ACTRATE4), ",", ".")
504                     sSQL = sSQL & "," & Replace(Trim(rsHlp!ACTRATE5), ",", ".")
505                     sSQL = sSQL & "," & Replace(Trim(rsHlp!ACTRATE6), ",", ".")
506                     sSQL = sSQL & "," & Replace(Trim(rsHlp!ACTRATE7), ",", ".")
507                     sSQL = sSQL & "," & Replace(Trim(rsHlp!ACTRATE8), ",", ".")
508                     sSQL = sSQL & "," & Replace(Trim(rsHlp!Off1), ",", ".")
509                     sSQL = sSQL & "," & Replace(Trim(rsHlp!Off2), ",", ".")
510                     sSQL = sSQL & "," & Replace(Trim(rsHlp!Off3), ",", ".")
511                     sSQL = sSQL & "," & Replace(Trim(rsHlp!Off4), ",", ".")
512                     sSQL = sSQL & "," & Replace(Trim(rsHlp!Off5), ",", ".")
513                     sSQL = sSQL & "," & Replace(Trim(rsHlp!Off6), ",", ".")
514                     sSQL = sSQL & "," & Replace(Trim(rsHlp!Off7), ",", ".")
515                     sSQL = sSQL & "," & Replace(Trim(rsHlp!Off8), ",", ".")
516                     sSQL = sSQL & ",'" & Trim(rsHlp!KLPKERJA1) & "'"
517                     sSQL = sSQL & ",'" & Trim(rsHlp!KLPKERJA2) & "'"
518                     sSQL = sSQL & ",'" & Trim(rsHlp!KLPKERJA3) & "'"
519                     sSQL = sSQL & ",'" & Trim(rsHlp!KLPKERJA4) & "'"
520                     sSQL = sSQL & ",'" & Trim(rsHlp!KLPKERJA5) & "'"
521                     sSQL = sSQL & ",'" & Trim(rsHlp!KLPKERJA6) & "'"
522                     sSQL = sSQL & ",'" & Trim(rsHlp!KLPKERJA7) & "'"
523                     sSQL = sSQL & ",'" & Trim(rsHlp!KLPKERJA8) & "'"
524                     sSQL = sSQL & ",'" & Trim(rsHlp!alasan) & "'"
525                     sSQL = sSQL & ",'" & Trim(rsHlp!FS_STATUS_INV) & "'"
526                     sSQL = sSQL & ",'" & Trim(rsHlp!FS_TH_MOBIL) & "'"
527                     sSQL = sSQL & ",'" & Trim(rsHlp!FN_RETURN) & "'"
528                     sSQL = sSQL & ",'" & Trim(rsHlp!FS_LOG_ADD) & "'"
529                     sSQL = sSQL & ",'" & Trim(rsHlp!FS_LOG_DEL) & "'"
530                     sSQL = sSQL & ",'" & Format(Trim(rsHlp!FD_TGL_DEL), "mm/dd/yyyy") & "'"
531                     sSQL = sSQL & ",'" & Trim(rsHlp!FS_RC) & "'"
532                     sSQL = sSQL & ",'" & Trim(rsHlp!FS_NM_HOBI) & "'"
533                     sSQL = sSQL & ",'" & Trim(rsHlp!FS_KD_PLG) & "'"
534                     sSQL = sSQL & ",'" & Trim(rsHlp!FS_KD_WIL) & "'"
535                     sSQL = sSQL & ",'" & Trim(rsHlp!KODETIPE) & "'"
536                     sSQL = sSQL & "," & Replace(Trim(rsHlp!pds), ",", ".")
537                     sSQL = sSQL & ",'" & Trim(rsHlp!THS) & "'"
538                     sSQL = sSQL & ",'" & Trim(rsHlp!SBE) & "'"
539                     sSQL = sSQL & ",'" & Trim(rsHlp!ANCUST) & "'"
540                     sSQL = sSQL & ",'" & Trim(rsHlp!ALASAN_GATE) & "'"
541                     sSQL = sSQL & ",'" & Trim(rsHlp!MOU) & "'"
542                     sSQL = sSQL & ",'" & Trim(rsHlp!TELPANCUST) & "'"
543                     sSQL = sSQL & ",'" & Trim(rsHlp!NOFK) & "'"
544                     sSQL = sSQL & ",'" & Format(Trim(rsHlp!TGLFK), "mm/dd/yyyy") & "'"
545                     sSQL = sSQL & ",'" & Trim(rsHlp!dBaseFK) & "'"
546                     sSQL = sSQL & ",'" & Trim(rsHlp!NOFKRETUR) & "'"
547                     sSQL = sSQL & ",'" & Format(Trim(rsHlp!TGLFKRETUR), "mm/dd/yyyy") & "'"
548                     sSQL = sSQL & ",'" & Trim(rsHlp!DBASEFKRETUR) & "')"
549                     Call objO.InsertDataHO(sSQL)
550                     Call objO.Update_LOG(Trim(rsT.Fields(1).value), (Trim(rsT.Fields(2).value)), Trim(rsT.Fields(0).value))
551                     rsHlp.MoveNext
552                 Loop
553             End If
                
                '==============================================================
554         End If
            
555     Else
556         If Trim(rsT.Fields(0).value) = "INV1" Then
557             Call objO.InsertDataHO(Replace(Trim(rsT.Fields(5).value), "~", "'"))
558             Call Transfer_INV1(Trim(rsT.Fields(1).value))
559             Call objO.Update_LOG((rsT.Fields(1).value), (Trim(rsT.Fields(2).value)), Trim(rsT.Fields(0).value))
560         ElseIf Trim(rsT.Fields(0).value) = "EXEC REPOTAF" Then
561             Call DoTransferExec(rsT)
562         ElseIf Trim(rsT.Fields(0).value) = "DETAILWAC" Then
563             If Trim(rsT.Fields(2).value) = "DELETE" Then
564                 Call objO.InsertDataHO(Replace(Trim(rsT.Fields(5).value), "~", "'"))
565                 Call objO.Update_LOG(Trim(rsT.Fields(1).value), (Trim(rsT.Fields(2).value)), Trim(rsT.Fields(0).value))

566             ElseIf Trim(rsT.Fields(2).value) = "INSERT" Then
                    
567                 If Log.detWAC = 1 Then
                       ' On Error GoTo error_simpanwac
                        Dim dbConHO As New ADODB.Connection
                        Dim rsTemp As New ADODB.Recordset
                        'errHO = False
568                     If dbConHO.State = adStateClosed Then dbConHO.Open (SQLConBengkel99)
                        
569                     dbConHO.BeginTrans
570                     isBegin = True
                        Dim rsWACHO As New ADODB.Recordset
571                     If rsWACHO.State Then rsWACHO.Close
572                     rsWACHO.Open "SELECT TOP 1 nowo,matriks,kondisi,bagian,posisi,dealer,pic,idantrian,adaFoto FROM detailwacho ", dbConHO, adOpenKeyset, adLockOptimistic
573                     Set rsTemp = objO.FetchData("SELECT no_wo,matrix,kondisi,bagian,posisi,pic,idantrian,adaFoto FROM detailwac WITH (NOLOCK) WHERE idwac = '" & Trim(rsT.Fields(5).value) & "'")
                        
574                     If Not rsTemp.EOF Then
575                         rsTemp.MoveFirst
                            ''------------------------------------------------
                            ''2017-04-04 hapus data lama
                            ''------------------------------------------------
576                         sSQL = "DELETE FROM DETAILWACHO WHERE idwac = '" & Trim(rsT.Fields(5).value) & "'"
                            'Call objHO.AddNew_WO(sSQL)
      '                      dbConHO.Execute sSQL
                            
                            ''------------------------------------------------
577                         Do While Not rsTemp.EOF
578                           If Trim(rsTemp!adaFoto) = 1 Then
579                               Set fotoStream = New ADODB.Stream
580                               With fotoStream
581                                   .Type = adTypeBinary
582                                   .Open
583                                   .Write rsTemp!PIC
584                                   .SaveToFile App.Path & "simpanWAC.bmp", adSaveCreateOverWrite
                                      
585                               End With
586                               Set readStream = New ADODB.Stream
587                               With readStream
588                                   .Type = adTypeBinary
589                                   .Open
590                                   .LoadFromFile App.Path & "simpanWAC.bmp"
591                               End With
                                            
592                               With rsWACHO
593                                   .AddNew
594                                   .Fields("nowo").value = Trim(rsTemp!no_wo)
595                                   .Fields("matriks").value = Nul(Trim(rsTemp!matrix))
596                                   .Fields("kondisi").value = Nul(Trim(rsTemp!kondisi))
597                                   .Fields("bagian").value = Nul(Trim(rsTemp!bagian))
598                                   .Fields("posisi").value = Nul(Trim(rsTemp!Posisi))
599                                   .Fields("dealer").value = gKdCbg
600                                   .Fields("pic").value = readStream.Read
601                                   .Fields("idantrian").value = Trim(rsTemp!idAntrian)
602                                   .Fields("adaFoto").value = 1
603                                   .Update
604                               End With
605                               Kill App.Path & "simpanWAC.bmp"
606                           Else
607                               With rsWACHO
608                                   .AddNew
609                                   .Fields("nowo").value = Trim(rsTemp!no_wo)
610                                   .Fields("matriks").value = Nul(Trim(rsTemp!matrix))
611                                   .Fields("kondisi").value = Nul(Trim(rsTemp!kondisi))
612                                   .Fields("bagian").value = Nul(Trim(rsTemp!bagian))
613                                   .Fields("posisi").value = Nul(Trim(rsTemp!Posisi))
614                                   .Fields("dealer").value = gKdCbg
615                                   .Fields("idantrian").value = Trim(rsTemp!idAntrian)
616                                   .Fields("adaFoto").value = 0
617                                   .Update
618                               End With
619                           End If
                                                                  
                                

620                             rsTemp.MoveNext
621                         Loop
622                     End If
                        
623                     dbConHO.CommitTrans
624                     Call objO.Update_LOG(Trim(rsT.Fields(1).value), (Trim(rsT.Fields(2).value)), Trim(rsT.Fields(0).value))
625                 End If
                    
626             End If
627         Else
                'Call objHT.InsertData_Ho_Only(Replace(Trim(rsT.Fields(5).value), "~", "'"))
628             Call objO.InsertDataHO(Replace(Trim(rsT.Fields(5).value), "~", "'"))
629             Call objO.Update_LOG(Trim(rsT.Fields(1).value), (Trim(rsT.Fields(2).value)), Trim(rsT.Fields(0).value))
630             If Trim(rsT.Fields(0).value) = "CLAIM" Then
631                 Call objO.InsertData("UPDATE CLAIMWARRANTY SET TRANSFER = '1' WHERE NO_WO = '" & Trim(rsT.Fields(1).value) & "'")
632             End If
633         End If
634     End If
635     rsT.MoveNext
636       Loop

637      On Error GoTo 0
638      Exit Sub

ProsesTransfer_HO_Error:

639       If isBegin Then
640           dbConHO.RollbackTrans
641       End If

642       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure ProsesTransfer_HO of Form Layar")
End Sub

Private Sub Transfer_INV1(nWO As String)
          Dim rsPr As ADODB.Recordset
1        On Error GoTo Transfer_INV1_Error

2         Set rsPr = objO.FetchData("Select * FROm T_Part WITH (NOLOCK) Where No_WO = '" & nWO & "'")
3         If rsPr.RecordCount <> 0 Then
4             rsPr.MoveFirst
5             Do Until rsPr.EOF
6                 Call objO.InsertDataHO("Insert Into T_PARTnew (NO_WO,KODEITEM,KETERANGAN,QTY,HARGA,OLI,NOINV,FN_EXT_DISKON,FS_LOG_DEL,FD_TGL_DEL,CABANG) Values ('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "'," & Val(rsPr.Fields(3).value) & "," & RepRp(rsPr.Fields(4).value) & "," & IIf(rsPr.Fields(5).value = False, 0, 1) & ",'" & Trim(rsPr.Fields(6).value) & "'," & Val(rsPr.Fields(7).value) & ",'','01/01/3000','" & Log.LogBengkel & "')")
7                 rsPr.MoveNext
8             Loop
9         End If
          
10        Set rsPr = objO.FetchData("Select * FROm Hist_Sublet WITH (NOLOCK) Where No_WO = '" & nWO & "'")
11        If rsPr.RecordCount <> 0 Then
12            rsPr.MoveFirst
13            Do Until rsPr.EOF
14                Call objO.InsertDataHO("Insert Into Hist_SUBLETnew (NOPO,NO_WO,NOINV,NO_CHASSIS,NOFAK,KODEITEM,NAMABARANG,KODESATUAN,QTY,HARGA,NONOTA,KETERANGAN,FLAG,SUPPLIER,TAGIH,CABANG) Values ('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "','" & Trim(rsPr.Fields(5).value) & "','" & Trim(rsPr.Fields(6).value) & "','" & Trim(rsPr.Fields(7).value) & "'," & Val(rsPr.Fields(8).value) & "," & RepRp(rsPr.Fields(9).value) & ",'" & Trim(rsPr.Fields(10).value) & "','" & Trim(rsPr.Fields(11).value) & "','" & Trim(rsPr.Fields(12).value) & "','" & Trim(rsPr.Fields(13).value) & "','" & Trim(rsPr.Fields(14).value) & "','" & Log.LogBengkel & "')")
15                rsPr.MoveNext
16            Loop
17        End If
          
18        Set rsPr = objO.FetchData("Select * FROm History WITH (NOLOCK) Where No_WO = '" & nWO & "'")
19        If rsPr.RecordCount <> 0 Then
20            rsPr.MoveFirst
21            Do Until rsPr.EOF
22                Call objO.InsertDataHO("Insert Into Historynew (CUSTOMER,CUST_ADRES,KOTA,NO_POLISI,NO_ENGINE,NO_CHASSIS,FRONT,MECH,KM,[DATE],PART_G,PART_NG,OIL_G,OIL_NG,MTRL,LABOR,BODY,SBLT,MET,TOTAL,JOB,REP1,REP2,REP3,REP4,REP5,REP6,REP7,MEKANIK1,MEKANIK2,MEKANIK3,MEKANIK4,MEKANIK5,MEKANIK6,MEKANIK7,MEKANIK8,KLPKERJA1,KLPKERJA2,KLPKERJA3,KLPKERJA4,KLPKERJA5,KLPKERJA6,KLPKERJA7,KLPKERJA8,NO_WO,NOINV,NOFAK,[STATUS],PAID,DISC,DISCT,REMARK1,FN_RETURN,FS_RC,FS_KD_DEALER,FS_NM_DEALER,TIPEWO,KODETIPE,SBE,FS_KD_PLG,FS_KD_WIL,THS,MOU,PDS,FS_LOG_DEL,FD_TGL_DEL,ALASAN_DEL,SUGGEST,INVDATE,CABANG) Values " & _
                  "('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "','" & Trim(rsPr.Fields(5).value) & "', " & _
                  "'" & Trim(rsPr.Fields(6).value) & "','" & Trim(rsPr.Fields(7).value) & "','" & Trim(rsPr.Fields(8).value) & "','" & TglNul2(rsPr.Fields(9).value) & "'," & RepRp(rsPr.Fields(10).value) & "," & RepRp(rsPr.Fields(11).value) & ", " & _
                  " " & RepRp(rsPr.Fields(12).value) & "," & RepRp(rsPr.Fields(13).value) & "," & RepRp(rsPr.Fields(14).value) & "," & RepRp(rsPr.Fields(15).value) & "," & RepRp(rsPr.Fields(16).value) & "," & RepRp(rsPr.Fields(17).value) & " , " & _
                  " " & RepRp(rsPr.Fields(18).value) & "," & RepRp(rsPr.Fields(19).value) & ",'" & Trim(rsPr.Fields(20).value) & "','" & Trim(rsPr.Fields(21).value) & "','" & Trim(rsPr.Fields(22).value) & "','" & Trim(rsPr.Fields(23).value) & "', " & _
                  "'" & Trim(rsPr.Fields(24).value) & "','" & Trim(rsPr.Fields(25).value) & "','" & Trim(rsPr.Fields(26).value) & "','" & Trim(rsPr.Fields(27).value) & "','" & Trim(rsPr.Fields(28).value) & "','" & Trim(rsPr.Fields(29).value) & "', " & _
                  "'" & Trim(rsPr.Fields(30).value) & "','" & Trim(rsPr.Fields(31).value) & "','" & Trim(rsPr.Fields(32).value) & "','" & Trim(rsPr.Fields(33).value) & "','" & Trim(rsPr.Fields(34).value) & "','" & Trim(rsPr.Fields(35).value) & "', " & _
                  "'" & Trim(rsPr.Fields(36).value) & "','" & Trim(rsPr.Fields(37).value) & "','" & Trim(rsPr.Fields(38).value) & "','" & Trim(rsPr.Fields(39).value) & "','" & Trim(rsPr.Fields(40).value) & "','" & Trim(rsPr.Fields(41).value) & "', " & _
                  "'" & Trim(rsPr.Fields(42).value) & "','" & Trim(rsPr.Fields(43).value) & "','" & Trim(rsPr.Fields(44).value) & "','" & Trim(rsPr.Fields(45).value) & "','" & Trim(rsPr.Fields(46).value) & "','" & Trim(rsPr.Fields(47).value) & "', " & _
                  "'" & Trim(rsPr.Fields(48).value) & "'," & Val(rsPr.Fields(49).value) & "," & Val(rsPr.Fields(50).value) & ",'" & Trim(rsPr.Fields(51).value) & "'," & Val(rsPr.Fields(52).value) & ",'" & Trim(rsPr.Fields(53).value) & "', " & _
                  "'" & Trim(rsPr.Fields(54).value) & "','" & Trim(rsPr.Fields(55).value) & "','" & Trim(rsPr.Fields(56).value) & "','" & Trim(rsPr.Fields(57).value) & "','" & Trim(rsPr.Fields(58).value) & "','" & Trim(rsPr.Fields(59).value) & "', " & _
                  "'" & Trim(rsPr.Fields(60).value) & "'," & Val(rsPr.Fields(61).value) & ",'" & Trim(rsPr.Fields(62).value) & "'," & Val(rsPr.Fields(63).value) & ",'" & Trim(rsPr.Fields(64).value) & "','" & Format(rsPr.Fields(65).value, "MM/DD/YYYY") & "', " & _
                  "'" & Trim(rsPr.Fields(66).value) & "','" & Trim(rsPr.Fields(67).value) & "','" & Format(rsPr.Fields(68).value, "MM/DD/YYYY") & "','" & Log.LogBengkel & "')")
                  
23                rsPr.MoveNext
24            Loop
25        End If
          'Call objHT.CloseTransactionHO
          'Call objHT.OpenTransactionHO
26        Call objO.InsertDataHO("Delete FROM JOBDTA Where No_Wo = '" & nWO & "' and FS_NM_DEALER = '" & Trim(gNmCbg) & "'")
          'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & nWO & "'")
27        Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & nWO & "'")
28        Set rsPr = objO.FetchData("Select * from JOBDTA WITH (NOLOCK) Where No_Wo = '" & nWO & "'")
29        Call objO.InsertDataHO("Insert Into JOBDTA ([CUSTOMER], [CUST_ADRES], [KOTA], [TELPON], [NO_POLISI], [NO_ENGINE], [NO_CHASSIS], [MDL], [MODEL], [FMDL], [COLOR], [FS_KD_DEALER], [FS_NM_DEALER], [DDATE], [DOTAM], [CODE], [NPWP], [PPN], [DISC], [DISCT], [CRDTCD], [FRONT], [GROUP], [MECH] " & _
                ", [KM], [DATE], [TIME], [TIME_M], [TIME_S], [TIME_A], [TIME_AT], [TIME_SA], [PART], [MTRL], [OIL], [MET], [SBLT], [LABOR], [BODY], [TOTAL], [JOB], [REQ1], [REQ2], [REQ3], [REQ4], [REQ5], [JOBCD], [CR]" & _
                ", [REP1], [REP2], [REP3], [REP4], [REP5], [REP6], [REP7], [REP8], [SLIP], [RTE1], [RTE2], [RTE3], [RTE4], [RTE5], [RTE6], [RTE7], [RTE8], [COND1], [COND2]" & _
                ", [COND3], [COND4], [COND5], [COND6], [COND7], [INVDATE], [NO_WO], [TPSVC], [NOINV], [STATUS], [PAID], [NOFAK], [REMARK1], [SUGGEST], [CATAT], [DELIV], [SPK], [RPR_COND], [KD_K], [MEKANIK1], [MEKANIK2], [MEKANIK3], [MEKANIK4], [MEKANIK5]" & _
                ", [MEKANIK6], [MEKANIK7], [MEKANIK8], [GRP1], [GRP2], [GRP3], [GRP4], [GRP5], [GRP6], [GRP7], [GRP8], [ACTRATE1], [ACTRATE2], [ACTRATE3], [ACTRATE4], [ACTRATE5], [ACTRATE6], [ACTRATE7], [ACTRATE8], [OFF1], [OFF2], [OFF3], [OFF4], [OFF5]" & _
                ", [OFF6], [OFF7], [OFF8], [KLPKERJA1], [KLPKERJA2], [KLPKERJA3], [KLPKERJA4], [KLPKERJA5], [KLPKERJA6], [KLPKERJA7], [KLPKERJA8], [ALASAN], [FS_STATUS_INV], [FS_TH_MOBIL], [FN_RETURN], [FS_LOG_ADD], [FS_LOG_DEL], [FD_TGL_DEL], [FS_RC]" & _
                ", [FS_NM_HOBI], [FS_KD_PLG], [FS_KD_WIL], [KODETIPE], [PDS], [THS], [SBE], [ANCUST], [ALASAN_GATE], [MOU], [TELPANCUST], [NOFK], [TGLFK], [DBASEFK], [NOFKRETUR], [TGLFKRETUR], [DBASEFKRETUR]) Values " & _
                  "('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "','" & Trim(rsPr.Fields(5).value) & "','" & Trim(rsPr.Fields(6).value) & "','" & Trim(rsPr.Fields(7).value) & "','" & Trim(rsPr.Fields(8).value) & "','" & Trim(rsPr.Fields(9).value) & "','" & Trim(rsPr.Fields(10).value) & "','" & Trim(rsPr.Fields(11).value) & "', " & _
                  "'" & Trim(rsPr.Fields(12).value) & "','" & Trim(rsPr.Fields(13).value) & "','" & Trim(rsPr.Fields(14).value) & "','" & Trim(rsPr.Fields(15).value) & "','" & Trim(rsPr.Fields(16).value) & "'," & RepRp(rsPr.Fields(17).value) & "," & RepRp(rsPr.Fields(18).value) & "," & RepRp(rsPr.Fields(19).value) & "," & RepRp(rsPr.Fields(20).value) & ",'" & Trim(rsPr.Fields(21).value) & "','" & Trim(rsPr.Fields(22).value) & "','" & Trim(rsPr.Fields(23).value) & "', " & _
                  "'" & Trim(rsPr.Fields(24).value) & "','" & TglNul2(rsPr.Fields(25).value) & "','" & Trim(rsPr.Fields(26).value) & "','" & Trim(rsPr.Fields(27).value) & "','" & Trim(rsPr.Fields(28).value) & "','" & Trim(rsPr.Fields(29).value) & "','" & Trim(rsPr.Fields(30).value) & "','" & Trim(rsPr.Fields(31).value) & "'," & RepRp(rsPr.Fields(32).value) & "," & RepRp(rsPr.Fields(33).value) & "," & RepRp(rsPr.Fields(34).value) & "," & RepRp(rsPr.Fields(35).value) & ", " & _
                  " " & RepRp(rsPr.Fields(36).value) & "," & RepRp(rsPr.Fields(37).value) & "," & RepRp(rsPr.Fields(38).value) & "," & RepRp(rsPr.Fields(39).value) & ",'" & Trim(rsPr.Fields(40).value) & "','" & Trim(rsPr.Fields(41).value) & "','" & Trim(rsPr.Fields(42).value) & "','" & Trim(rsPr.Fields(43).value) & "','" & Trim(rsPr.Fields(44).value) & "','" & Trim(rsPr.Fields(45).value) & "','" & Trim(rsPr.Fields(46).value) & "','" & Trim(rsPr.Fields(47).value) & "', " & _
                  "'" & Trim(rsPr.Fields(48).value) & "','" & Trim(rsPr.Fields(49).value) & "','" & Trim(rsPr.Fields(50).value) & "','" & Trim(rsPr.Fields(51).value) & "','" & Trim(rsPr.Fields(52).value) & "','" & Trim(rsPr.Fields(53).value) & "','" & Trim(rsPr.Fields(54).value) & "','" & Trim(rsPr.Fields(55).value) & "','" & Trim(rsPr.Fields(56).value) & "'," & RepRp(rsPr.Fields(57).value) & "," & RepRp(rsPr.Fields(58).value) & "," & RepRp(rsPr.Fields(59).value) & ", " & _
                  " " & RepRp(rsPr.Fields(60).value) & "," & RepRp(rsPr.Fields(61).value) & "," & RepRp(rsPr.Fields(62).value) & "," & RepRp(rsPr.Fields(63).value) & "," & RepRp(rsPr.Fields(64).value) & ",'" & Trim(rsPr.Fields(65).value) & "','" & Trim(rsPr.Fields(66).value) & "','" & Trim(rsPr.Fields(67).value) & "','" & Trim(rsPr.Fields(68).value) & "','" & Trim(rsPr.Fields(69).value) & "','" & Trim(rsPr.Fields(70).value) & "','" & Trim(rsPr.Fields(71).value) & "', " & _
                  "'" & TglNul2(rsPr.Fields(72).value) & "','" & Trim(rsPr.Fields(73).value) & "','" & Trim(rsPr.Fields(74).value) & "','" & Trim(rsPr.Fields(75).value) & "','" & Trim(rsPr.Fields(76).value) & "','" & Trim(rsPr.Fields(77).value) & "','" & Trim(rsPr.Fields(78).value) & "','" & Trim(rsPr.Fields(79).value) & "','" & Trim(rsPr.Fields(80).value) & "','" & Trim(rsPr.Fields(81).value) & "','" & Trim(rsPr.Fields(82).value) & "','" & Trim(rsPr.Fields(83).value) & "', " & _
                  "'" & Trim(rsPr.Fields(84).value) & "','" & Trim(rsPr.Fields(85).value) & "','" & Trim(rsPr.Fields(86).value) & "','" & Trim(rsPr.Fields(87).value) & "','" & Trim(rsPr.Fields(88).value) & "','" & Trim(rsPr.Fields(89).value) & "','" & Trim(rsPr.Fields(90).value) & "','" & Trim(rsPr.Fields(91).value) & "','" & Trim(rsPr.Fields(92).value) & "','" & Trim(rsPr.Fields(93).value) & "','" & Trim(rsPr.Fields(94).value) & "','" & Trim(rsPr.Fields(95).value) & "', " & _
                  "'" & Trim(rsPr.Fields(96).value) & "','" & Trim(rsPr.Fields(97).value) & "','" & Trim(rsPr.Fields(98).value) & "','" & Trim(rsPr.Fields(99).value) & "','" & Trim(rsPr.Fields(100).value) & "','" & Trim(rsPr.Fields(101).value) & "'," & RepRp(rsPr.Fields(102).value) & "," & RepRp(rsPr.Fields(103).value) & "," & RepRp(rsPr.Fields(104).value) & "," & RepRp(rsPr.Fields(105).value) & "," & RepRp(rsPr.Fields(106).value) & "," & RepRp(rsPr.Fields(107).value) & ", " & _
                  " " & RepRp(rsPr.Fields(108).value) & "," & RepRp(rsPr.Fields(109).value) & "," & RepRp(rsPr.Fields(110).value) & "," & RepRp(rsPr.Fields(111).value) & "," & RepRp(rsPr.Fields(112).value) & "," & RepRp(rsPr.Fields(113).value) & "," & RepRp(rsPr.Fields(114).value) & "," & RepRp(rsPr.Fields(115).value) & "," & RepRp(rsPr.Fields(116).value) & "," & RepRp(rsPr.Fields(117).value) & ",'" & Trim(rsPr.Fields(118).value) & "','" & Trim(rsPr.Fields(119).value) & "', " & _
                  "'" & Trim(rsPr.Fields(120).value) & "','" & Trim(rsPr.Fields(121).value) & "','" & Trim(rsPr.Fields(122).value) & "','" & Trim(rsPr.Fields(123).value) & "','" & Trim(rsPr.Fields(124).value) & "','" & Trim(rsPr.Fields(125).value) & "','" & Trim(rsPr.Fields(126).value) & "','" & Trim(rsPr.Fields(127).value) & "','" & Trim(rsPr.Fields(128).value) & "'," & Val(rsPr.Fields(129).value) & ",'" & Trim(rsPr.Fields(130).value) & "','" & Trim(rsPr.Fields(131).value) & "', " & _
                  "'" & TglNul2(rsPr.Fields(132).value) & "','" & Trim(rsPr.Fields(133).value) & "','" & Trim(rsPr.Fields(134).value) & "','" & Trim(rsPr.Fields(135).value) & "','" & Trim(rsPr.Fields(136).value) & "','" & Trim(rsPr.Fields(137).value) & "'," & RepRp(rsPr.Fields(138).value) & ",'" & Trim(rsPr.Fields(139).value) & "','" & Trim(rsPr.Fields(140).value) & "','" & Trim(rsPr.Fields(141).value) & "','" & Trim(rsPr.Fields(142).value) & "','" & Trim(rsPr.Fields(143).value) & "', " & _
                  "'" & Trim(rsPr.Fields(144).value) & "','" & Trim(rsPr.Fields(145).value) & "','" & TglNul2(rsPr.Fields(146).value) & "','" & Trim(rsPr.Fields(147).value) & "','" & Trim(rsPr.Fields(148).value) & "','" & TglNul2(rsPr.Fields(149).value) & "','" & Trim(rsPr.Fields(150).value) & "')")
                  
30        Set rsPr = objO.FetchData("Select * from Hist_JobAdd WITH (NOLOCK) Where No_Wo = '" & nWO & "'")
31        If rsPr.RecordCount <> 0 Then
32            rsPr.MoveFirst
33            Do Until rsPr.EOF
34                Call objO.InsertDataHO("Insert Into Hist_JobAddnew (NO_WO,NO_CHASSIS,[DATE],KLPKERJA,REP,RTE,ACTRTE,MEKANIK,GRP,OFFn,DISC,[TIME],VALID,FREEC,FS_LOG_DEL,FD_TGL_DEL,No_inv,CABANG)  Values " & _
                  "('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & TglNul2(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "'," & RepRp(rsPr.Fields(5).value) & ", " & _
                  " " & RepRp(rsPr.Fields(6).value) & ",'" & Trim(rsPr.Fields(7).value) & "','" & Trim(rsPr.Fields(8).value) & "'," & RepRp(rsPr.Fields(9).value) & "," & RepRp(rsPr.Fields(10).value) & ",'" & Trim(rsPr.Fields(11).value) & "', " & _
                  "'" & Trim(rsPr.Fields(12).value) & "','" & Trim(rsPr.Fields(13).value) & "','" & Trim(rsPr.Fields(14).value) & "','" & TglNul2(rsPr.Fields(15).value) & "','" & Trim(rsPr.Fields(16).value) & "','" & Log.LogBengkel & "')")
35                rsPr.MoveNext
36            Loop
37        End If

38       On Error GoTo 0
39       Exit Sub

Transfer_INV1_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Transfer_INV1 of Form Layar")
End Sub

Private Sub DoTransferExec(rsEx As ADODB.Recordset)
      '"EXEC REPOTAF" | "Bulan = 12 And Tahun = 2012" | "EXEC" | 0 | 01/01/1900 | "EXEC REPOTAF nnnn"
          Dim rsTr As New ADODB.Recordset
1        On Error GoTo DoTransferExec_Error

2         Set rsTr = objO.FetchData(Replace(rsEx.Fields(5).value, "~", "'"))
3         If rsTr.RecordCount > 0 Then
4             Call objHO.OpenTransactionLog_NRM
5             Call objHO.AddNew_WO("Delete Bengkel99..Jam_TAF WHere " & Trim(rsEx.Fields(1).value) & " And Cabang = '" & Log.LogNamaNas & "'")
6             rsTr.MoveFirst
7             Do Until rsTr.EOF
8                 Call objHO.AddNew_WO("Insert Into Bengkel99..Jam_TAF Values (" & rsTr.Fields(0).value & ", " & rsTr.Fields(1).value & ",'" & rsTr.Fields(2).value & "','" & rsTr.Fields(3).value & "'," & RepRp(rsTr.Fields(4).value) & ")")
9                 rsTr.MoveNext
10            Loop
11            Call objHO.CloseTransactionLog_NRM
12            Call objO.Update_LOG(Trim(rsEx.Fields(1).value), (Trim(rsEx.Fields(2).value)), Trim(rsEx.Fields(0).value))
13        Else
14            Call objO.InsertData("Update LogHO Set LogHo = 2, transfer = Getdate() Where Kode = '" & Trim(rsEx.Fields(1).value) & "' And Status = '" & (Trim(rsEx.Fields(2).value)) & "' And Tipe = '" & Trim(rsEx.Fields(0).value) & "'")
15        End If

16       On Error GoTo 0
17       Exit Sub

DoTransferExec_Error:

18        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure DoTransferExec of Form Layar")
End Sub

Private Sub ViewDataAR()
    Dim fAR As New aLockAR
    fAR.show vbModal
End Sub

Function GensSql2Layar(xTabel As String, xUpdate As Boolean) As String
    Dim aa As String
    Dim xHasil As String
    Dim i As Integer
    
    xHasil = ""
    If xUpdate Then
        xHasil = "UPDATE " + xTabel + " SET "
        i = 1
        Do While gsSql1(i) <> ""
            If i <> 1 Then
                xHasil = xHasil + ","
            End If
            
            Select Case UCase(gsSql2(i))
                Case "GETDATE()"
                    xHasil = Left(xHasil, Len(xHasil) - 1)
            End Select
                
            aa = UCase(gsSql1(i))
            Select Case aa
                Case "PPN", "DISC", "DISCT", "PART_G", "PART_NG", "OIL_G", "OIL_NG", "MTRL", "MET", "SBLT", "LABOR", "BODY", "TOTAL", "FN_EXT_DISKON"
                    xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
                
                Case Else
                    If UCase(gsSql2(i)) <> "GETDATE()" Then
                        xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
                    End If
            End Select
            
            i = i + 1
        Loop
    Else
        xHasil = "INSERT INTO " + xTabel + " ("
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            xHasil = xHasil + gsSql1(i)
            i = i + 1
        Loop
        xHasil = xHasil + ") VALUES ("
        
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            Select Case UCase(gsSql2(i))
                Case "GETDATE()"
                    gsSql2(i) = "Convert(VarChar(12), getdate(), 101)"
                    xHasil = xHasil + gsSql2(i)
            End Select
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC", "DISCT", "PART_G", "PART_NG", "OIL_G", "OIL_NG", "MTRL", "MET", "SBLT", "LABOR", "BODY", "TOTAL", "FN_EXT_DISKON", "RTE1", "RTE2", "RTE3", "RTE4", "RTE5", "RTE6", "RTE7", "RTE8", "ACTRATE1", "ACTRATE2", "ACTRATE3", "ACTRATE4", "ACTRATE5", "ACTRATE6", "ACTRATE7", "ACTRATE8", "OFF1", "Off2", "Off3", "Off4", "Off5", "Off6", "Off7", "Off8", "FN_PAJAK", "FN_JML_ORDER", "FN_JML_RATE", "THS", "pds", "DISC1", "DISC2", "DISC3", "DISC4", "DISC5", "DISC6", "DISC7", "DISC8"
            
                xHasil = xHasil + gsSql2(i)
            
            Case Else
                If UCase(gsSql2(i)) <> "CONVERT(VARCHAR(12), GETDATE(), 101)" Then
                    xHasil = xHasil + "'" + gsSql2(i) + "'"
                End If
            End Select
            i = i + 1
        Loop
        xHasil = xHasil + ")"
    End If
    GensSql2Layar = xHasil
End Function
