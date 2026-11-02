VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmInfoNoPolisi 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ganti No. Polisi"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.CommandButton TblSimpan 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Simpan"
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
      TabIndex        =   6
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
      TabIndex        =   7
      Top             =   7065
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   525
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">??"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Index           =   0
      Left            =   3840
      TabIndex        =   9
      Top             =   540
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
      Left            =   2160
      TabIndex        =   3
      Top             =   930
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">??"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   1
      Top             =   525
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   2
      Top             =   525
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      Top             =   930
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   5
      Left            =   3360
      TabIndex        =   5
      Top             =   930
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCC"
      PromptChar      =   ">"
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7455
      Left            =   0
      Picture         =   "frmInfoNoPolisi.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Polisi Baru"
      Height          =   255
      Left            =   735
      TabIndex        =   10
      Top             =   990
      Width           =   1290
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Polisi Lama"
      Height          =   255
      Left            =   660
      TabIndex        =   8
      Top             =   585
      Width           =   1365
   End
End
Attribute VB_Name = "frmInfoNoPolisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nPol As String
Dim nPolBaru As String

Private nEst As String

Public Property Let passingXX(VALUES As String)
    nEst = VALUES
End Property

Private Sub hapusForm()
    For i = 0 To 5
        text(i) = ""
    Next
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub Form_Resize() '
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean
Dim pesan As String
Dim rsHlp As ADODB.Recordset

    boleh = True
    pesan = "Lengkapi Data " & vbCr
    If text(0) = "" And text(2) = "" And text(3) = "" Then
        boleh = False
        pesan = pesan + "No. Polisi Lama. "
    End If
    If text(1) = "" And text(4) = "" And text(5) = "" Then
        boleh = False
        pesan = pesan + "No. Polisi Baru. "
    End If
    
    If boleh Then
        nPol = Trim(text(0)) & IIf(Trim(text(0)) <> "", "-", "") & Trim(text(2)) & IIf(Trim(text(3)) <> "", "-", "") & Trim(text(3))
        nPolBaru = Trim(text(1)) & IIf(Trim(text(1)) <> "", "-", "") & Trim(text(4)) & IIf(Trim(text(5)) <> "", "-", "") & Trim(text(5))
        If MsgBox("Yakin mengubah No.Polisi " & Trim(nPol) & " menjadi " & Trim(nPolBaru) & " !", vbYesNo, "konfirmasi") = vbYes Then
            
            Set rsHlp = objO.FetchData("select * from CUSTDAT WITH(NOLOCK) where NO_POLISI = '" & nPol & "' AND NO_CHASSIS = '" & Trim(nEst) & "'")
            If Not rsHlp.BOF Then
                Call objO.InsertData("INSERT INTO GANTINOPOL VALUES ('" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', '" & gSA & "', '" & nPol & "', '" & nPolBaru & "')")
                Call objO.UpdateData("UPDATE CUSTDAT SET NO_POLISI = '" & nPolBaru & "' WHERE NO_POLISI = '" & nPol & "' AND NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "'")
                Call objO.InsertData("INSERT INTO HIST_PEMILIK SELECT DISTINCT FS_KD_PLG, FS_KD_WIL,CUSTOMER, CUST_ADRES, KOTA, KODEPOS, TELPON, GSM, CDMA, USEAGAMA, USEJOBDSC, USEFS_NM_HOBI, NO_CHASSIS, NO_POLISI, " _
                             & "NO_ENGINE, MDL, MODEL, FMDL, COLOR, KEYNO, FS_KM_AKHIR, FS_TH_MOBIL, FS_KD_DEALER, FS_NM_DEALER, DDATE, CODE, DOTAM, NPWP, PPN, DISC, USERADD, USENAME, USEADDR, USECITY, USEPOS, " _
                             & "USEPSTN, USEGSM, USECDMA, USET4LAHIR, USETGLLAHIR, ASURANSI, ASR_ADDR, ASR_CITY, ASR_TGLM, ASR_TGLS, ASR_LIMIT, ASR_NOTE, NOTE, '', '', '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "',PKP, KDPLG_ASR, USEEMAIL FROM CUSTDAT WITH(NOLOCK) WHERE NO_POLISI = '" & nPolBaru & "'")
                
                
                If Log.LogHOConn = 1 Then
                    If ConHO = False Then
                        Call objO.UpdateDataHO("UPDATE " & HODb & "..CUSTDAT SET NO_POLISI = '" & nPolBaru & "' WHERE NO_POLISI = '" & nPol & "' AND NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "'")
                        Dim rsHlp2 As ADODB.Recordset
                        Set rsHlp2 = objO.FetchData("SELECT * FROM CUSTDAT WITH(NOLOCK) WHERE NO_POLISI = '" & nPolBaru & "' AND NO_CHASSIS = '" & Trim(nEst) & "'")
                        If Not rsHlp2.EOF Then
                            Call objO.InsertDataHO("INSERT INTO " & HODb & "..HIST_PEMILIK (FS_KD_PLG, FS_KD_WIL,CUSTOMER, CUST_ADRES, KOTA, KODEPOS, TELPON, GSM, CDMA, USEAGAMA, USEJOBDSC, USEFS_NM_HOBI, NO_CHASSIS, NO_POLISI, " _
                                       & "NO_ENGINE, MDL, MODEL, FMDL, COLOR, KEYNO, FS_KM_AKHIR, FS_TH_MOBIL, FS_KD_DEALER, FS_NM_DEALER, DDATE, CODE, DOTAM, NPWP, PPN, DISC, USERADD, USENAME, USEADDR, USECITY, USEPOS, " _
                                       & "USEPSTN, USEGSM, USECDMA, USET4LAHIR, USETGLLAHIR, ASURANSI, ASR_ADDR, ASR_CITY, ASR_TGLM, ASR_TGLS, ASR_LIMIT, ASR_NOTE, NOTE, FS_LOG_DEL, FD_TGL_DEL, KDPLG_ASR) VALUES ('" & rsHlp2("FS_KD_PLG") & "', '" & rsHlp2("FS_KD_WIL") & "', '" & rsHlp2("CUSTOMER") & "', '" & rsHlp2("CUST_ADRES") & "', '" & rsHlp2("KOTA") & "', '" & rsHlp2("KODEPOS") & "', '" & rsHlp2("TELPON") & "', '" & rsHlp2("GSM") & "', '" & rsHlp2("CDMA") & "', '" & rsHlp2("USEAGAMA") & "', '" & rsHlp2("USEJOBDSC") & "', '" & rsHlp2("USEFS_NM_HOBI") & "', '" & rsHlp2("NO_CHASSIS") & "', '" & rsHlp2("NO_POLISI") & "', " _
                                       & "'" & rsHlp2("NO_ENGINE") & "', '" & rsHlp2("MDL") & "', '" & rsHlp2("MODEL") & "', '" & rsHlp2("FMDL") & "', '" & rsHlp2("COLOR") & "', '" & rsHlp2("KEYNO") & "', '" & rsHlp2("FS_KM_AKHIR") & "', '" & rsHlp2("FS_TH_MOBIL") & "', '" & rsHlp2("FS_KD_DEALER") & "', '" & rsHlp2("FS_NM_DEALER") & "', '" & rsHlp2("DDATE") & "', '" & rsHlp2("CODE") & "', '" & rsHlp2("DOTAM") & "', '" & rsHlp2("NPWP") & "', '" & rsHlp2("PPN") & "', '" & rsHlp2("DISC") & "', '" & rsHlp2("USERADD") & "', '" & rsHlp2("USENAME") & "', '" & rsHlp2("USEADDR") & "', '" & rsHlp2("USECITY") & "', '" & rsHlp2("USEPOS") & "', " _
                                       & "'" & rsHlp2("USEPSTN") & "', '" & rsHlp2("USEGSM") & "', '" & rsHlp2("USECDMA") & "', '" & rsHlp2("USET4LAHIR") & "', '" & rsHlp2("USETGLLAHIR") & "', '" & rsHlp2("ASURANSI") & "', '" & rsHlp2("ASR_ADDR") & "', '" & rsHlp2("ASR_CITY") & "', '" & rsHlp2("ASR_TGLM") & "', '" & rsHlp2("ASR_TGLS") & "', '" & rsHlp2("ASR_LIMIT") & "', '" & rsHlp2("ASR_NOTE") & "', '" & rsHlp2("NOTE") & "', '', '','" & rsHlp2("KDPLG_ASR") & "')")
                        End If
                        rsHlp2.Close
                    End If
                End If
                MsgBox "Proses Selesai !", vbOKOnly, ""
            Else
                MsgBox "No. Polisi yang mau diganti belum ada !!!"
            End If
            Call hapusForm
        End If
    Else
        MsgBox pesan, vbCritical, "KELENGKAPAN"
        Exit Sub
    End If
    text(0).SetFocus
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            If Index = 0 Then
                Dim nDpn1 As String
                Dim nTgh1 As String
                Dim nBlk1 As String
                Call pecahNoPol(gMainKeyCari, nDpn1, nTgh1, nBlk1)
                text(0) = Trim(nDpn1)
                text(2) = Trim(nTgh1)
                text(3) = Trim(nBlk1)
            End If
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
                gMainKeyCari = text(Index)
                gSearchOption = "CUSTOMER3"
                frmCari.Show (vbModal)
            Case 2, 3
                text(0).SetFocus
                Exit Sub
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
        Case 3
            If Trim(text(Index)) <> "" And Trim(text(0)) <> "" And Trim(text(2)) <> "" Then
                nPol = Trim(text(0)) & "-" & Trim(text(2)) & "-" & Trim(text(3))
                sSQL = "select nO_pOLISI, No_Chassis, Customer, CUST_ADRES, No_engine from Custdat WITH(NOLOCK) where NO_POLISI ='" & nPol & "' AND NO_CHASSIS = '" & Trim(nEst) & "'"
                Set rsSearch = objO.FetchData(sSQL)
                If rsSearch.BOF Then
                    Cancel = True
                    gMainKeyCari = nPol
                    gSearchOption = "CUSTOMER3"
                    frmCari.Show (vbModal)
                End If
            End If
    End Select
     
End Sub
