VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form frmKonversi 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7470
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11310
   Begin VB.CommandButton btnNextSvc 
      BackColor       =   &H0080C0FF&
      Caption         =   "Update Hitungan &Next Service"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton btnUpdate 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Update NoWO + NoFPS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   3495
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Konversi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin TDBDate6Ctl.TDBDate TglAwal 
      Height          =   330
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "KONVERSI.frx":0000
      Caption         =   "KONVERSI.frx":012C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "KONVERSI.frx":0198
      Keys            =   "KONVERSI.frx":01B6
      Spin            =   "KONVERSI.frx":0214
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   2118189057
      Value           =   2.12482692446619E-314
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TglAkhir 
      Height          =   330
      Left            =   4065
      TabIndex        =   1
      Top             =   240
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "KONVERSI.frx":023C
      Caption         =   "KONVERSI.frx":0368
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "KONVERSI.frx":03D4
      Keys            =   "KONVERSI.frx":03F2
      Spin            =   "KONVERSI.frx":0450
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   2118189057
      Value           =   2.12482692446619E-314
      CenturyMode     =   0
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Invoice dari :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   285
      Width           =   1605
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3585
      TabIndex        =   3
      Top             =   285
      Width           =   345
   End
End
Attribute VB_Name = "frmKonversi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objBilling As New BClsBilling
Dim rsHlp As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim nNilai As Double
Dim nDuoInv As Boolean
Dim subletOnly As Boolean
Dim vNamaTabel As String

Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
Dim hlp5 As String
Dim hlp6 As String
Dim hlp7 As String
Dim hlp8 As String
Dim hlp9 As String

Private Sub isiTINVOICE(Tgl1 As Date, Tgl2 As Date)

      Dim i As Integer

      '&&&&&&&&&&&&&&&&&&&&&
      Dim noInv As String
      Dim nmCust As String
      Dim alCust As String
      Dim ktCust As String
      Dim npwpCust As String
      Dim kdCust As String
      Dim nStatus As String
      Dim paid As String

          
1         vNamaTabel = "T_INVOICE2"
          
2         If Log.RPLHistori = 0 Then
3             Set rsTemp = objO.FetchData("SELECT CASE DELIV WHEN 'BTL' THEN CASE WHEN FD_RETUR BETWEEN '" & Format(Tgl1, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(Tgl2, "MM/DD/YYYY 23:59:59") & "' THEN 0 ELSE 1 END ELSE 1 END AS HIT_UNIT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, Z.FS_ALAMAT AS ALMPENJUAL, Z.FS_KOTA AS KTPENJUAL, Z.FS_NPWP AS NPWPPENJUAL, Z.KDPLG AS KDLGPENJUAL, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP, FD_RETUR " _
                                   & "FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C WITH(NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE LEN(NOINV) = 8 AND INVDATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND ((A.TPSVC = 'P' AND C.DOUBLEINV = 1) OR INVDATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND (FS_RC = 'R001')) ORDER BY NOINV")
4         Else
5             Set rsTemp = objO.FetchData("SELECT CASE DELIV WHEN 'BTL' THEN CASE WHEN FD_RETUR BETWEEN '" & Format(Tgl1, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(Tgl2, "MM/DD/YYYY 23:59:59") & "' THEN 0 ELSE 1 END ELSE 1 END AS HIT_UNIT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, Z.FS_ALAMAT AS ALMPENJUAL, Z.FS_KOTA AS KTPENJUAL, Z.FS_NPWP AS NPWPPENJUAL, Z.KDPLG AS KDLGPENJUAL, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP, FD_RETUR " _
                                   & "FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM C WITH(NOLOCK) ON C.FS_LOG_DEL = '' AND A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE LEN(NOINV) = 8 AND INVDATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND ((A.TPSVC = 'P' AND C.DOUBLEINV = 1) OR INVDATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND (FS_RC = 'R001')) ORDER BY NOINV")
6         End If
                                   '--:: 23 Agustus 2010 ; Menampilkan nilai retur ; DELIV <> 'BTL' AND ::--
7         If Not rsTemp.EOF Then
8             Do While Not rsTemp.EOF
9                 subletOnly = True
      '            If Trim(rsTemp("NOINV")) = "02211188" Then
      '                MsgBox "TEST"
      '            End If
10                For i = 1 To 8
11                    hlp = "REP" & i
12                    If rsTemp(hlp) <> "" Then
13                        subletOnly = False
14                        hlp2 = "KLPKERJA" & i
15                        hlp3 = "RTE" & i
16                        hlp4 = "MEKANIK" & i
17                        hlp5 = "OFF" & i
18                        hlp6 = "DISC" & i
19                        hlp7 = "FREEC" & i
20                        hlp8 = "GRP" & i
21                        hlp9 = "ACTRATE" & i
22                        hlp10 = "VALID" & i
                          'If createTemp(vNamaTabel) Then
23                            nDuoInv = False
24                            If rsTemp("NDUOINV") = 1 Then nDuoInv = True
25                            If nDuoInv Then
26                                If rsTemp(hlp7) = 0 Then
27                                    nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
28                                    nNilai = CDbl(rsTemp(hlp5)) * (1 - IIf((nDuoInv And IIf(Trim(rsTemp(hlp7)) = 1, True, False)), CDbl(rsTemp(hlp6)) / 100, 0))
29                                    noInv = Trim(rsTemp("NOINV")) & "-B"
30                                    nmCust = Nul(Trim(rsTemp("NMPENJUAL")))
31                                    alCust = Nul(Trim(rsTemp("ALMPENJUAL")))
32                                    ktCust = Nul(Trim(rsTemp("KTPENJUAL")))
33                                    npwpCust = Nul(Trim(rsTemp("NPWPPENJUAL")))
34                                    kdCust = Nul(Trim(rsTemp("KDLGPENJUAL")))
35                                Else
36                                    nStatus = rsTemp("STATUS")
37                                    nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100))
38                                    noInv = Trim(rsTemp("NOINV")) & "-A"
39                                    nmCust = Trim(rsTemp("CUSTOMER"))
40                                    alCust = Nul(Trim(rsTemp("CUST_ADRES")))
41                                    ktCust = Nul(Trim(rsTemp("KOTA")))
42                                    npwpCust = Nul(Trim(rsTemp("NPWP")))
43                                    kdCust = Nul(Trim(rsTemp("FS_KD_PLG")))
44                                End If
45                            Else
46                                If rsTemp("FS_RC") = "R001" Then
47                                    nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
48                                    nNilai = CDbl(rsTemp(hlp5)) * (1 - IIf((nDuoInv And IIf(Trim(rsTemp(hlp7)) = 1, True, False)), CDbl(rsTemp(hlp6)) / 100, 0))
49                                    noInv = Trim(rsTemp("NOINV"))
50                                    nmCust = Trim(rsTemp("NMPENJUAL"))
51                                    alCust = Nul(Trim(rsTemp("ALMPENJUAL")))
52                                    ktCust = Nul(Trim(rsTemp("KTPENJUAL")))
53                                    npwpCust = Nul(Trim(rsTemp("NPWPPENJUAL")))
54                                    kdCust = Nul(Trim(rsTemp("KDLGPENJUAL")))
55                                Else
56                                    nStatus = rsTemp("STATUS")
57                                    nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100))
58                                    noInv = Trim(rsTemp("NOINV"))
59                                    nmCust = Trim(rsTemp("CUSTOMER"))
60                                    alCust = Nul(Trim(rsTemp("CUST_ADRES")))
61                                    ktCust = Nul(Trim(rsTemp("KOTA")))
62                                    npwpCust = Nul(Trim(rsTemp("NPWP")))
63                                    kdCust = Nul(Trim(rsTemp("FS_KD_PLG")))
64                                End If
65                            End If
                              
66                            If nStatus = "AFFNONPPN" Then
67                                nNilai = nNilai * (100 / 110)
68                            End If
                              
                               
69                            Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
70                            If Not rsHlp.EOF Then
71                                sSQL = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET CUSTOMER = '" & trimSS(nmCust) & "',CUST_ADRES = '" & trimSS(alCust) & "',KOTA = '" & trimSS(ktCust) & "',TELPON = '" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("TELPON"))) & "',NO_POLISI = '" & trimSS(rsTemp("NO_POLISI")) & "',NO_ENGINE = '" & trimSS(rsTemp("NO_ENGINE")) & "', " & _
                                         "MDL = '" & trimSS(rsTemp("MDL")) & "',MODEL = '" & trimSS(rsTemp("MODEL")) & "',FMDL = '" & trimSS(rsTemp("FMDL")) & "',COLOR = '" & trimSS(rsTemp("COLOR")) & "',FS_KD_DEALER = '" & trimSS(rsTemp("FS_KD_DEALER")) & "',FS_NM_DEALER = '" & trimSS(rsTemp("FS_NM_DEALER")) & "',DDATE = '" & trimSS(rsTemp("DDATE")) & "',CODE = '" & trimSS(rsTemp("CODE")) & "', " & _
                                         "NPWP = '" & trimSS(npwpCust) & "', PPN = '" & setKoma(rsTemp("PPN")) & "',DISC = '" & setKoma(rsTemp("DISC")) & "',DISCT = '" & setKoma(rsTemp("DISCT")) & "',FRONT = '" & trimSS(rsTemp("FRONT")) & "', [GROUP] = '" & trimSS(rsTemp("GROUP")) & "',MECH = '" & trimSS(rsTemp("MECH")) & "',KM = '" & trimSS(rsTemp("KM")) & "', " & _
                                         "DATE = '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "',TIME = '" & trimSS(rsTemp("TIME")) & "',TIME_M = '" & trimSS(rsTemp("TIME_M")) & "',TIME_S = '" & trimSS(rsTemp("TIME_S")) & "',TIME_A = '" & trimSS(rsTemp("TIME_A")) & "',PART_G = " & setKoma(cekDuo(CDbl(rsTemp("PART_G")), noInv)) & ",PART_NG = " & setKoma(cekDuo(CDbl(rsTemp("PART_NG")), noInv)) & ",MTRL = " & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", " & _
                                         "OIL_G = " & setKoma(cekDuo(CDbl(rsTemp("OIL_G")), noInv)) & ",OIL_NG= " & setKoma(cekDuo(CDbl(rsTemp("OIL_NG")), noInv)) & ",MET = " & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & ", SBLT = " & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", " & _
                                         "LABOR = LABOR + " & setKoma(CDbl(nNilai)) & ",BODY = " & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & ",TOTAL = TOTAL + " & setKoma(CDbl(nNilai)) & ", JOB = '" & trimSS(rsTemp("JOB")) & "',SLIP = '" & trimSS(rsTemp("SLIP")) & "',INVDATE = '" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', " & _
                                         "STATUS = '" & trimSS(nStatus) & "',PAID = '" & trimSS(paid) & "',NOFAK = '" & trimSS(rsTemp("NOFAK")) & "',REMARK1 = '" & trimSS(rsTemp("REMARK1")) & "',SUGGEST = '" & trimSS(rsTemp("SUGGEST")) & "',DELIV = '" & trimSS(rsTemp("DELIV")) & "',SPK = '" & trimSS(rsTemp("SPK")) & "', " & _
                                         "RPR_COND = '" & trimSS(rsTemp("RPR_COND")) & "',FS_KREDIT = '" & IIf(Right(noInv, 2) = "-B", 1, trimSS(rsTemp("FS_KREDIT"))) & "',FN_PAJAK = FN_PAJAK + '" & setKoma(CDbl(nNilai) * IIf(nStatus <> "AFFNONPPN", 10 / 110, 0)) & "',FS_TH_MOBIL = '" & trimSS(rsTemp("FS_TH_MOBIL")) & "',FN_RETURN = '" & trimSS(rsTemp("FN_RETURN")) & "',FS_CATATAN = '" & trimSS(rsTemp("FS_CATATAN")) & "', " & _
                                         "FS_LOG_ADD = '" & trimSS(rsTemp("FS_LOG_ADD")) & "',FS_LOG_DEL = '" & trimSS(rsTemp("FS_LOG_DEL")) & "',FN_JML_ORDER = '" & trimSS(rsTemp("FN_JML_ORDER")) & "',FN_JML_RATE = '" & setKoma(rsTemp("FN_JML_RATE")) & "',ALASAN_INV = '" & trimSS(rsTemp("ALASAN_INV")) & "',FS_RC = '" & trimSS(rsTemp("FS_RC")) & "',FS_KD_PLG = '" & trimSS(kdCust) & "',FS_KD_WIL = '" & Left(trimSS(kdCust), 3) & "',FD_TGL_PART = '" & Format(rsTemp("FD_TGL_PART"), "MM/DD/YYYY") & "', " & _
                                         "FS_NO_RETUR = '" & trimSS(rsTemp("FS_NO_RETUR")) & "',FD_RETUR = '" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") & "',FS_KD_RETUR = '" & trimSS(rsTemp("FS_KD_RETUR")) & "',TPSVC = '" & trimSS(rsTemp("TPSVC")) & "',THS = '" & trimSS(rsTemp("THS")) & "',PDS = '" & trimSS(rsTemp("PDS")) & "',MOU = '" & trimSS(rsTemp("MOU")) & "',KODETIPE = '" & trimSS(rsTemp("KODETIPE")) & "',INVTIME = '" & trimSS(rsTemp("INVTIME")) & "',NOFAKTUR = '" & trimSS(rsTemp("NOFAKTUR")) & "', " & _
                                         "NOFP = '" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("NOFP"))) & "',TGLKASIR = '" & Format(rsTemp("TGLKASIR"), "MM/DD/YYYY HH:NN:SS") & "',NOFP2 = '" & IIf(Right(noInv, 2) = "-B", trimSS(rsTemp("NOFP2")), "") & "',KASIRADD = '" & trimSS(rsTemp("KASIRADD")) & "',CASHADD = '" & trimSS(rsTemp("CASHADD")) & "',TGLCASH = '" & Format(rsTemp("TGLCASH"), "MM/DD/YYYY HH:NN:SS") & "', " & _
                                         "" & hlp & " = '" & trimSS(rsTemp(hlp)) & "'," & hlp3 & " = '" & setKoma(CDbl(rsTemp(hlp3))) & "'," & hlp4 & " = '" & trimSS(rsTemp(hlp4)) & "'," & hlp8 & " = '" & trimSS(rsTemp(hlp8)) & "'," & hlp9 & " = '" & setKoma(CDbl(rsTemp(hlp9))) & "'," & hlp5 & " = '" & trimSS(rsTemp(hlp5)) & "'," & hlp2 & " = '" & trimSS(rsTemp(hlp2)) & "'," & hlp6 & " = '" & IIf(Right(noInv, 2) = "-B", 0, setKoma(CDbl(rsTemp(hlp6)))) & "'," & hlp10 & " = '" & trimSS(rsTemp(hlp10)) & "'," & hlp7 & " = '" & trimSS(rsTemp(hlp7)) & "' " & _
                                         "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'"
72                            Else
73                                sSQL = "INSERT INTO [TEMPDB]..[" & vNamaTabel & "] (CUSTOMER,CUST_ADRES,KOTA,TELPON,NO_POLISI,NO_ENGINE,NO_CHASSIS,MDL,MODEL,FMDL,COLOR,FS_KD_DEALER,FS_NM_DEALER,DDATE,CODE,NPWP,PPN,DISC,DISCT,FRONT,[GROUP],MECH,KM,DATE,TIME,TIME_M,TIME_S,TIME_A, " & _
                                         "PART_G,PART_NG,MTRL,OIL_G,OIL_NG,MET,SBLT,LABOR,BODY,TOTAL,JOB,SLIP,INVDATE,NO_WO,NOINV,STATUS,PAID,NOFAK,REMARK1,SUGGEST,DELIV,SPK,RPR_COND,FS_KREDIT,FN_PAJAK,FS_TH_MOBIL,FN_RETURN,FS_CATATAN,FS_LOG_ADD,FS_LOG_DEL,FN_JML_ORDER, " & _
                                         "FN_JML_RATE,ALASAN_INV,FS_RC,FS_KD_PLG,FS_KD_WIL,FD_TGL_PART,FS_NO_RETUR,FD_RETUR,FS_KD_RETUR,TPSVC,THS,PDS,MOU,KODETIPE,INVTIME,NOFAKTUR,NOFP,NOMB,TGLKASIR,NOFP2,KASIRADD,CASHADD,TGLCASH, " & _
                                         "" & hlp & ", " & hlp3 & ", " & hlp4 & ", " & hlp8 & ", " & hlp9 & ", " & hlp5 & ", " & hlp2 & ", " & hlp6 & ", " & hlp10 & ", " & hlp7 & ") VALUES (" & _
                                         "'" & trimSS(nmCust) & "','" & trimSS(alCust) & "','" & trimSS(ktCust) & "','" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("TELPON"))) & "','" & trimSS(rsTemp("NO_POLISI")) & "','" & trimSS(rsTemp("NO_ENGINE")) & "','" & trimSS(rsTemp("NO_CHASSIS")) & "', " & _
                                         "'" & trimSS(rsTemp("MDL")) & "','" & trimSS(rsTemp("MODEL")) & "','" & trimSS(rsTemp("FMDL")) & "','" & trimSS(rsTemp("COLOR")) & "','" & trimSS(rsTemp("FS_KD_DEALER")) & "','" & trimSS(rsTemp("FS_NM_DEALER")) & "','" & trimSS(rsTemp("DDATE")) & "','" & trimSS(rsTemp("CODE")) & "', " & _
                                         "'" & trimSS(npwpCust) & "', '" & setKoma(rsTemp("PPN")) & "','" & setKoma(rsTemp("DISC")) & "','" & setKoma(rsTemp("DISCT")) & "','" & trimSS(rsTemp("FRONT")) & "','" & trimSS(rsTemp("GROUP")) & "','" & trimSS(rsTemp("MECH")) & "','" & trimSS(rsTemp("KM")) & "', " & _
                                         "'" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "','" & trimSS(rsTemp("TIME")) & "','" & trimSS(rsTemp("TIME_M")) & "','" & trimSS(rsTemp("TIME_S")) & "','" & trimSS(rsTemp("TIME_A")) & "'," & setKoma(cekDuo(CDbl(rsTemp("PART_G")), noInv)) & "," & setKoma(cekDuo(CDbl(rsTemp("PART_NG")), noInv)) & "," & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", " & _
                                         "" & setKoma(cekDuo(CDbl(rsTemp("OIL_G")), noInv)) & ", " & setKoma(cekDuo(CDbl(rsTemp("OIL_NG")), noInv)) & ", " & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & ", " & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", " & _
                                         "" & setKoma(CDbl(nNilai)) & "," & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "," & setKoma(CDbl(nNilai)) & ", '" & trimSS(rsTemp("JOB")) & "','" & trimSS(rsTemp("SLIP")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', " & _
                                         "'" & trimSS(rsTemp("NO_WO")) & "','" & trimSS(noInv) & "','" & trimSS(nStatus) & "','" & trimSS(paid) & "','" & trimSS(rsTemp("NOFAK")) & "','" & trimSS(rsTemp("REMARK1")) & "','" & trimSS(rsTemp("SUGGEST")) & "','" & trimSS(rsTemp("DELIV")) & "','" & trimSS(rsTemp("SPK")) & "', " & _
                                         "'" & trimSS(rsTemp("RPR_COND")) & "','" & IIf(Right(noInv, 2) = "-B", 1, trimSS(rsTemp("FS_KREDIT"))) & "','" & setKoma(CDbl(nNilai) * IIf(nStatus <> "AFFNONPPN", 10 / 110, 0)) & "','" & trimSS(rsTemp("FS_TH_MOBIL")) & "','" & trimSS(rsTemp("FN_RETURN")) & "','" & trimSS(rsTemp("FS_CATATAN")) & "', " & _
                                         "'" & trimSS(rsTemp("FS_LOG_ADD")) & "','" & trimSS(rsTemp("FS_LOG_DEL")) & "','" & trimSS(rsTemp("FN_JML_ORDER")) & "','" & setKoma(rsTemp("FN_JML_RATE")) & "','" & trimSS(rsTemp("ALASAN_INV")) & "','" & trimSS(rsTemp("FS_RC")) & "','" & trimSS(kdCust) & "','" & Left(trimSS(kdCust), 3) & "','" & Format(rsTemp("FD_TGL_PART"), "MM/DD/YYYY") & "', " & _
                                         "'" & trimSS(rsTemp("FS_NO_RETUR")) & "','" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") & "','" & trimSS(rsTemp("FS_KD_RETUR")) & "','" & trimSS(rsTemp("TPSVC")) & "','" & trimSS(rsTemp("THS")) & "','" & trimSS(rsTemp("PDS")) & "','" & trimSS(rsTemp("MOU")) & "','" & trimSS(rsTemp("KODETIPE")) & "','" & trimSS(rsTemp("INVTIME")) & "','" & trimSS(rsTemp("NOFAKTUR")) & "', "
                                         
74                                 sSQL = sSQL + "'" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("NOFP"))) & "','','" & Format(rsTemp("TGLKASIR"), "MM/DD/YYYY HH:NN:SS") & "','" & IIf(Right(noInv, 2) = "-B", trimSS(rsTemp("NOFP2")), "") & "','" & trimSS(rsTemp("KASIRADD")) & "','" & trimSS(rsTemp("CASHADD")) & "','" & Format(rsTemp("TGLCASH"), "MM/DD/YYYY HH:NN:SS") & "', " & _
                                         "'" & trimSS(rsTemp(hlp)) & "', '" & setKoma(CDbl(rsTemp(hlp3))) & "', '" & trimSS(rsTemp(hlp4)) & "', '" & trimSS(rsTemp(hlp8)) & "', '" & setKoma(CDbl(rsTemp(hlp9))) & "', '" & trimSS(rsTemp(hlp5)) & "', '" & trimSS(rsTemp(hlp2)) & "', '" & IIf(Right(noInv, 2) = "-B", 0, setKoma(CDbl(rsTemp(hlp6)))) & "', '" & trimSS(rsTemp(hlp10)) & "', '" & trimSS(rsTemp(hlp7)) & "')"
75                            End If
76                            Call objO.UpdateData(sSQL)
                               
                              
77                            If nDuoInv Then
78                                If rsTemp(hlp7) = 0 Then
79                                    nStatus = rsTemp("STATUS")
80                                    nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100))
81                                    noInv = trimSS(rsTemp("NOINV")) & "-A"
82                                    nmCust = Trim(rsTemp("CUSTOMER"))
83                                    alCust = Nul(Trim(rsTemp("CUST_ADRES")))
84                                    ktCust = Nul(Trim(rsTemp("KOTA")))
85                                    npwpCust = Nul(Trim(rsTemp("NPWP")))
86                                    kdCust = Nul(Trim(rsTemp("FS_KD_PLG")))
                                       
87                                    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsTemp("NO_WO")) & "' AND NOINV = '" & Trim(noInv) & "'")
88                                    If Not rsHlp.EOF Then
89                                        sSQL = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET CUSTOMER = '" & trimSS(nmCust) & "',CUST_ADRES = '" & trimSS(alCust) & "',KOTA = '" & trimSS(ktCust) & "',TELPON = '" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("TELPON"))) & "',NO_POLISI = '" & trimSS(rsTemp("NO_POLISI")) & "',NO_ENGINE = '" & trimSS(rsTemp("NO_ENGINE")) & "', " & _
                                                 "MDL = '" & trimSS(rsTemp("MDL")) & "',MODEL = '" & trimSS(rsTemp("MODEL")) & "',FMDL = '" & trimSS(rsTemp("FMDL")) & "',COLOR = '" & trimSS(rsTemp("COLOR")) & "',FS_KD_DEALER = '" & trimSS(rsTemp("FS_KD_DEALER")) & "',FS_NM_DEALER = '" & trimSS(rsTemp("FS_NM_DEALER")) & "',DDATE = '" & trimSS(rsTemp("DDATE")) & "',CODE = '" & trimSS(rsTemp("CODE")) & "', " & _
                                                 "NPWP = '" & trimSS(npwpCust) & "', PPN = '" & setKoma(rsTemp("PPN")) & "',DISC = '" & setKoma(rsTemp("DISC")) & "',DISCT = '" & setKoma(rsTemp("DISCT")) & "',FRONT = '" & trimSS(rsTemp("FRONT")) & "', [GROUP] = '" & trimSS(rsTemp("GROUP")) & "',MECH = '" & trimSS(rsTemp("MECH")) & "',KM = '" & trimSS(rsTemp("KM")) & "', " & _
                                                 "DATE = '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "',TIME = '" & trimSS(rsTemp("TIME")) & "',TIME_M = '" & trimSS(rsTemp("TIME_M")) & "',TIME_S = '" & trimSS(rsTemp("TIME_S")) & "',TIME_A = '" & trimSS(rsTemp("TIME_A")) & "',PART_G = '" & setKoma(cekDuo(CDbl(rsTemp("PART_G")), noInv)) & "',PART_NG = '" & setKoma(cekDuo(CDbl(rsTemp("PART_NG")), noInv)) & "',MTRL = '" & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', " & _
                                                 "OIL_G = '" & setKoma(cekDuo(CDbl(rsTemp("OIL_G")), noInv)) & "',OIL_NG= '" & setKoma(cekDuo(CDbl(rsTemp("OIL_NG")), noInv)) & "',MET = '" & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & "', SBLT = '" & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', " & _
                                                 "LABOR = LABOR + '" & setKoma(CDbl(nNilai)) & "',BODY = '" & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "',TOTAL = TOTAL + '" & setKoma(CDbl(nNilai)) & "', JOB = '" & trimSS(rsTemp("JOB")) & "',SLIP = '" & trimSS(rsTemp("SLIP")) & "',INVDATE = '" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', " & _
                                                 "STATUS = '" & trimSS(nStatus) & "',PAID = '" & trimSS(paid) & "',NOFAK = '" & trimSS(rsTemp("NOFAK")) & "',REMARK1 = '" & trimSS(rsTemp("REMARK1")) & "',SUGGEST = '" & trimSS(rsTemp("SUGGEST")) & "',DELIV = '" & trimSS(rsTemp("DELIV")) & "',SPK = '" & trimSS(rsTemp("SPK")) & "', " & _
                                                 "RPR_COND = '" & trimSS(rsTemp("RPR_COND")) & "',FS_KREDIT = '" & IIf(Right(noInv, 2) = "-B", 1, trimSS(rsTemp("FS_KREDIT"))) & "',FN_PAJAK = FN_PAJAK + '" & setKoma(CDbl(nNilai) * IIf(nStatus <> "AFFNONPPN", 10 / 110, 0)) & "',FS_TH_MOBIL = '" & trimSS(rsTemp("FS_TH_MOBIL")) & "',FN_RETURN = '" & trimSS(rsTemp("FN_RETURN")) & "',FS_CATATAN = '" & trimSS(rsTemp("FS_CATATAN")) & "', " & _
                                                 "FS_LOG_ADD = '" & trimSS(rsTemp("FS_LOG_ADD")) & "',FS_LOG_DEL = '" & trimSS(rsTemp("FS_LOG_DEL")) & "',FN_JML_ORDER = '" & trimSS(rsTemp("FN_JML_ORDER")) & "',FN_JML_RATE = '" & setKoma(rsTemp("FN_JML_RATE")) & "',ALASAN_INV = '" & trimSS(rsTemp("ALASAN_INV")) & "',FS_RC = '" & trimSS(rsTemp("FS_RC")) & "',FS_KD_PLG = '" & trimSS(kdCust) & "',FS_KD_WIL = '" & Left(trimSS(kdCust), 3) & "',FD_TGL_PART = '" & Format(rsTemp("FD_TGL_PART"), "MM/DD/YYYY") & "', " & _
                                                 "FS_NO_RETUR = '" & trimSS(rsTemp("FS_NO_RETUR")) & "',FD_RETUR = '" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") & "',FS_KD_RETUR = '" & trimSS(rsTemp("FS_KD_RETUR")) & "',TPSVC = '" & trimSS(rsTemp("TPSVC")) & "',THS = '" & trimSS(rsTemp("THS")) & "',PDS = '" & trimSS(rsTemp("PDS")) & "',MOU = '" & trimSS(rsTemp("MOU")) & "',KODETIPE = '" & trimSS(rsTemp("KODETIPE")) & "',INVTIME = '" & trimSS(rsTemp("INVTIME")) & "',NOFAKTUR = '" & trimSS(rsTemp("NOFAKTUR")) & "', " & _
                                                 "NOFP = '" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("NOFP"))) & "',TGLKASIR = '" & Format(rsTemp("TGLKASIR"), "MM/DD/YYYY HH:NN:SS") & "',NOFP2 = '" & IIf(Right(noInv, 2) = "-B", trimSS(rsTemp("NOFP2")), "") & "',KASIRADD = '" & trimSS(rsTemp("KASIRADD")) & "',CASHADD = '" & trimSS(rsTemp("CASHADD")) & "',TGLCASH = '" & Format(rsTemp("TGLCASH"), "MM/DD/YYYY HH:NN:SS") & "', " & _
                                                 "" & hlp & " = '" & trimSS(rsTemp(hlp)) & "'," & hlp3 & " = '" & setKoma(CDbl(rsTemp(hlp3))) & "'," & hlp4 & " = '" & trimSS(rsTemp(hlp4)) & "'," & hlp8 & " = '" & trimSS(rsTemp(hlp8)) & "'," & hlp9 & " = '" & setKoma(CDbl(rsTemp(hlp9))) & "'," & hlp5 & " = '" & trimSS(rsTemp(hlp5)) & "'," & hlp2 & " = '" & trimSS(rsTemp(hlp2)) & "'," & hlp6 & " = '" & IIf(Right(noInv, 2) = "-B", 0, setKoma(CDbl(rsTemp(hlp6)))) & "'," & hlp10 & " = '" & trimSS(rsTemp(hlp10)) & "'," & hlp7 & " = '" & trimSS(rsTemp(hlp7)) & "' " & _
                                                 "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'"
90                                    Else
91                                        sSQL = "INSERT INTO [TEMPDB]..[" & vNamaTabel & "] (CUSTOMER,CUST_ADRES,KOTA,TELPON,NO_POLISI,NO_ENGINE,NO_CHASSIS,MDL,MODEL,FMDL,COLOR,FS_KD_DEALER,FS_NM_DEALER,DDATE,CODE,NPWP,PPN,DISC,DISCT,FRONT,[GROUP],MECH,KM,DATE,TIME,TIME_M,TIME_S,TIME_A, " & _
                                                 "PART_G,PART_NG,MTRL,OIL_G,OIL_NG,MET,SBLT,LABOR,BODY,TOTAL,JOB,SLIP,INVDATE,NO_WO,NOINV,STATUS,PAID,NOFAK,REMARK1,SUGGEST,DELIV,SPK,RPR_COND,FS_KREDIT,FN_PAJAK,FS_TH_MOBIL,FN_RETURN,FS_CATATAN,FS_LOG_ADD,FS_LOG_DEL,FN_JML_ORDER, " & _
                                                 "FN_JML_RATE,ALASAN_INV,FS_RC,FS_KD_PLG,FS_KD_WIL,FD_TGL_PART,FS_NO_RETUR,FD_RETUR,FS_KD_RETUR,TPSVC,THS,PDS,MOU,KODETIPE,INVTIME,NOFAKTUR,NOFP,NOMB,TGLKASIR,NOFP2,KASIRADD,CASHADD,TGLCASH, " & _
                                                 "" & hlp & ", " & hlp3 & ", " & hlp4 & ", " & hlp8 & ", " & hlp9 & ", " & hlp5 & ", " & hlp2 & ", " & hlp6 & ", " & hlp10 & ", " & hlp7 & ") VALUES (" & _
                                                 "'" & trimSS(nmCust) & "','" & trimSS(alCust) & "','" & trimSS(ktCust) & "','" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("TELPON"))) & "','" & trimSS(rsTemp("NO_POLISI")) & "','" & trimSS(rsTemp("NO_ENGINE")) & "','" & trimSS(rsTemp("NO_CHASSIS")) & "', " & _
                                                 "'" & trimSS(rsTemp("MDL")) & "','" & trimSS(rsTemp("MODEL")) & "','" & trimSS(rsTemp("FMDL")) & "','" & trimSS(rsTemp("COLOR")) & "','" & trimSS(rsTemp("FS_KD_DEALER")) & "','" & trimSS(rsTemp("FS_NM_DEALER")) & "','" & trimSS(rsTemp("DDATE")) & "','" & trimSS(rsTemp("CODE")) & "', " & _
                                                 "'" & trimSS(npwpCust) & "', '" & setKoma(rsTemp("PPN")) & "','" & setKoma(rsTemp("DISC")) & "','" & setKoma(rsTemp("DISCT")) & "','" & trimSS(rsTemp("FRONT")) & "','" & trimSS(rsTemp("GROUP")) & "','" & trimSS(rsTemp("MECH")) & "','" & trimSS(rsTemp("KM")) & "', " & _
                                                 "'" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "','" & trimSS(rsTemp("TIME")) & "','" & trimSS(rsTemp("TIME_M")) & "','" & trimSS(rsTemp("TIME_S")) & "','" & trimSS(rsTemp("TIME_A")) & "'," & setKoma(cekDuo(CDbl(rsTemp("PART_G")), noInv)) & "," & setKoma(cekDuo(CDbl(rsTemp("PART_NG")), noInv)) & "," & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", " & _
                                                 "" & setKoma(cekDuo(CDbl(rsTemp("OIL_G")), noInv)) & ", " & setKoma(cekDuo(CDbl(rsTemp("OIL_NG")), noInv)) & ", " & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & ", " & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", " & _
                                                 "" & setKoma(CDbl(nNilai)) & "," & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "," & setKoma(CDbl(nNilai)) & ", '" & trimSS(rsTemp("JOB")) & "','" & trimSS(rsTemp("SLIP")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', " & _
                                                 "'" & trimSS(rsTemp("NO_WO")) & "','" & trimSS(noInv) & "','" & trimSS(nStatus) & "','" & trimSS(paid) & "','" & trimSS(rsTemp("NOFAK")) & "','" & trimSS(rsTemp("REMARK1")) & "','" & trimSS(rsTemp("SUGGEST")) & "','" & trimSS(rsTemp("DELIV")) & "','" & trimSS(rsTemp("SPK")) & "', " & _
                                                 "'" & trimSS(rsTemp("RPR_COND")) & "','" & IIf(Right(noInv, 2) = "-B", 1, trimSS(rsTemp("FS_KREDIT"))) & "','" & setKoma(CDbl(nNilai) * IIf(nStatus <> "AFFNONPPN", 10 / 110, 0)) & "','" & trimSS(rsTemp("FS_TH_MOBIL")) & "','" & trimSS(rsTemp("FN_RETURN")) & "','" & trimSS(rsTemp("FS_CATATAN")) & "', " & _
                                                 "'" & trimSS(rsTemp("FS_LOG_ADD")) & "','" & trimSS(rsTemp("FS_LOG_DEL")) & "','" & trimSS(rsTemp("FN_JML_ORDER")) & "','" & setKoma(rsTemp("FN_JML_RATE")) & "','" & trimSS(rsTemp("ALASAN_INV")) & "','" & trimSS(rsTemp("FS_RC")) & "','" & trimSS(kdCust) & "','" & Left(trimSS(kdCust), 3) & "','" & Format(rsTemp("FD_TGL_PART"), "MM/DD/YYYY") & "', " & _
                                                 "'" & trimSS(rsTemp("FS_NO_RETUR")) & "','" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") & "','" & trimSS(rsTemp("FS_KD_RETUR")) & "','" & trimSS(rsTemp("TPSVC")) & "','" & trimSS(rsTemp("THS")) & "','" & trimSS(rsTemp("PDS")) & "','" & trimSS(rsTemp("MOU")) & "','" & trimSS(rsTemp("KODETIPE")) & "','" & trimSS(rsTemp("INVTIME")) & "','" & trimSS(rsTemp("NOFAKTUR")) & "', " & _
                                                 "'" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("NOFP"))) & "','','" & Format(rsTemp("TGLKASIR"), "MM/DD/YYYY HH:NN:SS") & "','" & IIf(Right(noInv, 2) = "-B", trimSS(rsTemp("NOFP2")), "") & "','" & trimSS(rsTemp("KASIRADD")) & "','" & trimSS(rsTemp("CASHADD")) & "','" & Format(rsTemp("TGLCASH"), "MM/DD/YYYY HH:NN:SS") & "', "
92                                         sSQL = sSQL + "'" & trimSS(rsTemp(hlp)) & "', '" & setKoma(CDbl(rsTemp(hlp3))) & "', '" & trimSS(rsTemp(hlp4)) & "', '" & trimSS(rsTemp(hlp8)) & "', '" & setKoma(CDbl(rsTemp(hlp9))) & "', '" & trimSS(rsTemp(hlp5)) & "', '" & trimSS(rsTemp(hlp2)) & "', '" & IIf(Right(noInv, 2) = "-B", 0, setKoma(CDbl(rsTemp(hlp6)))) & "', '" & trimSS(rsTemp(hlp10)) & "', '" & trimSS(rsTemp(hlp7)) & "')"
93                                    End If
94                                    Call objO.UpdateData(sSQL)
                                       
95                                ElseIf rsTemp(hlp7) = 1 Then
96                                    nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
97                                    nNilai = CDbl(rsTemp(hlp5)) * (1 - IIf((nDuoInv And IIf(Trim(rsTemp(hlp7)) = 1, True, False)), CDbl(rsTemp(hlp6)) / 100, 0))
98                                    noInv = Trim(rsTemp("NOINV")) & "-B"
99                                    nmCust = Nul(Trim(rsTemp("NMPENJUAL")))
100                                   alCust = Nul(Trim(rsTemp("ALMPENJUAL")))
101                                   ktCust = Nul(Trim(rsTemp("KTPENJUAL")))
102                                   npwpCust = Nul(Trim(rsTemp("NPWPPENJUAL")))
103                                   kdCust = Nul(Trim(rsTemp("KDLGPENJUAL")))
                                       
104                                   Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsTemp("NO_WO")) & "' AND NOINV = '" & Trim(noInv) & "'")
105                                   If Not rsHlp.EOF Then
106                                       sSQL = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET CUSTOMER = '" & trimSS(nmCust) & "',CUST_ADRES = '" & trimSS(alCust) & "',KOTA = '" & trimSS(ktCust) & "',TELPON = '" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("TELPON"))) & "',NO_POLISI = '" & trimSS(rsTemp("NO_POLISI")) & "',NO_ENGINE = '" & trimSS(rsTemp("NO_ENGINE")) & "', " & _
                                                 "MDL = '" & trimSS(rsTemp("MDL")) & "',MODEL = '" & trimSS(rsTemp("MODEL")) & "',FMDL = '" & trimSS(rsTemp("FMDL")) & "',COLOR = '" & trimSS(rsTemp("COLOR")) & "',FS_KD_DEALER = '" & trimSS(rsTemp("FS_KD_DEALER")) & "',FS_NM_DEALER = '" & trimSS(rsTemp("FS_NM_DEALER")) & "',DDATE = '" & trimSS(rsTemp("DDATE")) & "',CODE = '" & trimSS(rsTemp("CODE")) & "', " & _
                                                 "NPWP = '" & trimSS(npwpCust) & "', PPN = '" & setKoma(rsTemp("PPN")) & "',DISC = '" & setKoma(rsTemp("DISC")) & "',DISCT = '" & setKoma(rsTemp("DISCT")) & "',FRONT = '" & trimSS(rsTemp("FRONT")) & "', [GROUP] = '" & trimSS(rsTemp("GROUP")) & "',MECH = '" & trimSS(rsTemp("MECH")) & "',KM = '" & trimSS(rsTemp("KM")) & "', " & _
                                                 "DATE = '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "',TIME = '" & trimSS(rsTemp("TIME")) & "',TIME_M = '" & trimSS(rsTemp("TIME_M")) & "',TIME_S = '" & trimSS(rsTemp("TIME_S")) & "',TIME_A = '" & trimSS(rsTemp("TIME_A")) & "',PART_G = " & setKoma(cekDuo(CDbl(rsTemp("PART_G")), noInv)) & ",PART_NG = " & setKoma(cekDuo(CDbl(rsTemp("PART_NG")), noInv)) & ",MTRL = " & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", " & _
                                                 "OIL_G = " & setKoma(cekDuo(CDbl(rsTemp("OIL_G")), noInv)) & ",OIL_NG= " & setKoma(cekDuo(CDbl(rsTemp("OIL_NG")), noInv)) & ",MET = " & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & ", SBLT = " & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", " & _
                                                 "LABOR = LABOR + " & setKoma(CDbl(nNilai)) & ",BODY = " & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & ",TOTAL = TOTAL + " & setKoma(CDbl(nNilai)) & ", JOB = '" & trimSS(rsTemp("JOB")) & "',SLIP = '" & trimSS(rsTemp("SLIP")) & "',INVDATE = '" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', " & _
                                                 "STATUS = '" & trimSS(nStatus) & "',PAID = '" & trimSS(paid) & "',NOFAK = '" & trimSS(rsTemp("NOFAK")) & "',REMARK1 = '" & trimSS(rsTemp("REMARK1")) & "',SUGGEST = '" & trimSS(rsTemp("SUGGEST")) & "',DELIV = '" & trimSS(rsTemp("DELIV")) & "',SPK = '" & trimSS(rsTemp("SPK")) & "', " & _
                                                 "RPR_COND = '" & trimSS(rsTemp("RPR_COND")) & "',FS_KREDIT = '" & IIf(Right(noInv, 2) = "-B", 1, trimSS(rsTemp("FS_KREDIT"))) & "',FN_PAJAK = FN_PAJAK + '" & setKoma(CDbl(nNilai) * IIf(nStatus <> "AFFNONPPN", 10 / 110, 0)) & "',FS_TH_MOBIL = '" & trimSS(rsTemp("FS_TH_MOBIL")) & "',FN_RETURN = '" & trimSS(rsTemp("FN_RETURN")) & "',FS_CATATAN = '" & trimSS(rsTemp("FS_CATATAN")) & "', " & _
                                                 "FS_LOG_ADD = '" & trimSS(rsTemp("FS_LOG_ADD")) & "',FS_LOG_DEL = '" & trimSS(rsTemp("FS_LOG_DEL")) & "',FN_JML_ORDER = '" & trimSS(rsTemp("FN_JML_ORDER")) & "',FN_JML_RATE = '" & setKoma(rsTemp("FN_JML_RATE")) & "',ALASAN_INV = '" & trimSS(rsTemp("ALASAN_INV")) & "',FS_RC = '" & trimSS(rsTemp("FS_RC")) & "',FS_KD_PLG = '" & trimSS(kdCust) & "',FS_KD_WIL = '" & Left(trimSS(kdCust), 3) & "',FD_TGL_PART = '" & Format(rsTemp("FD_TGL_PART"), "MM/DD/YYYY") & "', " & _
                                                 "FS_NO_RETUR = '" & trimSS(rsTemp("FS_NO_RETUR")) & "',FD_RETUR = '" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") & "',FS_KD_RETUR = '" & trimSS(rsTemp("FS_KD_RETUR")) & "',TPSVC = '" & trimSS(rsTemp("TPSVC")) & "',THS = '" & trimSS(rsTemp("THS")) & "',PDS = '" & trimSS(rsTemp("PDS")) & "',MOU = '" & trimSS(rsTemp("MOU")) & "',KODETIPE = '" & trimSS(rsTemp("KODETIPE")) & "',INVTIME = '" & trimSS(rsTemp("INVTIME")) & "',NOFAKTUR = '" & trimSS(rsTemp("NOFAKTUR")) & "', " & _
                                                 "NOFP = '" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("NOFP"))) & "',TGLKASIR = '" & Format(rsTemp("TGLKASIR"), "MM/DD/YYYY HH:NN:SS") & "',NOFP2 = '" & IIf(Right(noInv, 2) = "-B", trimSS(rsTemp("NOFP2")), "") & "',KASIRADD = '" & trimSS(rsTemp("KASIRADD")) & "',CASHADD = '" & trimSS(rsTemp("CASHADD")) & "',TGLCASH = '" & Format(rsTemp("TGLCASH"), "MM/DD/YYYY HH:NN:SS") & "', " & _
                                                 "" & hlp & " = '" & trimSS(rsTemp(hlp)) & "'," & hlp3 & " = '" & setKoma(CDbl(rsTemp(hlp3))) & "'," & hlp4 & " = '" & trimSS(rsTemp(hlp4)) & "'," & hlp8 & " = '" & trimSS(rsTemp(hlp8)) & "'," & hlp9 & " = '" & setKoma(CDbl(rsTemp(hlp9))) & "'," & hlp5 & " = '" & trimSS(rsTemp(hlp5)) & "'," & hlp2 & " = '" & trimSS(rsTemp(hlp2)) & "'," & hlp6 & " = '" & setKoma(CDbl(rsTemp(hlp6))) & "'," & hlp10 & " = '" & trimSS(rsTemp(hlp10)) & "'," & hlp7 & " = '" & trimSS(rsTemp(hlp7)) & "' " & _
                                                 "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & trimSS(noInv) & "'"
107                                   Else
108                                       sSQL = "INSERT INTO [TEMPDB]..[" & vNamaTabel & "] (CUSTOMER,CUST_ADRES,KOTA,TELPON,NO_POLISI,NO_ENGINE,NO_CHASSIS,MDL,MODEL,FMDL,COLOR,FS_KD_DEALER,FS_NM_DEALER,DDATE,CODE,NPWP,PPN,DISC,DISCT,FRONT,[GROUP],MECH,KM,DATE,TIME,TIME_M,TIME_S,TIME_A, " & _
                                                 "PART_G,PART_NG,MTRL,OIL_G,OIL_NG,MET,SBLT,LABOR,BODY,TOTAL,JOB,SLIP,INVDATE,NO_WO,NOINV,STATUS,PAID,NOFAK,REMARK1,SUGGEST,DELIV,SPK,RPR_COND,FS_KREDIT,FN_PAJAK,FS_TH_MOBIL,FN_RETURN,FS_CATATAN,FS_LOG_ADD,FS_LOG_DEL,FN_JML_ORDER, " & _
                                                 "FN_JML_RATE,ALASAN_INV,FS_RC,FS_KD_PLG,FS_KD_WIL,FD_TGL_PART,FS_NO_RETUR,FD_RETUR,FS_KD_RETUR,TPSVC,THS,PDS,MOU,KODETIPE,INVTIME,NOFAKTUR,NOFP,NOMB,TGLKASIR,NOFP2,KASIRADD,CASHADD,TGLCASH, " & _
                                                 "" & hlp & ", " & hlp3 & ", " & hlp4 & ", " & hlp8 & ", " & hlp9 & ", " & hlp5 & ", " & hlp2 & ", " & hlp6 & ", " & hlp10 & ", " & hlp7 & ") VALUES (" & _
                                                 "'" & trimSS(nmCust) & "','" & trimSS(alCust) & "','" & trimSS(ktCust) & "','" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("TELPON"))) & "','" & trimSS(rsTemp("NO_POLISI")) & "','" & trimSS(rsTemp("NO_ENGINE")) & "','" & trimSS(rsTemp("NO_CHASSIS")) & "', " & _
                                                 "'" & trimSS(rsTemp("MDL")) & "','" & trimSS(rsTemp("MODEL")) & "','" & trimSS(rsTemp("FMDL")) & "','" & trimSS(rsTemp("COLOR")) & "','" & trimSS(rsTemp("FS_KD_DEALER")) & "','" & trimSS(rsTemp("FS_NM_DEALER")) & "','" & trimSS(rsTemp("DDATE")) & "','" & trimSS(rsTemp("CODE")) & "', " & _
                                                 "'" & trimSS(npwpCust) & "', '" & setKoma(rsTemp("PPN")) & "','" & setKoma(rsTemp("DISC")) & "','" & setKoma(rsTemp("DISCT")) & "','" & trimSS(rsTemp("FRONT")) & "','" & trimSS(rsTemp("GROUP")) & "','" & trimSS(rsTemp("MECH")) & "','" & trimSS(rsTemp("KM")) & "', " & _
                                                 "'" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "','" & trimSS(rsTemp("TIME")) & "','" & trimSS(rsTemp("TIME_M")) & "','" & trimSS(rsTemp("TIME_S")) & "','" & trimSS(rsTemp("TIME_A")) & "'," & setKoma(cekDuo(CDbl(rsTemp("PART_G")), noInv)) & "," & setKoma(cekDuo(CDbl(rsTemp("PART_NG")), noInv)) & "," & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", " & _
                                                 "" & setKoma(cekDuo(CDbl(rsTemp("OIL_G")), noInv)) & ", " & setKoma(cekDuo(CDbl(rsTemp("OIL_NG")), noInv)) & ", " & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & ", " & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", " & _
                                                 "" & setKoma(CDbl(nNilai)) & "," & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "," & setKoma(CDbl(nNilai)) & ", '" & trimSS(rsTemp("JOB")) & "','" & trimSS(rsTemp("SLIP")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', " & _
                                                 "'" & trimSS(rsTemp("NO_WO")) & "','" & trimSS(noInv) & "','" & trimSS(nStatus) & "','" & trimSS(paid) & "','" & trimSS(rsTemp("NOFAK")) & "','" & trimSS(rsTemp("REMARK1")) & "','" & trimSS(rsTemp("SUGGEST")) & "','" & trimSS(rsTemp("DELIV")) & "','" & trimSS(rsTemp("SPK")) & "', " & _
                                                 "'" & trimSS(rsTemp("RPR_COND")) & "','" & IIf(Right(noInv, 2) = "-B", 1, trimSS(rsTemp("FS_KREDIT"))) & "','" & setKoma(CDbl(nNilai) * IIf(nStatus <> "AFFNONPPN", 10 / 110, 0)) & "','" & trimSS(rsTemp("FS_TH_MOBIL")) & "','" & trimSS(rsTemp("FN_RETURN")) & "','" & trimSS(rsTemp("FS_CATATAN")) & "', " & _
                                                 "'" & trimSS(rsTemp("FS_LOG_ADD")) & "','" & trimSS(rsTemp("FS_LOG_DEL")) & "','" & trimSS(rsTemp("FN_JML_ORDER")) & "','" & setKoma(rsTemp("FN_JML_RATE")) & "','" & trimSS(rsTemp("ALASAN_INV")) & "','" & trimSS(rsTemp("FS_RC")) & "','" & trimSS(kdCust) & "','" & Left(trimSS(kdCust), 3) & "','" & Format(rsTemp("FD_TGL_PART"), "MM/DD/YYYY") & "', " & _
                                                 "'" & trimSS(rsTemp("FS_NO_RETUR")) & "','" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") & "','" & trimSS(rsTemp("FS_KD_RETUR")) & "','" & trimSS(rsTemp("TPSVC")) & "','" & trimSS(rsTemp("THS")) & "','" & trimSS(rsTemp("PDS")) & "','" & trimSS(rsTemp("MOU")) & "','" & trimSS(rsTemp("KODETIPE")) & "','" & trimSS(rsTemp("INVTIME")) & "','" & trimSS(rsTemp("NOFAKTUR")) & "', " & _
                                                 "'" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("NOFP"))) & "','','" & Format(rsTemp("TGLKASIR"), "MM/DD/YYYY HH:NN:SS") & "','" & IIf(Right(noInv, 2) = "-B", trimSS(rsTemp("NOFP2")), "") & "','" & trimSS(rsTemp("KASIRADD")) & "','" & trimSS(rsTemp("CASHADD")) & "','" & Format(rsTemp("TGLCASH"), "MM/DD/YYYY HH:NN:SS") & "', "
109                                        sSQL = sSQL + "'" & trimSS(rsTemp(hlp)) & "', '" & setKoma(CDbl(rsTemp(hlp3))) & "', '" & trimSS(rsTemp(hlp4)) & "', '" & trimSS(rsTemp(hlp8)) & "', '" & setKoma(CDbl(rsTemp(hlp9))) & "', '" & trimSS(rsTemp(hlp5)) & "', '" & trimSS(rsTemp(hlp2)) & "', '" & IIf(Right(noInv, 2) = "-B", 0, setKoma(CDbl(rsTemp(hlp6)))) & "', '" & trimSS(rsTemp(hlp10)) & "', '" & trimSS(rsTemp(hlp7)) & "')"
110                                   End If
111                                   Call objO.UpdateData(sSQL)
                                       
112                               End If
113                           End If
                          'End If
114                   End If
115               Next
116               Call addJob(rsTemp)
117               rsTemp.MoveNext
118           Loop
119           MsgBox "DATA BERHASIL DIKONVERSI "
120       Else
121           MsgBox "Tidak ada data !!"
122           Exit Sub
123       End If
End Sub

Private Sub addJob(rsTemp As ADODB.Recordset)
    Dim noInv As String
    Dim nmCust As String
    Dim alCust As String
    Dim ktCust As String
    Dim npwpCust As String
    Dim kdCust As String
    Dim nStatus As String
    Dim paid As String

    
    ' DARI ADD JOB
    Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WITH(NOLOCK) WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND FS_LOG_DEL = '" & Trim(rsTemp("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
    If Not rsAdd.EOF Then
        Do While Not rsAdd.EOF
            hlp2 = "KLPKERJA"
            hlp3 = "RTE"
            hlp4 = "MEKANIK"
            hlp5 = "OFFN"
            hlp6 = "FREEC"
            
            nDuoInv = False
            If rsTemp("NDUOINV") = 1 Then
                nDuoInv = True
            End If
            
            If nDuoInv Then
                If rsAdd(hlp6) = 0 Then
                    nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
                    nNilai = CDbl(rsAdd(hlp5)) * (1 - IIf((nDuoInv And IIf(Trim(rsAdd(hlp6)) = 1, True, False)), CDbl(rsAdd("DISC")) / 100, 0))
                    noInv = Trim(rsTemp("NOINV")) & "-B"
                    nmCust = Nul(Trim(rsTemp("NMPENJUAL")))
                    alCust = Nul(Trim(rsTemp("ALMPENJUAL")))
                    ktCust = Nul(Trim(rsTemp("KTPENJUAL")))
                    npwpCust = Nul(Trim(rsTemp("NPWPPENJUAL")))
                    kdCust = Nul(Trim(rsTemp("KDLGPENJUAL")))
                Else
                    nStatus = rsTemp("STATUS")
                    nNilai = CDbl(rsAdd(hlp5)) * (1 - (CDbl(rsAdd("DISC")) / 100))
                    noInv = Trim(rsTemp("NOINV")) & "-A"
                    nmCust = Trim(rsTemp("CUSTOMER"))
                    alCust = Nul(Trim(rsTemp("CUST_ADRES")))
                    ktCust = Nul(Trim(rsTemp("KOTA")))
                    npwpCust = Nul(Trim(rsTemp("NPWP")))
                    kdCust = Nul(Trim(rsTemp("FS_KD_PLG")))
                End If
            Else
                nStatus = rsTemp("STATUS")
                nNilai = CDbl(rsAdd(hlp5)) * (1 - (CDbl(rsAdd("DISC")) / 100))
                noInv = Trim(rsTemp("NOINV"))
                nmCust = Trim(rsTemp("CUSTOMER"))
                alCust = Nul(Trim(rsTemp("CUST_ADRES")))
                ktCust = Nul(Trim(rsTemp("KOTA")))
                npwpCust = Nul(Trim(rsTemp("NPWP")))
                kdCust = Nul(Trim(rsTemp("FS_KD_PLG")))
            End If
            
            If Trim(nStatus) = "AFFNONPPN" Then
                nNilai = nNilai * (100 / 110)
            End If
            
             
            Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsTemp("NO_WO")) & "' AND NOINV = '" & Trim(noInv) & "'")
            If Not rsHlp.EOF Then
                sSQL = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET CUSTOMER = '" & trimSS(nmCust) & "',CUST_ADRES = '" & trimSS(alCust) & "',KOTA = '" & trimSS(ktCust) & "',TELPON = '" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("TELPON"))) & "',NO_POLISI = '" & trimSS(rsTemp("NO_POLISI")) & "',NO_ENGINE = '" & trimSS(rsTemp("NO_ENGINE")) & "', " & _
                       "MDL = '" & trimSS(rsTemp("MDL")) & "',MODEL = '" & trimSS(rsTemp("MODEL")) & "',FMDL = '" & trimSS(rsTemp("FMDL")) & "',COLOR = '" & trimSS(rsTemp("COLOR")) & "',FS_KD_DEALER = '" & trimSS(rsTemp("FS_KD_DEALER")) & "',FS_NM_DEALER = '" & trimSS(rsTemp("FS_NM_DEALER")) & "',DDATE = '" & trimSS(rsTemp("DDATE")) & "',CODE = '" & trimSS(rsTemp("CODE")) & "', " & _
                       "NPWP = '" & trimSS(npwpCust) & "', PPN = '" & trimSS(rsTemp("PPN")) & "',DISC = '" & trimSS(rsTemp("DISC")) & "',DISCT = '" & trimSS(rsTemp("DISCT")) & "',FRONT = '" & trimSS(rsTemp("FRONT")) & "', [GROUP] = '" & trimSS(rsTemp("GROUP")) & "',MECH = '" & trimSS(rsTemp("MECH")) & "',KM = '" & trimSS(rsTemp("KM")) & "', " & _
                       "DATE = '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "',TIME = '" & trimSS(rsTemp("TIME")) & "',TIME_M = '" & trimSS(rsTemp("TIME_M")) & "',TIME_S = '" & trimSS(rsTemp("TIME_S")) & "',TIME_A = '" & trimSS(rsTemp("TIME_A")) & "',PART_G = " & setKoma(cekDuo(CDbl(rsTemp("PART_G")), noInv)) & ",PART_NG = " & setKoma(cekDuo(CDbl(rsTemp("PART_NG")), noInv)) & ",MTRL = " & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", " & _
                       "OIL_G = " & setKoma(cekDuo(CDbl(rsTemp("OIL_G")), noInv)) & ",OIL_NG= " & setKoma(cekDuo(CDbl(rsTemp("OIL_NG")), noInv)) & ",MET = " & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & ", SBLT = " & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", " & _
                       "LABOR = LABOR + " & setKoma(CDbl(nNilai)) & ",BODY = " & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & ",TOTAL = TOTAL + " & setKoma(CDbl(nNilai)) & ", JOB = '" & trimSS(rsTemp("JOB")) & "',SLIP = '" & trimSS(rsTemp("SLIP")) & "',INVDATE = '" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', " & _
                       "STATUS = '" & trimSS(nStatus) & "',PAID = '" & trimSS(paid) & "',NOFAK = '" & trimSS(rsTemp("NOFAK")) & "',REMARK1 = '" & trimSS(rsTemp("REMARK1")) & "',SUGGEST = '" & trimSS(rsTemp("SUGGEST")) & "',DELIV = '" & trimSS(rsTemp("DELIV")) & "',SPK = '" & trimSS(rsTemp("SPK")) & "', " & _
                       "RPR_COND = '" & trimSS(rsTemp("RPR_COND")) & "',FS_KREDIT = '" & IIf(Right(noInv, 2) = "-B", 1, trimSS(rsTemp("FS_KREDIT"))) & "',FN_PAJAK = FN_PAJAK + '" & setKoma(CDbl(nNilai) * IIf(nStatus <> "AFFNONPPN", 10 / 110, 0)) & "',FS_TH_MOBIL = '" & trimSS(rsTemp("FS_TH_MOBIL")) & "',FN_RETURN = '" & trimSS(rsTemp("FN_RETURN")) & "',FS_CATATAN = '" & trimSS(rsTemp("FS_CATATAN")) & "', " & _
                       "FS_LOG_ADD = '" & trimSS(rsTemp("FS_LOG_ADD")) & "',FS_LOG_DEL = '" & trimSS(rsTemp("FS_LOG_DEL")) & "',FN_JML_ORDER = '" & trimSS(rsTemp("FN_JML_ORDER")) & "',FN_JML_RATE = '" & trimSS(rsTemp("FN_JML_RATE")) & "',ALASAN_INV = '" & trimSS(rsTemp("ALASAN_INV")) & "',FS_RC = '" & trimSS(rsTemp("FS_RC")) & "',FS_KD_PLG = '" & trimSS(kdCust) & "',FS_KD_WIL = '" & Left(trimSS(kdCust), 3) & "',FD_TGL_PART = '" & Format(rsTemp("FD_TGL_PART"), "MM/DD/YYYY") & "', " & _
                       "FS_NO_RETUR = '" & trimSS(rsTemp("FS_NO_RETUR")) & "',FD_RETUR = '" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") & "',FS_KD_RETUR = '" & trimSS(rsTemp("FS_KD_RETUR")) & "',TPSVC = '" & trimSS(rsTemp("TPSVC")) & "',THS = '" & trimSS(rsTemp("THS")) & "',PDS = '" & trimSS(rsTemp("PDS")) & "',MOU = '" & trimSS(rsTemp("MOU")) & "',KODETIPE = '" & trimSS(rsTemp("KODETIPE")) & "',INVTIME = '" & trimSS(rsTemp("INVTIME")) & "',NOFAKTUR = '" & trimSS(rsTemp("NOFAKTUR")) & "', " & _
                       "NOFP = '" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("NOFP"))) & "',TGLKASIR = '" & Format(rsTemp("TGLKASIR"), "MM/DD/YYYY HH:NN:SS") & "',NOFP2 = '" & IIf(Right(noInv, 2) = "-B", trimSS(rsTemp("NOFP2")), "") & "',KASIRADD = '" & trimSS(rsTemp("KASIRADD")) & "',CASHADD = '" & trimSS(rsTemp("CASHADD")) & "',TGLCASH = '" & Format(rsTemp("TGLCASH"), "MM/DD/YYYY HH:NN:SS") & "' " & _
                       "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'"
            Else
                sSQL = "INSERT INTO [TEMPDB]..[" & vNamaTabel & "] (CUSTOMER,CUST_ADRES,KOTA,TELPON,NO_POLISI,NO_ENGINE,NO_CHASSIS,MDL,MODEL,FMDL,COLOR,FS_KD_DEALER,FS_NM_DEALER,DDATE,CODE,NPWP,PPN,DISC,DISCT,FRONT,[GROUP],MECH,KM,DATE,TIME,TIME_M,TIME_S,TIME_A, " & _
                       "PART_G,PART_NG,MTRL,OIL_G,OIL_NG,MET,SBLT,LABOR,BODY,TOTAL,JOB,SLIP,INVDATE,NO_WO,NOINV,STATUS,PAID,NOFAK,REMARK1,SUGGEST,DELIV,SPK,RPR_COND,FS_KREDIT,FN_PAJAK,FS_TH_MOBIL,FN_RETURN,FS_CATATAN,FS_LOG_ADD,FS_LOG_DEL,FN_JML_ORDER, " & _
                       "FN_JML_RATE,ALASAN_INV,FS_RC,FS_KD_PLG,FS_KD_WIL,FD_TGL_PART,FS_NO_RETUR,FD_RETUR,FS_KD_RETUR,TPSVC,THS,PDS,MOU,KODETIPE,INVTIME,NOFAKTUR,NOFP,NOMB,TGLKASIR,NOFP2,KASIRADD,CASHADD,TGLCASH) VALUES (" & _
                       "'" & trimSS(nmCust) & "','" & trimSS(alCust) & "','" & trimSS(ktCust) & "','" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("TELPON"))) & "','" & trimSS(rsTemp("NO_POLISI")) & "','" & trimSS(rsTemp("NO_ENGINE")) & "','" & trimSS(rsTemp("NO_CHASSIS")) & "', " & _
                       "'" & trimSS(rsTemp("MDL")) & "','" & trimSS(rsTemp("MODEL")) & "','" & trimSS(rsTemp("FMDL")) & "','" & trimSS(rsTemp("COLOR")) & "','" & trimSS(rsTemp("FS_KD_DEALER")) & "','" & trimSS(rsTemp("FS_NM_DEALER")) & "','" & trimSS(rsTemp("DDATE")) & "','" & trimSS(rsTemp("CODE")) & "', " & _
                       "'" & trimSS(npwpCust) & "', '" & trimSS(rsTemp("PPN")) & "','" & trimSS(rsTemp("DISC")) & "','" & trimSS(rsTemp("DISCT")) & "','" & trimSS(rsTemp("FRONT")) & "','" & trimSS(rsTemp("GROUP")) & "','" & trimSS(rsTemp("MECH")) & "','" & trimSS(rsTemp("KM")) & "', " & _
                       "'" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "','" & trimSS(rsTemp("TIME")) & "','" & trimSS(rsTemp("TIME_M")) & "','" & trimSS(rsTemp("TIME_S")) & "','" & trimSS(rsTemp("TIME_A")) & "'," & setKoma(cekDuo(CDbl(rsTemp("PART_G")), noInv)) & "," & setKoma(cekDuo(CDbl(rsTemp("PART_NG")), noInv)) & "," & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", " & _
                       "" & setKoma(cekDuo(CDbl(rsTemp("OIL_G")), noInv)) & ", " & setKoma(cekDuo(CDbl(rsTemp("OIL_NG")), noInv)) & ", " & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & ", " & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", " & _
                       "" & setKoma(CDbl(nNilai)) & "," & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "," & setKoma(CDbl(nNilai)) & ", '" & trimSS(rsTemp("JOB")) & "','" & trimSS(rsTemp("SLIP")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', " & _
                       "'" & trimSS(rsTemp("NO_WO")) & "','" & trimSS(noInv) & "','" & trimSS(nStatus) & "','" & trimSS(paid) & "','" & trimSS(rsTemp("NOFAK")) & "','" & trimSS(rsTemp("REMARK1")) & "','" & trimSS(rsTemp("SUGGEST")) & "','" & trimSS(rsTemp("DELIV")) & "','" & trimSS(rsTemp("SPK")) & "', " & _
                       "'" & trimSS(rsTemp("RPR_COND")) & "','" & IIf(Right(noInv, 2) = "-B", 1, trimSS(rsTemp("FS_KREDIT"))) & "','" & setKoma(CDbl(nNilai) * IIf(nStatus <> "AFFNONPPN", 10 / 110, 0)) & "','" & trimSS(rsTemp("FS_TH_MOBIL")) & "','" & trimSS(rsTemp("FN_RETURN")) & "','" & trimSS(rsTemp("FS_CATATAN")) & "', " & _
                       "'" & trimSS(rsTemp("FS_LOG_ADD")) & "','" & trimSS(rsTemp("FS_LOG_DEL")) & "','" & trimSS(rsTemp("FN_JML_ORDER")) & "','" & trimSS(rsTemp("FN_JML_RATE")) & "','" & trimSS(rsTemp("ALASAN_INV")) & "','" & trimSS(rsTemp("FS_RC")) & "','" & trimSS(kdCust) & "','" & Left(trimSS(kdCust), 3) & "','" & Format(rsTemp("FD_TGL_PART"), "MM/DD/YYYY") & "', " & _
                       "'" & trimSS(rsTemp("FS_NO_RETUR")) & "','" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") & "','" & trimSS(rsTemp("FS_KD_RETUR")) & "','" & trimSS(rsTemp("TPSVC")) & "','" & trimSS(rsTemp("THS")) & "','" & trimSS(rsTemp("PDS")) & "','" & trimSS(rsTemp("MOU")) & "','" & trimSS(rsTemp("KODETIPE")) & "','" & trimSS(rsTemp("INVTIME")) & "','" & trimSS(rsTemp("NOFAKTUR")) & "', " & _
                       "'" & IIf(Right(noInv, 2) = "-B", "", trimSS(rsTemp("NOFP"))) & "','','" & Format(rsTemp("TGLKASIR"), "MM/DD/YYYY HH:NN:SS") & "','" & IIf(Right(noInv, 2) = "-B", trimSS(rsTemp("NOFP2")), "") & "','" & trimSS(rsTemp("KASIRADD")) & "','" & trimSS(rsTemp("CASHADD")) & "','" & Format(rsTemp("TGLCASH"), "MM/DD/YYYY HH:NN:SS") & "')"
            End If
            Call objO.UpdateData(sSQL)
             
            rsAdd.MoveNext
        Loop
    End If
    If rsAdd.State <> adStateClosed Then rsAdd.Close
End Sub

Private Function statusDealer(codeStatus As Integer)
    
    Select Case codeStatus
        Case 1
            statusDealer = "AFFNONPPN"
        Case 2
            statusDealer = "AFFPPN"
        Case 3
            statusDealer = "EXT"
    End Select
End Function

Private Function cekDuo(hrg As Double, nINV As String, Optional stat As Integer)
    
    If Right(Trim(nINV), 2) = "-B" Then
        cekDuo = 0
        If stat = 1 Then
            cekDuo = CDbl(hrg)
        End If
    Else
        cekDuo = CDbl(hrg)
        If stat = 1 Then
            cekDuo = 0
        End If
    End If
End Function

Private Sub btnNextSvc_Click()
Dim rsSearch As New ADODB.Recordset
Dim n As Integer

    'On Error GoTo vbError
    
    If Log.RPLHistori = 0 Then
        Set rsSearch = objO.FetchData("SELECT distinct A.NO_CHASSIS, DATE, a.KM, A.FS_RC, FS_NILAI, CASE WHEN (SELECT HEAVYUSE FROM JOBDTL WITH(NOLOCK) WHERE NO_WO = A.NO_wO) = 1 THEN 5000 ELSE 10000 END AS KELIP, A.NOINV, A.NO_WO, A.CUSTOMER, A.CUST_ADRES, A.NO_POLISI, B.TELPON, B.GSM, B.CDMA " & _
                                 "FROM HISTORY A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_REGULAR C WITH(NOLOCK) ON A.FS_RC = C.FS_RC WHERE A.FS_RC <> '' AND A.FS_LOG_DEL = '' " & _
                                 "and a.no_chassis = 'MHF11KF8320062181' ") 'and b.FS_TH_MOBIL = '1991'")
    Else
        Set rsSearch = objO.FetchData("SELECT distinct A.NO_CHASSIS, DATE, a.KM, A.FS_RC, FS_NILAI, CASE WHEN (SELECT HEAVYUSE FROM JOBDTL WITH(NOLOCK) WHERE NO_WO = A.NO_wO) = 1 THEN 5000 ELSE 10000 END AS KELIP, A.NOINV, A.NO_WO, A.CUSTOMER, A.CUST_ADRES, A.NO_POLISI, B.TELPON, B.GSM, B.CDMA " & _
                                 "FROM HISTORY A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN [" & dbReplikasiHO & "]..T_REGULAR C WITH(NOLOCK) ON A.FS_RC = C.FS_RC WHERE A.FS_RC <> '' AND A.FS_LOG_DEL = '' " & _
                                 "and a.no_chassis = 'MHF11KF8320062181' ") 'and b.FS_TH_MOBIL = '1991'")
    End If
    If Not rsSearch.EOF Then
        n = rsSearch.RecordCount
        Do While Not rsSearch.EOF
            objO.FetchData ("EXEC " & Trim(hTipePenghitunganDealer) & " '" & Trim(rsSearch("NO_CHASSIS")) & "', '" & Format(rsSearch("DATE"), "MM/DD/YYYY") & "', '" & CDbl(Trim(rsSearch("KM"))) & "', '" & Trim(rsSearch("FS_NILAI")) & "', '" & Trim(rsSearch("KELIP")) & "', '" & Trim(rsSearch("NOINV")) & "', '" & Trim(rsSearch("NO_WO")) & "','" & Trim(rsSearch("CUSTOMER")) & "','" & Trim(rsSearch("CUST_ADRES")) & "', '" & Trim(rsSearch("NO_POLISI")) & "', '" & Trim(rsSearch("TELPON")) & "', '" & Trim(rsSearch("GSM")) & "','" & Trim(rsSearch("CDMA")) & "'")
            n = n - 1
            rsSearch.MoveNext
        Loop
    End If
     
    MsgBox "COMPLETE TRANSFER"
    Exit Sub
    
vbError:
    MsgBox "BELUM COMPLETE"
    Exit Sub
End Sub

Private Sub btnUpdate_Click()
Dim accDb As String
Dim rsSearch As New ADODB.Recordset

    On Error GoTo vbError
    If Month(TglAwal) <> Month(TglAkhir) Then
        If year(TglAwal) <> year(TglAkhir) Then
            MsgBox "Isi Range Tanggal pada bulan & tahun yang sama.. "
        Else
            MsgBox "Isi Range Tanggal pada bulan yang sama.. "
        End If
    Else
             
            Set rsHlp = objO.FetchDataACC("SELECT Distinct NoBukti FROM GLTRN WITH(NOLOCK) WHERE isnull(NOWO,'') = '' AND left(NoFaktur,2) = 'E-' AND LEFT(NOBUKTI,2) = '" & Trim(gKodeFaktur) & "' AND TGLTRN BETWEEN '" & Format(TglAwal, "MM/DD/YYYY") & "' AND '" & Format(TglAkhir, "MM/DD/YYYY") & "' ORDER BY NOBUKTI")
            If Not rsHlp.EOF Then
                Do While Not rsHlp.EOF
                     
                    Set rsSearch = objO.FetchData("SELECT NO_WO, CASE WHEN NOFP = '' THEN NOFP2 ELSE NOFP END AS NOFPS FROM T_INVOICE WITH(NOLOCK) WHERE NOINV = '" & Trim(Replace(Replace(rsHlp("NOBUKTI"), gKodeFaktur, ""), "R", "")) & "'")
                    If Not rsSearch.EOF Then
                        Call objO.UpdateDataACC("UPDATE GLTRN SET NOWO = '" & Trim(rsSearch("NO_WO")) & "', NOFPS = '" & Trim(rsSearch("NOFPS")) & "' WHERE NOBUKTI = '" & Trim(rsHlp("NOBUKTI")) & "'")
                    End If
                    rsHlp.MoveNext
                Loop
                MsgBox "Data Berhasil diupdate. "
            Else
                MsgBox "Tidak ada data yang perlu diupdate. "
            End If
            
             
            Set rsHlp = objO.FetchDataACC("SELECT Distinct NoBukti FROM GLTRN WITH(NOLOCK) WHERE isnull(NOWO,'') = '' AND KODESUMBER = 'SV' AND LEFT(NOBUKTI,2) = 'D-' AND TGLTRN BETWEEN '" & Format(TglAwal, "MM/DD/YYYY") & "' AND '" & Format(TglAkhir, "MM/DD/YYYY") & "' ORDER BY NOBUKTI")
            If Not rsHlp.EOF Then
                Do While Not rsHlp.EOF
                     
                    Set rsSearch = objO.FetchData("SELECT NO_WO, CASE WHEN NOFP = '' THEN NOFP2 ELSE NOFP END AS NOFPS FROM T_INVOICE WITH(NOLOCK) WHERE NOINV = '" & Trim(Replace(Replace(rsHlp("NOBUKTI"), gKodeFaktur, ""), "R", "")) & "'")
                    If Not rsSearch.EOF Then
                        Call objO.UpdateDataACC("UPDATE GLTRN SET NOWO = '" & Trim(rsSearch("NO_WO")) & "', NOFPS = '" & Trim(rsSearch("NOFPS")) & "' WHERE NOBUKTI = '" & Trim(rsHlp("NOBUKTI")) & "'")
                    End If
                     
                    rsHlp.MoveNext
                Loop
                MsgBox "Data Berhasil diupdate. "
            Else
                MsgBox "Tidak ada data yang perlu diupdate. "
            End If
        End If
    
    Exit Sub

vbError:
    Exit Sub
End Sub

Private Sub cmdRefresh_Click()
    If MsgBox("Akan Konversi Data ? ", vbQuestion + vbDefaultButton2, "Konversi") = vbYes Then
        Call isiTINVOICE(TglAwal, TglAkhir)
    End If
End Sub

Private Sub Form_Load()
    
    TglAwal = Format(isnow, "DD/MM/YYYY")
    TglAkhir = Format(isnow, "DD/MM/YYYY")
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub
