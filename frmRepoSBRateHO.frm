VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmRepoSBRateHO 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form Ratantion Rate"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   13275
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
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
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
      Height          =   270
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   7575
      DisplayGroupTree=   0   'False
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
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   12375
      Begin VB.CommandButton cmdCetak 
         BackColor       =   &H0080FF80&
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   375
         Left            =   1650
         TabIndex        =   7
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   97517569
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   97517569
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1000
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSAdodcLib.Adodc AdcKode 
         Height          =   330
         Left            =   8160
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DCKodeGroup 
         Height          =   375
         Left            =   1650
         TabIndex        =   14
         Top             =   600
         Width           =   3195
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorEven   =   16711680
         BackColorOdd    =   8438015
         RowHeight       =   503
         Columns(0).Width=   3200
         _ExtentX        =   5644
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "- Pilih Nama Dealer -"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Dealer :  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   15
         Top             =   645
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   1560
         Width           =   60
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   60
      End
   End
   Begin VB.CommandButton TblSimpan 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Cetak"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1890
   End
   Begin VB.CommandButton tblRefresh 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Refresh"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1890
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   9075
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   16007
      StartColor      =   33023
      EndColor        =   12648447
      FormatString    =   "   Laporan Ratention Rate"
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
      Width           =   18540
      _ExtentX        =   32703
      _ExtentY        =   503
      BackColor       =   255
      HighlightColor  =   255
      ShadowColor     =   65535
   End
End
Attribute VB_Name = "frmRepoSBRateHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim Report As Object
Private rsTemp As ADODB.Recordset
Private KodeDealer As String
Private namaDealer As String
Private aConn As New ADODB.Connection
Private database As String

Private Sub DCKodeGroup_Change()
    'dealer = ip dealer
    'dbase  = database dealer
    KodeDealer = ""
    database = ""
    namaDealer = ""
    
    namaDealer = DCKodeGroup.Columns(1).value
    KodeDealer = DCKodeGroup.Columns(0).value
    database = DCKodeGroup.Columns(3).value
    
End Sub

Private Sub DCKodeGroup_Click()
    
    Call DCKodeGroup_Change
End Sub

Private Sub Form_Load()
    Call FetchCombo
    Tgl1 = isnow
    Tgl2 = isnow
End Sub

Private Sub FetchCombo()
    AdcKode.ConnectionString = SQLConHO
    AdcKode.RecordSource = "select FS_KD_DEALER, FS_NM_DEALER,VPNCABANG,DBASECBG,USERCBG,PASSCBG,NEWENKRIP from t_dealer WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1"
    DCKodeGroup.DataSourceList = AdcKode
    DCKodeGroup.DataFieldList = "FS_NM_DEALER"
    DCKodeGroup.Columns(0).Width = 1200
    DCKodeGroup.Columns(1).Width = 4000
    DCKodeGroup.Columns(0).Caption = "Kode Dealer"
    DCKodeGroup.Columns(1).Caption = "Nama Dealer"
    DCKodeGroup.Columns(2).Visible = False
    DCKodeGroup.Columns(3).Visible = False
    DCKodeGroup.Columns(4).Visible = False
    DCKodeGroup.Columns(5).Visible = False
    DCKodeGroup.Columns(6).Visible = False
End Sub

Private Sub cmdCetak_Click()
          
1        On Error GoTo cmdCetak_Click_Error
2         If Month(Tgl1) <> Month(Tgl2) Or year(Tgl1) <> year(Tgl2) Then
3             MsgBox "Range periode harus 1 bulan", vbInformation
4             Exit Sub
5         End If
          
6         If KodeDealer = "" Then
7             MsgBox "Silahkan pilih dealer terlebih dahulu.", vbInformation
8             Exit Sub
9         End If
          
10        Screen.MousePointer = vbHourglass
11        cmdCetak.Enabled = False
12        ProgressBar1.Min = 0
13        ProgressBar1.Max = 10
14        ProgressBar1.Visible = True
15        ProgressBar1.value = 1
16        DoEvents
17        Set Report = New rptRepoSBRateHO
          
          Dim oRs As ADODB.Recordset
          Dim rsTempAll As ADODB.Recordset
          Dim rsTemp As ADODB.Recordset
          
          Dim konter As Integer
          Dim isTambah As Boolean
          Dim TotalTemp As Double
          
          Dim total1 As Double
          Dim total2 As Double
          Dim total3 As Double
          Dim total4 As Double
          Dim total5 As Double
          
          Dim Demand1 As Double
          Dim Demand2 As Double
          Dim Demand3 As Double
          Dim Demand4 As Double
          Dim Demand5 As Double
          
          Dim SM1 As Double
          Dim SM2 As Double
          Dim SM3 As Double
          Dim SM4 As Double
          Dim SM5 As Double
          
          Dim NCM1 As Double
          Dim NCM2 As Double
          Dim NCM3 As Double
          Dim NCM4 As Double
          Dim NCM5 As Double
          
          Dim NonNCM1 As Double
          Dim NonNCM2 As Double
          Dim NonNCM3 As Double
          Dim NonNCM4 As Double
          Dim NonNCM5 As Double
          
          Dim Cetak As Double
          'Call createRS
          
          Dim i As Integer
              
18            Demand1 = 0
19            Demand2 = 0
20            Demand3 = 0
21            Demand4 = 0
22            Demand5 = 0
              
23            SM1 = 0
24            SM2 = 0
25            SM3 = 0
26            SM4 = 0
27            SM5 = 0
              
28            NCM1 = 0
29            NCM2 = 0
30            NCM3 = 0
31            NCM4 = 0
32            NCM5 = 0
              
33            NonNCM1 = 0
34            NonNCM2 = 0
35            NonNCM3 = 0
36            NonNCM4 = 0
37            NonNCM5 = 0
              
             ' lblPengambilanDataRet.Visible = True
              ''proses pengambilan data same outlate
              
38            sSQL = "select COUNT(noinv) total,fs_rc"
39            sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK) "
40            sSQL = sSQL & " where NOINV <> '' AND FS_KD_DEALER = '" & KodeDealer & "' AND FS_NO_RETUR = '' "
41            sSQL = sSQL & " and FS_RC <> '' and noinv not like '%-B' and FS_RC in ('R010','R020','R030','R040','R050') "
42            sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'"
43            sSQL = sSQL & " grOUP BY FS_RC"
44            sSQL = sSQL & " order by FS_RC"
45            Set rsTemp = objO.FetchDataHO(sSQL)
46            Cetak = 0
47            ProgressBar1.value = 2
48            DoEvents
49            If Not rsTemp.EOF Then
                  ''SBE 10
50                rsTemp.filter = "fs_rc = 'R010'"
51                Cetak = 0
52                If Not rsTemp.EOF Then
53                    Cetak = CDbl(rsTemp("TOTAL"))
54                End If
55                rsTemp.filter = ""
56                Report.txtSO1.SetText Cetak
57                total1 = total1 + Cetak
58                SM1 = SM1 + Cetak
                  
                  ''SBE 20
59                rsTemp.filter = "fs_rc = 'R020'"
60                Cetak = 0
61                If Not rsTemp.EOF Then
62                    Cetak = CDbl(rsTemp("TOTAL"))
63                End If
64                rsTemp.filter = ""
65                Report.txtSO2.SetText Cetak
66                total2 = total2 + Cetak
67                SM2 = SM2 + Cetak
                  
                  ''SBE 30
68                rsTemp.filter = "fs_rc = 'R030'"
69                Cetak = 0
70                If Not rsTemp.EOF Then
71                    Cetak = CDbl(rsTemp("TOTAL"))
72                End If
73                rsTemp.filter = ""
74                Report.txtSO3.SetText Cetak
75                total3 = total3 + Cetak
76                SM3 = SM3 + Cetak
                  
                  ''SBE 40
77                rsTemp.filter = "fs_rc = 'R040'"
78                Cetak = 0
79                If Not rsTemp.EOF Then
80                    Cetak = CDbl(rsTemp("TOTAL"))
81                End If
82                rsTemp.filter = ""
83                Report.txtSO4.SetText Cetak
84                total4 = total4 + Cetak
85                SM4 = SM4 + Cetak
                  
                  ''SBE 50
86                rsTemp.filter = "fs_rc = 'R050'"
87                Cetak = 0
88                If Not rsTemp.EOF Then
89                    Cetak = CDbl(rsTemp("TOTAL"))
90                End If
91                rsTemp.filter = ""
92                Report.txtSO5.SetText Cetak
93                total5 = total5 + Cetak
94                SM5 = SM5 + Cetak
95            Else
96                Report.txtSO1.SetText Cetak
97                Report.txtSO2.SetText Cetak
98                Report.txtSO3.SetText Cetak
99                Report.txtSO4.SetText Cetak
100               Report.txtSO5.SetText Cetak
101           End If
              
102           sSQL = "select fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
103           sSQL = sSQL & " from T_INVOICEnew A WITH (NOLOCK)"
104           sSQL = sSQL & " left join  OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKompSPK & ";Uid=" & Log.LogHOUserSPK & ";PWD=" & Log.LogHOPassSPK & "',[" & Log.LogHODbSPK & "].dbo.SUMMARY) B   ON A.NO_CHASSIS = B.NORANGKA"
105           sSQL = sSQL & " where invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
106           sSQL = sSQL & " and fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC < 'R055' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
107           sSQL = sSQL & " group by fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end"
108           sSQL = sSQL & " order by fs_Rc"
109           Set rsTemp = objO.FetchDataHO(sSQL)
110           ProgressBar1.value = 3
111           DoEvents
112           Cetak = 0
113           If Not rsTemp.EOF Then
                  ''SBE 10
114               rsTemp.filter = "fs_rc = 'R010' AND STATUS = 'NCM'"
115               Cetak = 0
116               If Not rsTemp.EOF Then
117                   Cetak = CDbl(rsTemp("unit"))
118               End If
119               Report.txtNCM1.SetText Cetak
120               total1 = total1 + Cetak
121               NCM1 = NCM1 + Cetak
122               rsTemp.filter = ""
                  
123               rsTemp.filter = "fs_rc = 'R010' AND STATUS = 'Non NCM'"
124               Cetak = 0
125               If Not rsTemp.EOF Then
126                   Cetak = CDbl(rsTemp("unit"))
127               End If
128               Report.txtNonNCM1.SetText Cetak
129               total1 = total1 + Cetak
130               NonNCM1 = NonNCM1 + Cetak
131               rsTemp.filter = ""
                  
                  ''SBE 20
132               rsTemp.filter = "fs_rc = 'R020' AND STATUS = 'NCM'"
133               Cetak = 0
134               If Not rsTemp.EOF Then
135                   Cetak = CDbl(rsTemp("unit"))
136               End If
137               Report.txtNCM2.SetText Cetak
138               total2 = total2 + Cetak
139               NCM2 = NCM2 + Cetak
140               rsTemp.filter = ""
                  
141               rsTemp.filter = "fs_rc = 'R020' AND STATUS = 'Non NCM'"
142               Cetak = 0
143               If Not rsTemp.EOF Then
144                   Cetak = CDbl(rsTemp("unit"))
145               End If
146               Report.txtNonNCM2.SetText Cetak
147               total2 = total2 + Cetak
148               NonNCM2 = NonNCM2 + Cetak
149               rsTemp.filter = ""
                  
                  ''SBE 30
150               rsTemp.filter = "fs_rc = 'R030' AND STATUS = 'NCM'"
151               Cetak = 0
152               If Not rsTemp.EOF Then
153                   Cetak = CDbl(rsTemp("unit"))
154               End If
155               Report.txtNCM3.SetText Cetak
156               total3 = total3 + Cetak
157               NCM3 = NCM3 + Cetak
158               rsTemp.filter = "" '
                  
159               rsTemp.filter = "fs_rc = 'R030' AND STATUS = 'Non NCM'"
160               Cetak = 0
161               If Not rsTemp.EOF Then
162                   Cetak = CDbl(rsTemp("unit"))
163               End If
164               Report.txtNonNCM3.SetText Cetak
165               total3 = total3 + Cetak
166               NonNCM3 = NonNCM3 + Cetak
167               rsTemp.filter = ""
                  
                  ''SBE 40
168               rsTemp.filter = "fs_rc = 'R040' AND STATUS = 'NCM'"
169               Cetak = 0
170               If Not rsTemp.EOF Then
171                   Cetak = CDbl(rsTemp("unit"))
172               End If
173               Report.txtNCM4.SetText Cetak
174               total4 = total4 + Cetak
175               NCM4 = NCM4 + Cetak
176               rsTemp.filter = ""
                  
177               rsTemp.filter = "fs_rc = 'R040' AND STATUS = 'Non NCM'"
178               Cetak = 0
179               If Not rsTemp.EOF Then
180                   Cetak = CDbl(rsTemp("unit"))
181               End If
182               Report.txtNonNCM4.SetText Cetak
183               total4 = total4 + Cetak
184               NonNCM4 = NonNCM4 + Cetak
185               rsTemp.filter = ""
                  
                  ''SBE 50
186               rsTemp.filter = "fs_rc = 'R050' AND STATUS = 'NCM'"
187               Cetak = 0
188               If Not rsTemp.EOF Then
189                   Cetak = CDbl(rsTemp("unit"))
190               End If
191               Report.txtNCM5.SetText Cetak
192               total5 = total5 + Cetak
193               NCM5 = NCM5 + Cetak
194               rsTemp.filter = ""
                  
195               rsTemp.filter = "fs_rc = 'R050' AND STATUS = 'Non NCM'"
196               Cetak = 0
197               If Not rsTemp.EOF Then
198                   Cetak = CDbl(rsTemp("unit"))
199               End If
200               Report.txtNonNCM5.SetText Cetak
201               total5 = total5 + Cetak
202               NonNCM5 = NonNCM5 + Cetak
203               rsTemp.filter = ""
204           Else
205               Report.txtNCM1.SetText Cetak
206               Report.txtNonNCM1.SetText Cetak
207               Report.txtNCM2.SetText Cetak
208               Report.txtNonNCM2.SetText Cetak
209               Report.txtNCM3.SetText Cetak
210               Report.txtNonNCM3.SetText Cetak
211               Report.txtNCM4.SetText Cetak
212               Report.txtNonNCM4.SetText Cetak
213               Report.txtNCM5.SetText Cetak
214               Report.txtNonNCM5.SetText Cetak
215           End If
216           ProgressBar1.value = 4
217           DoEvents
218           Report.txtS41.SetText total1
219           Report.txtS42.SetText total2
220           Report.txtS43.SetText total3
221           Report.txtS44.SetText total4
222           Report.txtS45.SetText total5
              '' demand 1
223           If aConn.State = adStateClosed Then aConn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogHOUserSPK & ";password=" & Log.LogHOPassSPK & ";Initial Catalog=" & Log.LogHODbSPK & " ;server=" & Log.LogHOKompSPK
              
224           Set rsTemp = New ADODB.Recordset
225           rsTemp.CursorLocation = adUseClient
226           rsTemp.Open "select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a", aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
227           Set rsTemp.ActiveConnection = Nothing
228           If rsTemp.RecordCount > 0 Then
229               Demand1 = CDbl(rsTemp.Fields(0).value)
230           Else
231               Demand1 = 0
232           End If
233           If aConn.State = adStateOpen Then aConn.Close
              ''-------------
              
              '' demand 2
234           If aConn.State = adStateClosed Then aConn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogHOUserSPK & ";password=" & Log.LogHOPassSPK & ";Initial Catalog=" & Log.LogHODbSPK & " ;server=" & Log.LogHOKompSPK
              
235           Set rsTemp = New ADODB.Recordset
236           rsTemp.CursorLocation = adUseClient
237           rsTemp.Open "select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a", aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
238           Set rsTemp.ActiveConnection = Nothing
239           If rsTemp.RecordCount > 0 Then
240               Demand2 = CDbl(rsTemp.Fields(0).value)
241           Else
242               Demand2 = 0
243           End If
244           If aConn.State = adStateOpen Then aConn.Close
              
245           If aConn.State = adStateClosed Then aConn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogHOUserSPK & ";password=" & Log.LogHOPassSPK & ";Initial Catalog=" & Log.LogHODbSPK & " ;server=" & Log.LogHOKompSPK
              
246           Set rsTemp = New ADODB.Recordset
247           rsTemp.CursorLocation = adUseClient
248           rsTemp.Open "select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a", aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
249           Set rsTemp.ActiveConnection = Nothing
250           If rsTemp.RecordCount > 0 Then
251               Demand3 = CDbl(rsTemp.Fields(0).value)
252           Else
253               Demand3 = 0
254           End If
255           If aConn.State = adStateOpen Then aConn.Close
              
256           If aConn.State = adStateClosed Then aConn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogHOUserSPK & ";password=" & Log.LogHOPassSPK & ";Initial Catalog=" & Log.LogHODbSPK & " ;server=" & Log.LogHOKompSPK
              
257           Set rsTemp = New ADODB.Recordset
258           rsTemp.CursorLocation = adUseClient
259           rsTemp.Open "select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a", aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
260           Set rsTemp.ActiveConnection = Nothing
261           If rsTemp.RecordCount > 0 Then
262               Demand4 = CDbl(rsTemp.Fields(0).value)
263           Else
264               Demand4 = 0
265           End If
266           If aConn.State = adStateOpen Then aConn.Close
              
267           If aConn.State = adStateClosed Then aConn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogHOUserSPK & ";password=" & Log.LogHOPassSPK & ";Initial Catalog=" & Log.LogHODbSPK & " ;server=" & Log.LogHOKompSPK
              
268           Set rsTemp = New ADODB.Recordset
269           rsTemp.CursorLocation = adUseClient
270           rsTemp.Open "select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a", aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
271           Set rsTemp.ActiveConnection = Nothing
272           If rsTemp.RecordCount > 0 Then
273               Demand5 = CDbl(rsTemp.Fields(0).value)
274           Else
275               Demand5 = 0
276           End If
277           If aConn.State = adStateOpen Then aConn.Close
              
              'Demand1 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '        Demand2 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '        Demand3 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '        Demand4 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '        Demand5 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
              
278           Report.txtS61.SetText Demand1
279           Report.txtS62.SetText Demand2
280           Report.txtS63.SetText Demand3
281           Report.txtS64.SetText Demand4
282           Report.txtS65.SetText Demand5
283           ProgressBar1.value = 5
284           DoEvents
              ''ratention rate sameoutlate
285           If Demand1 = 0 Then
286             Report.txtS81.SetText 0 & " %"
287           Else
288             Report.txtS81.SetText Round(((SM1 / Demand1) * 100), 0) & " %"
289           End If
290           If Demand2 = 0 Then
291             Report.txtS82.SetText 0 & " %"
292           Else
293             Report.txtS82.SetText Round(((SM2 / Demand2) * 100), 0) & " %"
294           End If
295           If Demand3 = 0 Then
296             Report.txtS83.SetText 0 & " %"
297           Else
298             Report.txtS83.SetText Round(((SM3 / Demand3) * 100), 0) & " %"
299           End If
300           If Demand4 = 0 Then
301             Report.txtS84.SetText 0 & " %"
302           Else
303             Report.txtS84.SetText Round(((SM4 / Demand4) * 100), 0) & " %"
304           End If
305           If Demand5 = 0 Then
306             Report.txtS85.SetText 0 & " %"
307           Else
308             Report.txtS85.SetText Round(((SM5 / Demand5) * 100), 0) & " %"
309           End If
              '' MENGGUNAKAN IIF TIDAK BISA DIGUNAKAN
      '
      '        Report.txtS81.SetText IIf(Demand1 = 0, 0, Round(((SM1 / Demand1) * 100), 0)) & " %"
      '        Report.txtS82.SetText IIf(Demand2 = 0, 0, Round(((SM2 / Demand2) * 100), 0)) & " %"
      '        Report.txtS83.SetText IIf(Demand3 = 0, 0, Round(((SM3 / Demand3) * 100), 0)) & " %"
      '        Report.txtS84.SetText IIf(Demand4 = 0, 0, Round(((SM4 / Demand4) * 100), 0)) & " %"
      '        Report.txtS85.SetText IIf(Demand5 = 0, 0, Round(((SM5 / Demand5) * 100), 0)) & " %"
              ''--------------------------
              
              ''total ratention rate
310           If Demand1 = 0 Then
311               Report.txtS91.SetText 0 & " %"
312           Else
313               Report.txtS91.SetText Round(((total1 / Demand1) * 100), 0) & " %"
314           End If
315           If Demand2 = 0 Then
316               Report.txtS92.SetText 0 & " %"
317           Else
318               Report.txtS92.SetText Round(((total2 / Demand2) * 100), 0) & " %"
319           End If
320           If Demand3 = 0 Then
321               Report.txtS93.SetText 0 & " %"
322           Else
323               Report.txtS93.SetText Round(((total3 / Demand3) * 100), 0) & " %"
324           End If
325           If Demand4 = 0 Then
326               Report.txtS94.SetText 0 & " %"
327           Else
328               Report.txtS94.SetText Round(((total4 / Demand4) * 100), 0) & " %"
329           End If
330           If Demand5 = 0 Then
331               Report.txtS95.SetText 0 & " %"
332           Else
333               Report.txtS95.SetText Round(((total5 / Demand5) * 100), 0) & " %"
334           End If
              ''
      '        Report.txtS91.SetText IIf(Demand1 = 0, 0, Round(((total1 / Demand1) * 100), 0)) & " %"
      '        Report.txtS92.SetText IIf(Demand2 = 0, 0, Round(((total2 / Demand2) * 100), 0)) & " %"
      '        Report.txtS93.SetText IIf(Demand3 = 0, 0, Round(((total3 / Demand3) * 100), 0)) & " %"
      '        Report.txtS94.SetText IIf(Demand4 = 0, 0, Round(((total4 / Demand4) * 100), 0)) & " %"
      '        Report.txtS95.SetText IIf(Demand5 = 0, 0, Round(((total5 / Demand5) * 100), 0)) & " %"
              ''--------------------
              
              ''Non NCM / Total
335           If total1 = 0 Then
336               Report.txtS121.SetText 0 & " %"
337           Else
338               Report.txtS121.SetText Round(((NonNCM1 / total1) * 100), 0) & " %"
339           End If
340           If total2 = 0 Then
341               Report.txtS122.SetText 0 & " %"
342           Else
343               Report.txtS122.SetText Round(((NonNCM2 / total2) * 100), 0) & " %"
344           End If
345           If total3 = 0 Then
346               Report.txtS123.SetText 0 & " %"
347           Else
348               Report.txtS123.SetText Round(((NonNCM3 / total3) * 100), 0) & " %"
349           End If
350           If total4 = 0 Then
351               Report.txtS124.SetText 0 & " %"
352           Else
353               Report.txtS124.SetText Round(((NonNCM4 / total4) * 100), 0) & " %"
354           End If
355           If total5 = 0 Then
356               Report.txtS125.SetText 0 & " %"
357           Else
358               Report.txtS125.SetText Round(((NonNCM5 / total5) * 100), 0) & " %"
359           End If
      '        Report.txtS121.SetText IIf(total1 = 0, 0, Round(((NonNCM1 / total1) * 100), 0)) & " %"
      '        Report.txtS122.SetText IIf(total2 = 0, 0, Round(((NonNCM2 / total2) * 100), 0)) & " %"
      '        Report.txtS123.SetText IIf(total3 = 0, 0, Round(((NonNCM3 / total3) * 100), 0)) & " %"
      '        Report.txtS124.SetText IIf(total4 = 0, 0, Round(((NonNCM4 / total4) * 100), 0)) & " %"
      '        Report.txtS125.SetText IIf(total5 = 0, 0, Round(((NonNCM5 / total5) * 100), 0)) & " %"
              ''---------------
              
360           Report.txtSMTotal.SetText (SM1 + SM2 + SM3 + SM4 + SM5)
361           Report.Text10.SetText (NCM1 + NCM2 + NCM3 + NCM4 + NCM5)
362           Report.Text13.SetText (NonNCM1 + NonNCM2 + NonNCM3 + NonNCM4 + NonNCM5)
363           Report.Text14.SetText (total1 + total2 + total3 + total4 + total5)
              ''Report.Text15.SetText Round((Demand1 + Demand2 + Demand3 + Demand4 + Demand5) / 5, 0)
              
364           Report.Text15.SetText Round((Demand1 + Demand2 + Demand3 + Demand4 + Demand5), 0)
              
365           If Round((Demand1 + Demand2 + Demand3 + Demand4 + Demand5), 0) > 0 Then
366               Report.Text17.SetText Round(((SM1 + SM2 + SM3 + SM4 + SM5) / (Demand1 + Demand2 + Demand3 + Demand4 + Demand5)) * 100, 0) & " %"
367               Report.Text21.SetText Round(((total1 + total2 + total3 + total4 + total5) / (Demand1 + Demand2 + Demand3 + Demand4 + Demand5)) * 100, 0) & " %"
368           Else
369               Report.Text17.SetText "0 %"
370               Report.Text21.SetText "0 %"
371           End If
372           ProgressBar1.value = 6
373           DoEvents
374           TotalTemp = 0
375           If Demand1 > 0 Then
376               TotalTemp = TotalTemp + Round(((SM1 / Demand1) * 100), 0)
377           End If
378           If Demand2 > 0 Then
379               TotalTemp = TotalTemp + Round(((SM2 / Demand2) * 100), 0)
380           End If
381           If Demand3 > 0 Then
382               TotalTemp = TotalTemp + Round(((SM3 / Demand3) * 100), 0)
383           End If
384           If Demand4 > 0 Then
385               TotalTemp = TotalTemp + Round(((SM4 / Demand4) * 100), 0)
386           End If
387           If Demand5 > 0 Then
388               TotalTemp = TotalTemp + Round(((SM5 / Demand5) * 100), 0)
389           End If
              ''Report.Text17.SetText Round(TotalTemp / 5, 0) & " %"
              'Report.Text17.SetText Round((((IIf(Demand1 = 0, 0, Round(((SM1 / Demand1) * 100), 0))) + (IIf(Demand2 = 0, 0, Round(((SM2 / Demand2) * 100), 0))) + (IIf(Demand3 = 0, 0, Round(((SM3 / Demand3) * 100), 0))) + (IIf(Demand4 = 0, 0, Round(((SM4 / Demand4) * 100), 0))) + (IIf(Demand5 = 0, 0, Round(((SM5 / Demand5) * 100), 0)))) / 5), 0) & " %"
              
390           TotalTemp = 0
391           If Demand1 > 0 Then
392               TotalTemp = TotalTemp + Round(((total1 / Demand1) * 100), 0)
393           End If
394           If Demand2 > 0 Then
395               TotalTemp = TotalTemp + Round(((total2 / Demand2) * 100), 0)
396           End If
397           If Demand3 > 0 Then
398               TotalTemp = TotalTemp + Round(((total3 / Demand3) * 100), 0)
399           End If
400           If Demand4 > 0 Then
401               TotalTemp = TotalTemp + Round(((total4 / Demand4) * 100), 0)
402           End If
403           If Demand5 > 0 Then
404               TotalTemp = TotalTemp + Round(((total5 / Demand5) * 100), 0)
405           End If
              ''Report.Text21.SetText Round(TotalTemp / 5, 0) & " %"
              'Report.Text21.SetText Round((((IIf(Demand1 = 0, 0, Round(((total1 / Demand1) * 100), 0))) + (IIf(Demand2 = 0, 0, Round(((total2 / Demand2) * 100), 0))) + (IIf(Demand3 = 0, 0, Round(((total3 / Demand3) * 100), 0))) + (IIf(Demand4 = 0, 0, Round(((total4 / Demand4) * 100), 0))) + (IIf(Demand5 = 0, 0, Round(((total5 / Demand5) * 100), 0)))) / 5), 0) & " %"
              
406           TotalTemp = 0
407           If total1 > 0 Then
408               TotalTemp = TotalTemp + Round(((NonNCM1 / total1) * 100), 0)
409           End If
410           If total2 > 0 Then
411               TotalTemp = TotalTemp + Round(((NonNCM2 / total2) * 100), 0)
412           End If
413           If total3 > 0 Then
414               TotalTemp = TotalTemp + Round(((NonNCM3 / total3) * 100), 0)
415           End If
416           If total4 > 0 Then
417               TotalTemp = TotalTemp + Round(((NonNCM4 / total4) * 100), 0)
418           End If
419           If total5 > 0 Then
420               TotalTemp = TotalTemp + Round(((NonNCM5 / total5) * 100), 0)
421           End If
422           Report.Text23.SetText Round(TotalTemp / 5, 0) & " %"
              
              'Report.Text23.SetText Round((((IIf(total1 = 0, 0, Round(((NonNCM1 / total1) * 100), 0))) + (IIf(total2 = 0, 0, Round(((NonNCM2 / total2) * 100), 0))) + (IIf(total3 = 0, 0, Round(((NonNCM3 / total3) * 100), 0))) + (IIf(total4 = 0, 0, Round(((NonNCM4 / total4) * 100), 0))) + (IIf(total5 = 0, 0, Round(((NonNCM5 / total5) * 100), 0)))) / 5), 0) & " %"
423           ProgressBar1.value = 7
424           DoEvents
425           If aConn.State = adStateClosed Then aConn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogHOUserSPK & ";password=" & Log.LogHOPassSPK & ";Initial Catalog=" & Log.LogHODbSPK & " ;server=" & Log.LogHOKompSPK
              
426           Set rsTemp = New ADODB.Recordset
427           rsTemp.CursorLocation = adUseClient
428           rsTemp.Open "select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(YEAR,-8,'" & Format(Tgl1, "MM/DD/YYYY") & "') and '" & Format(Tgl2, "MM/DD/YYYY") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a", aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
429           Set rsTemp.ActiveConnection = Nothing
430           If rsTemp.RecordCount > 0 Then
431               Demand1 = CDbl(rsTemp.Fields(0).value)
432           Else
433               Demand1 = 0
434           End If
435           If aConn.State = adStateOpen Then aConn.Close
              
436           Demand1 = Round(Demand1 * 2.2, 0)
437           Report.Text47.SetText Demand1
              
438           ProgressBar1.value = 8
439           DoEvents
              
              ''total share
440           sSQL = "select COUNT(noinv) total"
441           sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK) "
442           sSQL = sSQL & " where noinv not like '%-B' and NOINV <> '' AND FS_KD_DEALER = '" & KodeDealer & "' AND FS_NO_RETUR = '' "
443           sSQL = sSQL & " AND FS_RC <> 'R001' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0 AND INVDATE between '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'"
              
444           Set rsTemp = objO.FetchData(sSQL)
              
445           If Not rsTemp.EOF Then
446               total1 = CDbl(rsTemp!total)
447           Else
448               total1 = 0
449           End If
              
450           Report.Text43.SetText total1
451           Report.Text11.SetText "LAPORAN RETENTION RATE " & namaDealer
              
452           sSQL = "select case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
453           sSQL = sSQL & " from T_INVOICEnew A WITH (NOLOCK)"
454           sSQL = sSQL & " left join  OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKompSPK & ";Uid=" & Log.LogHOUserSPK & ";PWD=" & Log.LogHOPassSPK & "',[" & Log.LogHODbSPK & "].dbo.SUMMARY) B   ON A.NO_CHASSIS = B.NORANGKA"
455           sSQL = sSQL & " where noinv not like '%-B' and invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
456           sSQL = sSQL & " and fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC < 'R055' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
457           sSQL = sSQL & " group by case when b.norangka is null then 'Non NCM' else 'NCM' end"
              
458           Set rsTemp = objO.FetchData(sSQL)
459           NCM1 = 0
460           NonNCM1 = 0
461           If Not rsTemp.EOF Then
                  ''ncm
462               rsTemp.filter = "status = 'NCM'"
463               If Not rsTemp.EOF Then
464                   NCM1 = CDbl(rsTemp!unit)
465               End If
466               rsTemp.filter = ""
                  ''non ncm
                  
467               rsTemp.filter = "status = 'Non NCM'"
468               If Not rsTemp.EOF Then
469                   NonNCM1 = CDbl(rsTemp!unit)
470               End If
471               rsTemp.filter = ""
                  
472           End If
473           Report.Text44.SetText NCM1
474           Report.Text45.SetText NonNCM1
              
              ''total share
475           total2 = 0
476           total2 = total1 + NCM1 + NonNCM1
477           Report.Text46.SetText total2
              
478           ProgressBar1.value = 9
479           DoEvents
              
480           If Demand1 > 0 Then
481               Report.Text48.SetText Round(((total1 / Demand1) * 100), 0) & " %"
482               Report.Text49.SetText Round(((total2 / Demand1) * 100), 0) & " %"
483           Else
484               Report.Text48.SetText "0 %"
485               Report.Text49.SetText "0 %"
486           End If
              
487           If total2 > 0 Then
488               Report.Text50.SetText Round(((NonNCM1 / total2) * 100), 0) & " %"
489           Else
490               Report.Text50.SetText "0 %"
              
491           End If
              
492           ProgressBar1.value = 10
493           DoEvents
494           ProgressBar1.Visible = False
495           cmdCetak.Enabled = True
496       With CRViewer1
497           .ReportSource = Report
498           .ViewReport
499           .Zoom 100
500           .Visible = True
501           .Left = 300
502           .Height = Me.Height - 400
503           .Width = Me.Width - .Left
504       End With
         
505               cmdTutup.Visible = True
506               cmdTutup.Left = CRViewer1.Left + 20
507               cmdTutup.Top = CRViewer1.Top + 400
                  
          'ProgressBar2.Visible = False
          'Frame2.Visible = False
508       Screen.MousePointer = vbDefault
          
509      Exit Sub

cmdCetak_Click_Error:
510       Screen.MousePointer = vbDefault
511       cmdCetak.Enabled = True
512       ProgressBar1.Visible = True
513       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCetak_Click of Form frmRepoSBRateHO")
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub Form_Activate()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
    asxPowerBanner1.Height = Me.Height - 500
End Sub

