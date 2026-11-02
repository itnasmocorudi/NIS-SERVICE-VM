VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoSBRate 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form Ratantion Rate"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1740
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2685
      Width           =   1440
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H000000FF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2685
      Width           =   1050
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5550
      TabIndex        =   15
      Top             =   2715
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdPrintOut 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8670
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2685
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3225
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2685
      Visible         =   0   'False
      Width           =   1185
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   4440
      TabIndex        =   16
      Top             =   2700
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
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
      MICON           =   "frmRepoSBRate.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   0
      Top             =   2400
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
      TabIndex        =   4
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   375
         Left            =   1650
         TabIndex        =   5
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
         Format          =   125042689
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   375
         Left            =   3240
         TabIndex        =   6
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
         Format          =   125042689
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblAmbilData 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   6255
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         Left            =   3120
         TabIndex        =   7
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
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
Attribute VB_Name = "frmRepoSBRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
 Dim Report As Object
Private rsTemp As ADODB.Recordset

Private Sub cmdCetak_Click()
1         If Month(Tgl1) <> Month(Tgl2) Or year(Tgl2) <> year(Tgl2) Then
2             MsgBox "Range periode dalam 1 bulan", vbInformation
3             Exit Sub
4         End If
5        On Error GoTo cmdCetak_Click_Error
6         Screen.MousePointer = vbHourglass
7         cmdCetak.Enabled = False
8         ProgressBar1.Min = 0
9         ProgressBar1.Max = 10
10        ProgressBar1.Visible = True
11        ProgressBar1.value = 1
12        DoEvents
13        Set Report = New rptRepoSBRate
          
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
              
14            Demand1 = 0
15            Demand2 = 0
16            Demand3 = 0
17            Demand4 = 0
18            Demand5 = 0
              
19            SM1 = 0
20            SM2 = 0
21            SM3 = 0
22            SM4 = 0
23            SM5 = 0
              
24            NCM1 = 0
25            NCM2 = 0
26            NCM3 = 0
27            NCM4 = 0
28            NCM5 = 0
              
29            NonNCM1 = 0
30            NonNCM2 = 0
31            NonNCM3 = 0
32            NonNCM4 = 0
33            NonNCM5 = 0
              
             ' lblPengambilanDataRet.Visible = True
              ''proses pengambilan data same outlate
34            If Log.RPLHistori = 0 Then
35                sSQL = "select COUNT(noinv) total,fs_rc"
36                sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
37                sSQL = sSQL & " INNER JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
38                sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
39                sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
40                sSQL = sSQL & " and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050') "
41                sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'"
42                sSQL = sSQL & " grOUP BY FS_RC"
43                sSQL = sSQL & " order by FS_RC"
44            Else
45                sSQL = "select COUNT(noinv) total,fs_rc"
46                sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
47                sSQL = sSQL & " INNER JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
48                sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
49                sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
50                sSQL = sSQL & " and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050') "
51                sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'"
52                sSQL = sSQL & " grOUP BY FS_RC"
53                sSQL = sSQL & " order by FS_RC"
54            End If
              
55            Set rsTemp = objO.FetchData(sSQL)
56            Cetak = 0
57            ProgressBar1.value = 2
58            DoEvents
59            If Not rsTemp.EOF Then
                  ''SBE 10
60                rsTemp.filter = "fs_rc = 'R010'"
61                Cetak = 0
62                If Not rsTemp.EOF Then
63                    Cetak = CDbl(rsTemp("TOTAL"))
64                End If
65                rsTemp.filter = ""
66                Report.txtSO1.SetText Cetak
67                total1 = total1 + Cetak
68                SM1 = SM1 + Cetak
                  
                  ''SBE 20
69                rsTemp.filter = "fs_rc = 'R020'"
70                Cetak = 0
71                If Not rsTemp.EOF Then
72                    Cetak = CDbl(rsTemp("TOTAL"))
73                End If
74                rsTemp.filter = ""
75                Report.txtSO2.SetText Cetak
76                total2 = total2 + Cetak
77                SM2 = SM2 + Cetak
                  
                  ''SBE 30
78                rsTemp.filter = "fs_rc = 'R030'"
79                Cetak = 0
80                If Not rsTemp.EOF Then
81                    Cetak = CDbl(rsTemp("TOTAL"))
82                End If
83                rsTemp.filter = ""
84                Report.txtSO3.SetText Cetak
85                total3 = total3 + Cetak
86                SM3 = SM3 + Cetak
                  
                  ''SBE 40
87                rsTemp.filter = "fs_rc = 'R040'"
88                Cetak = 0
89                If Not rsTemp.EOF Then
90                    Cetak = CDbl(rsTemp("TOTAL"))
91                End If
92                rsTemp.filter = ""
93                Report.txtSO4.SetText Cetak
94                total4 = total4 + Cetak
95                SM4 = SM4 + Cetak
                  
                  ''SBE 50
96                rsTemp.filter = "fs_rc = 'R050'"
97                Cetak = 0
98                If Not rsTemp.EOF Then
99                    Cetak = CDbl(rsTemp("TOTAL"))
100               End If
101               rsTemp.filter = ""
102               Report.txtSO5.SetText Cetak
103               total5 = total5 + Cetak
104               SM5 = SM5 + Cetak
105           Else
106               Report.txtSO1.SetText Cetak
107               Report.txtSO2.SetText Cetak
108               Report.txtSO3.SetText Cetak
109               Report.txtSO4.SetText Cetak
110               Report.txtSO5.SetText Cetak
111           End If
              
112           If Log.RPLHistori = 0 Then
113               sSQL = "select fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
114               sSQL = sSQL & " from t_invoice A WITH (NOLOCK)"
115               sSQL = sSQL & " left join SALESUNIT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NORANGKA"
116               sSQL = sSQL & " where invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
117               sSQL = sSQL & " and a.fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC < 'R055' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
118               sSQL = sSQL & " group by fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end"
119               sSQL = sSQL & " order by fs_Rc"
120           Else
121               sSQL = "select fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
122               sSQL = sSQL & " from t_invoice A WITH (NOLOCK)"
123               sSQL = sSQL & " left join [" & dbReplikasiHO & "]..SUMMARY B WITH (NOLOCK) ON A.NO_CHASSIS = B.NORANGKA"
124               sSQL = sSQL & " where invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
125               sSQL = sSQL & " and a.fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC < 'R055' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
126               sSQL = sSQL & " group by fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end"
127               sSQL = sSQL & " order by fs_Rc"
128           End If
129           Set rsTemp = objO.FetchData(sSQL)
130           ProgressBar1.value = 3
131           DoEvents
132           Cetak = 0
133           If Not rsTemp.EOF Then
                  ''SBE 10
134               rsTemp.filter = "fs_rc = 'R010' AND STATUS = 'NCM'"
135               Cetak = 0
136               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
137                   Cetak = CDbl(rsTemp("unit")) - SM1
138               End If
139               Report.txtNCM1.SetText Cetak
140               total1 = total1 + Cetak
141               NCM1 = NCM1 + Cetak
142               rsTemp.filter = ""
                  
143               rsTemp.filter = "fs_rc = 'R010' AND STATUS = 'Non NCM'"
144               Cetak = 0
145               If Not rsTemp.EOF Then
146                   Cetak = CDbl(rsTemp("unit"))
147               End If
148               Report.txtNonNCM1.SetText Cetak
149               total1 = total1 + Cetak
150               NonNCM1 = NonNCM1 + Cetak
151               rsTemp.filter = ""
                  
                  ''SBE 20
152               rsTemp.filter = "fs_rc = 'R020' AND STATUS = 'NCM'"
153               Cetak = 0
154               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
155                   Cetak = CDbl(rsTemp("unit")) - SM2
156               End If
157               Report.txtNCM2.SetText Cetak
158               total2 = total2 + Cetak
159               NCM2 = NCM2 + Cetak
160               rsTemp.filter = ""
                  
161               rsTemp.filter = "fs_rc = 'R020' AND STATUS = 'Non NCM'"
162               Cetak = 0
163               If Not rsTemp.EOF Then
164                   Cetak = CDbl(rsTemp("unit"))
165               End If
166               Report.txtNonNCM2.SetText Cetak
167               total2 = total2 + Cetak
168               NonNCM2 = NonNCM2 + Cetak
169               rsTemp.filter = ""
                  
                  ''SBE 30
170               rsTemp.filter = "fs_rc = 'R030' AND STATUS = 'NCM'"
171               Cetak = 0
172               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
173                   Cetak = CDbl(rsTemp("unit")) - SM3
174               End If
175               Report.txtNCM3.SetText Cetak
176               total3 = total3 + Cetak
177               NCM3 = NCM3 + Cetak
178               rsTemp.filter = "" '
                  
179               rsTemp.filter = "fs_rc = 'R030' AND STATUS = 'Non NCM'"
180               Cetak = 0
181               If Not rsTemp.EOF Then
182                   Cetak = CDbl(rsTemp("unit"))
183               End If
184               Report.txtNonNCM3.SetText Cetak
185               total3 = total3 + Cetak
186               NonNCM3 = NonNCM3 + Cetak
187               rsTemp.filter = ""
                  
                  ''SBE 40
188               rsTemp.filter = "fs_rc = 'R040' AND STATUS = 'NCM'"
189               Cetak = 0
190               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
191                   Cetak = CDbl(rsTemp("unit")) - SM4
192               End If
193               Report.txtNCM4.SetText Cetak
194               total4 = total4 + Cetak
195               NCM4 = NCM4 + Cetak
196               rsTemp.filter = ""
                  
197               rsTemp.filter = "fs_rc = 'R040' AND STATUS = 'Non NCM'"
198               Cetak = 0
199               If Not rsTemp.EOF Then
200                   Cetak = CDbl(rsTemp("unit"))
201               End If
202               Report.txtNonNCM4.SetText Cetak
203               total4 = total4 + Cetak
204               NonNCM4 = NonNCM4 + Cetak
205               rsTemp.filter = ""
                  
                  ''SBE 50
206               rsTemp.filter = "fs_rc = 'R050' AND STATUS = 'NCM'"
207               Cetak = 0
208               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
209                   Cetak = CDbl(rsTemp("unit")) - SM5
210               End If
211               Report.txtNCM5.SetText Cetak
212               total5 = total5 + Cetak
213               NCM5 = NCM5 + Cetak
214               rsTemp.filter = ""
                  
215               rsTemp.filter = "fs_rc = 'R050' AND STATUS = 'Non NCM'"
216               Cetak = 0
217               If Not rsTemp.EOF Then
218                   Cetak = CDbl(rsTemp("unit"))
219               End If
220               Report.txtNonNCM5.SetText Cetak
221               total5 = total5 + Cetak
222               NonNCM5 = NonNCM5 + Cetak
223               rsTemp.filter = ""
224           Else
225               Report.txtNCM1.SetText Cetak
226               Report.txtNonNCM1.SetText Cetak
227               Report.txtNCM2.SetText Cetak
228               Report.txtNonNCM2.SetText Cetak
229               Report.txtNCM3.SetText Cetak
230               Report.txtNonNCM3.SetText Cetak
231               Report.txtNCM4.SetText Cetak
232               Report.txtNonNCM4.SetText Cetak
233               Report.txtNCM5.SetText Cetak
234               Report.txtNonNCM5.SetText Cetak
235           End If
236           ProgressBar1.value = 4
237           DoEvents
238           Report.txtS41.SetText total1
239           Report.txtS42.SetText total2
240           Report.txtS43.SetText total3
241           Report.txtS44.SetText total4
242           Report.txtS45.SetText total5
              
243           Demand1 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
244           Demand2 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
245           Demand3 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
246           Demand4 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
247           Demand5 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
              
248           Report.txtS61.SetText Demand1
249           Report.txtS62.SetText Demand2
250           Report.txtS63.SetText Demand3
251           Report.txtS64.SetText Demand4
252           Report.txtS65.SetText Demand5
253           ProgressBar1.value = 5
254           DoEvents
              ''ratention rate sameoutlate
255           If Demand1 = 0 Then
256             Report.txtS81.SetText 0 & " %"
257           Else
258             Report.txtS81.SetText Round(((SM1 / Demand1) * 100), 0) & " %"
259           End If
260           If Demand2 = 0 Then
261             Report.txtS82.SetText 0 & " %"
262           Else
263             Report.txtS82.SetText Round(((SM2 / Demand2) * 100), 0) & " %"
264           End If
265           If Demand3 = 0 Then
266             Report.txtS83.SetText 0 & " %"
267           Else
268             Report.txtS83.SetText Round(((SM3 / Demand3) * 100), 0) & " %"
269           End If
270           If Demand4 = 0 Then
271             Report.txtS84.SetText 0 & " %"
272           Else
273             Report.txtS84.SetText Round(((SM4 / Demand4) * 100), 0) & " %"
274           End If
275           If Demand5 = 0 Then
276             Report.txtS85.SetText 0 & " %"
277           Else
278             Report.txtS85.SetText Round(((SM5 / Demand5) * 100), 0) & " %"
279           End If
              '' MENGGUNAKAN IIF TIDAK BISA DIGUNAKAN
      '
      '        Report.txtS81.SetText IIf(Demand1 = 0, 0, Round(((SM1 / Demand1) * 100), 0)) & " %"
      '        Report.txtS82.SetText IIf(Demand2 = 0, 0, Round(((SM2 / Demand2) * 100), 0)) & " %"
      '        Report.txtS83.SetText IIf(Demand3 = 0, 0, Round(((SM3 / Demand3) * 100), 0)) & " %"
      '        Report.txtS84.SetText IIf(Demand4 = 0, 0, Round(((SM4 / Demand4) * 100), 0)) & " %"
      '        Report.txtS85.SetText IIf(Demand5 = 0, 0, Round(((SM5 / Demand5) * 100), 0)) & " %"
              ''--------------------------
              
              ''total ratention rate
280           If Demand1 = 0 Then
281               Report.txtS91.SetText 0 & " %"
282           Else
283               Report.txtS91.SetText Round(((total1 / Demand1) * 100), 0) & " %"
284           End If
285           If Demand2 = 0 Then
286               Report.txtS92.SetText 0 & " %"
287           Else
288               Report.txtS92.SetText Round(((total2 / Demand2) * 100), 0) & " %"
289           End If
290           If Demand3 = 0 Then
291               Report.txtS93.SetText 0 & " %"
292           Else
293               Report.txtS93.SetText Round(((total3 / Demand3) * 100), 0) & " %"
294           End If
295           If Demand4 = 0 Then
296               Report.txtS94.SetText 0 & " %"
297           Else
298               Report.txtS94.SetText Round(((total4 / Demand4) * 100), 0) & " %"
299           End If
300           If Demand5 = 0 Then
301               Report.txtS95.SetText 0 & " %"
302           Else
303               Report.txtS95.SetText Round(((total5 / Demand5) * 100), 0) & " %"
304           End If
              ''
      '        Report.txtS91.SetText IIf(Demand1 = 0, 0, Round(((total1 / Demand1) * 100), 0)) & " %"
      '        Report.txtS92.SetText IIf(Demand2 = 0, 0, Round(((total2 / Demand2) * 100), 0)) & " %"
      '        Report.txtS93.SetText IIf(Demand3 = 0, 0, Round(((total3 / Demand3) * 100), 0)) & " %"
      '        Report.txtS94.SetText IIf(Demand4 = 0, 0, Round(((total4 / Demand4) * 100), 0)) & " %"
      '        Report.txtS95.SetText IIf(Demand5 = 0, 0, Round(((total5 / Demand5) * 100), 0)) & " %"
              ''--------------------
              
              ''Non NCM / Total
305           If total1 = 0 Then
306               Report.txtS121.SetText 0 & " %"
307           Else
308               Report.txtS121.SetText Round(((NonNCM1 / total1) * 100), 0) & " %"
309           End If
310           If total2 = 0 Then
311               Report.txtS122.SetText 0 & " %"
312           Else
313               Report.txtS122.SetText Round(((NonNCM2 / total2) * 100), 0) & " %"
314           End If
315           If total3 = 0 Then
316               Report.txtS123.SetText 0 & " %"
317           Else
318               Report.txtS123.SetText Round(((NonNCM3 / total3) * 100), 0) & " %"
319           End If
320           If total4 = 0 Then
321               Report.txtS124.SetText 0 & " %"
322           Else
323               Report.txtS124.SetText Round(((NonNCM4 / total4) * 100), 0) & " %"
324           End If
325           If total5 = 0 Then
326               Report.txtS125.SetText 0 & " %"
327           Else
328               Report.txtS125.SetText Round(((NonNCM5 / total5) * 100), 0) & " %"
329           End If
      '        Report.txtS121.SetText IIf(total1 = 0, 0, Round(((NonNCM1 / total1) * 100), 0)) & " %"
      '        Report.txtS122.SetText IIf(total2 = 0, 0, Round(((NonNCM2 / total2) * 100), 0)) & " %"
      '        Report.txtS123.SetText IIf(total3 = 0, 0, Round(((NonNCM3 / total3) * 100), 0)) & " %"
      '        Report.txtS124.SetText IIf(total4 = 0, 0, Round(((NonNCM4 / total4) * 100), 0)) & " %"
      '        Report.txtS125.SetText IIf(total5 = 0, 0, Round(((NonNCM5 / total5) * 100), 0)) & " %"
              ''---------------
              
330           Report.txtSMTotal.SetText (SM1 + SM2 + SM3 + SM4 + SM5)
331           Report.Text10.SetText (NCM1 + NCM2 + NCM3 + NCM4 + NCM5)
332           Report.Text13.SetText (NonNCM1 + NonNCM2 + NonNCM3 + NonNCM4 + NonNCM5)
333           Report.Text14.SetText (total1 + total2 + total3 + total4 + total5)
334           Report.Text15.SetText Round((Demand1 + Demand2 + Demand3 + Demand4 + Demand5) / 5, 0)
              ''Report.Text15.SetText Round((Demand1 + Demand2 + Demand3 + Demand4 + Demand5), 0)
              
335           If Round((Demand1 + Demand2 + Demand3 + Demand4 + Demand5), 0) > 0 Then
336               Report.Text17.SetText Round(((SM1 + SM2 + SM3 + SM4 + SM5) / (Demand1 + Demand2 + Demand3 + Demand4 + Demand5)) * 100, 0) & " %"
337               Report.Text21.SetText Round(((total1 + total2 + total3 + total4 + total5) / (Demand1 + Demand2 + Demand3 + Demand4 + Demand5)) * 100, 0) & " %"
338           Else
339               Report.Text17.SetText "0 %"
340               Report.Text21.SetText "0 %"
341           End If
              
342           ProgressBar1.value = 6
343           DoEvents
344           TotalTemp = 0
345           If Demand1 > 0 Then
346               TotalTemp = TotalTemp + Round(((SM1 / Demand1) * 100), 0)
347           End If
348           If Demand2 > 0 Then
349               TotalTemp = TotalTemp + Round(((SM2 / Demand2) * 100), 0)
350           End If
351           If Demand3 > 0 Then
352               TotalTemp = TotalTemp + Round(((SM3 / Demand3) * 100), 0)
353           End If
354           If Demand4 > 0 Then
355               TotalTemp = TotalTemp + Round(((SM4 / Demand4) * 100), 0)
356           End If
357           If Demand5 > 0 Then
358               TotalTemp = TotalTemp + Round(((SM5 / Demand5) * 100), 0)
359           End If
              ''Report.Text17.SetText Round(TotalTemp / 5, 0) & " %"
              'Report.Text17.SetText Round((((IIf(Demand1 = 0, 0, Round(((SM1 / Demand1) * 100), 0))) + (IIf(Demand2 = 0, 0, Round(((SM2 / Demand2) * 100), 0))) + (IIf(Demand3 = 0, 0, Round(((SM3 / Demand3) * 100), 0))) + (IIf(Demand4 = 0, 0, Round(((SM4 / Demand4) * 100), 0))) + (IIf(Demand5 = 0, 0, Round(((SM5 / Demand5) * 100), 0)))) / 5), 0) & " %"
              
360           TotalTemp = 0
361           If Demand1 > 0 Then
362               TotalTemp = TotalTemp + Round(((total1 / Demand1) * 100), 0)
363           End If
364           If Demand2 > 0 Then
365               TotalTemp = TotalTemp + Round(((total2 / Demand2) * 100), 0)
366           End If
367           If Demand3 > 0 Then
368               TotalTemp = TotalTemp + Round(((total3 / Demand3) * 100), 0)
369           End If
370           If Demand4 > 0 Then
371               TotalTemp = TotalTemp + Round(((total4 / Demand4) * 100), 0)
372           End If
373           If Demand5 > 0 Then
374               TotalTemp = TotalTemp + Round(((total5 / Demand5) * 100), 0)
375           End If
              ''Report.Text21.SetText Round(TotalTemp / 5, 0) & " %"
              'Report.Text21.SetText Round((((IIf(Demand1 = 0, 0, Round(((total1 / Demand1) * 100), 0))) + (IIf(Demand2 = 0, 0, Round(((total2 / Demand2) * 100), 0))) + (IIf(Demand3 = 0, 0, Round(((total3 / Demand3) * 100), 0))) + (IIf(Demand4 = 0, 0, Round(((total4 / Demand4) * 100), 0))) + (IIf(Demand5 = 0, 0, Round(((total5 / Demand5) * 100), 0)))) / 5), 0) & " %"
              
376           TotalTemp = 0
377           If total1 > 0 Then
378               TotalTemp = TotalTemp + Round(((NonNCM1 / total1) * 100), 0)
379           End If
380           If total2 > 0 Then
381               TotalTemp = TotalTemp + Round(((NonNCM2 / total2) * 100), 0)
382           End If
383           If total3 > 0 Then
384               TotalTemp = TotalTemp + Round(((NonNCM3 / total3) * 100), 0)
385           End If
386           If total4 > 0 Then
387               TotalTemp = TotalTemp + Round(((NonNCM4 / total4) * 100), 0)
388           End If
389           If total5 > 0 Then
390               TotalTemp = TotalTemp + Round(((NonNCM5 / total5) * 100), 0)
391           End If
392           Report.Text23.SetText Round(TotalTemp / 5, 0) & " %"
              
              'Report.Text23.SetText Round((((IIf(total1 = 0, 0, Round(((NonNCM1 / total1) * 100), 0))) + (IIf(total2 = 0, 0, Round(((NonNCM2 / total2) * 100), 0))) + (IIf(total3 = 0, 0, Round(((NonNCM3 / total3) * 100), 0))) + (IIf(total4 = 0, 0, Round(((NonNCM4 / total4) * 100), 0))) + (IIf(total5 = 0, 0, Round(((NonNCM5 / total5) * 100), 0)))) / 5), 0) & " %"
393           ProgressBar1.value = 7
394           DoEvents
395           sSQL = "select COUNT(noinv) total "
396           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
397           sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
398           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
399           sSQL = sSQL & " AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0 AND INVDATE between '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'"
              
400           total1 = 0
401           Set rsTemp = objO.FetchData(sSQL)
402           If Not rsTemp.EOF Then
403               rsTemp.MoveFirst
404               total1 = CDbl(rsTemp!total)
405           End If
406           Report.Text43.SetText total1
              
407           ProgressBar1.value = 8
408           DoEvents
              
409           If Log.RPLHistori = 0 Then
410               sSQL = "select case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
411               sSQL = sSQL & " from t_invoice A WITH (NOLOCK)"
412               sSQL = sSQL & " left join SALESUNIT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NORANGKA"
413               sSQL = sSQL & " where invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
414               sSQL = sSQL & " and a.fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
415               sSQL = sSQL & " group by case when b.norangka is null then 'Non NCM' else 'NCM' end"
416           Else
417               sSQL = "select case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
418               sSQL = sSQL & " from t_invoice A WITH (NOLOCK)"
419               sSQL = sSQL & " left join [" & dbReplikasiHO & "]..SUMMARY B WITH (NOLOCK) ON A.NO_CHASSIS = B.NORANGKA"
420               sSQL = sSQL & " where invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
421               sSQL = sSQL & " and a.fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
422               sSQL = sSQL & " group by case when b.norangka is null then 'Non NCM' else 'NCM' end"
              
423           End If
424           Set rsTemp = objO.FetchData(sSQL)
425           NCM1 = 0
426           NonNCM1 = 0
427           If Not rsTemp.EOF Then
428               rsTemp.MoveFirst
429               rsTemp.filter = "STATUS = 'NCM'"
430               If Not rsTemp.EOF Then
431                   NCM1 = CDbl(rsTemp!unit)
432               End If
433               rsTemp.filter = ""
434               rsTemp.filter = "STATUS = 'Non NCM'"
435               If Not rsTemp.EOF Then
436                   NonNCM1 = CDbl(rsTemp!unit)
437               End If
438           End If
              
439           total2 = 0
440           total2 = (total1 + NCM1 + NonNCM1)
              
441           Report.Text44.SetText NCM1
442           Report.Text45.SetText NonNCM1
443           Report.Text46.SetText total2 '''(total1 + NCM1 + NonNCM1)
              
444           If total2 = 0 Then
445               Report.Text50.SetText "0 %"
446           Else
447               Report.Text50.SetText Round((NonNCM1 / total2) * 100, 0) & " %"
448           End If
              ''demand service share : rata2 per bulan penjualan 8 thn kebelakang * 2.2
              
449           ProgressBar1.value = 9
450           DoEvents
              
451           Demand1 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(YEAR,-8,'" & Format(Tgl1, "MM/DD/YYYY") & "') and '" & Format(Tgl2, "MM/DD/YYYY") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
452           Demand1 = Round(Demand1 * 2.2, 0)
              
453           Report.Text47.SetText Demand1
              
454           If Demand1 > 0 Then
455               Report.Text48.SetText Round((total1 / Demand1) * 100, 0) & " %"
456               Report.Text49.SetText Round((total2 / Demand1) * 100, 0) & " %"
457           Else
458               Report.Text48.SetText "0 %"
459               Report.Text49.SetText "0 %"
460           End If
              
461           ProgressBar1.value = 10
462           DoEvents
463           ProgressBar1.Visible = False
464           cmdCetak.Enabled = True
              
              ''
465           Report.Text11.SetText "Laporan Retention Rate " & gPerush
466           Report.Text9.SetText GetNamaBulan(Month(Tgl1), 1) & " " & year(Tgl1)
467           Report.Text27.SetText "Laporan Service Share " & gPerush
468           Report.Text26.SetText GetNamaBulan(Month(Tgl1), 1) & " " & year(Tgl1)
              
469       With CRViewer1
470           .ReportSource = Report
471           .ViewReport
472           .Zoom 100
473           .Visible = True
474           .Left = 300
475           .Height = Me.Height - 400
476           .Width = Me.Width - .Left
477       End With
         
478               cmdTutup.Visible = True
479               cmdTutup.Left = CRViewer1.Left + 20
480               cmdTutup.Top = CRViewer1.Top + 400
                  
          'ProgressBar2.Visible = False
          'Frame2.Visible = False
481       Screen.MousePointer = vbDefault
482       Call DisableExport
483       Call setLokasiTombol
484      Exit Sub

cmdCetak_Click_Error:
485       Screen.MousePointer = vbDefault
486       cmdCetak.Enabled = True
487       ProgressBar1.Visible = True
488       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCetak_Click of Form frmRepoSBRate")
End Sub
 

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
End Sub

Private Sub Form_Activate()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height + 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > (asxPowerBanner1.Width + 100) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Load()
    Tgl1 = isnow
    Tgl2 = isnow
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
End Sub


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 345
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdPrintOut.Top = CRViewer1.Top + 345
    cmdPrintOut.Left = cboPrinter.Left + cboPrinter.Width + 30
End Sub

Private Sub DisableExport()
    Dim objPrinter As Printer
    Set objPrinter = Printer
    If InStr(UCase(objPrinter.DeviceName), "PDF") <> 0 Then
        cboPrinter = ""
    Else
        cboPrinter = objPrinter.DeviceName
    End If
    cboPrinter.Clear
    
    For Each pr In Printers
        If InStr(UCase(pr.DeviceName), "PDF") = 0 Then cboPrinter.AddItem pr.DeviceName
    Next pr
    If Log.PdfOnly = 0 Then
        CRViewer1.EnableExportButton = True
        CRViewer1.EnablePrintButton = True
        cmdExport.Visible = False
        cboPrinter.Visible = False
        cmdPrintOut.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdPrintOut_Click()
    For Each pr In Printers
        If UCase(pr.DeviceName) = UCase(cboPrinter) Then
            GoTo Cetak
        End If
    Next pr
    MsgBox "Printer tidak ditemukan Pilih Printer dahulu", vbInformation, "Proses Cetak Gagal"
    cboPrinter.SetFocus
    Exit Sub
    
Cetak:
    Dim Psize As Integer
    Dim POr As Integer
    
    Psize = Report.PaperSize
    POr = Report.PaperOrientation
    
    Set Printer = Printers(cboPrinter.ListIndex)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    Report.PaperOrientation = POr
    Report.PaperSize = Psize
    Report.PrintOut
    
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Salah
    Dim NamaFile As String
    NamaFile = ""
    
    Dim f As New FormExplorer
    f.Show (vbModal)
    Dim MyPath As String
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    
    If Log.PdfOnly = 1 Then
         
        NamaFile = MyPath & f.FileName
        
        If UCase(Right(Trim(f.FileName), 5)) = ".XLSX" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 5)
        If UCase(Right(Trim(f.FileName), 4)) = ".XLS" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 4)
         
        
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = NamaFile & ".Doc"
        Report.ExportOptions.RTFExportAllPages = True
        Report.ExportOptions.FormatType = crEFTExactRichText
        Report.Export False
        MsgBox NamaFile & ".Doc" & "  ... Data sudah di Export !", vbInformation
        Exit Sub
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


