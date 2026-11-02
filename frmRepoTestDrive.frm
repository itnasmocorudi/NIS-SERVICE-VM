VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoTestDrive 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Test Drive dan FI"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
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
      Left            =   3855
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4785
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCetak 
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
      Left            =   9300
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4785
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   6180
      TabIndex        =   17
      Top             =   4815
      Visible         =   0   'False
      Width           =   3015
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   5070
      TabIndex        =   16
      Top             =   4800
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
      MICON           =   "frmRepoTestDrive.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4785
      Width           =   1050
   End
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
      Left            =   2370
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4785
      Width           =   1440
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   600
      TabIndex        =   0
      Top             =   4440
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
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   600
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
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
      TabIndex        =   9
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
      TabIndex        =   7
      Top             =   1440
      Width           =   1890
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
      TabIndex        =   8
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   12375
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1650
         TabIndex        =   2
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   124125185
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   124125185
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1650
         TabIndex        =   10
         Top             =   600
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   2475
         TabIndex        =   11
         Top             =   600
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
         Index           =   6
         Left            =   2595
         TabIndex        =   12
         Top             =   600
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SA"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   645
         Width           =   195
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
         TabIndex        =   6
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Invoice"
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
         Left            =   210
         TabIndex        =   5
         Top             =   240
         Width           =   1350
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
         TabIndex        =   4
         Top             =   1560
         Width           =   60
      End
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
      Left            =   480
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "frmRepoTestDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmRepoTestDrive
' DateTime  : 17/04/2015 14:16
' Author    : RUDI KRISTANTO
' Purpose   :
'---------------------------------------------------------------------------------------

Dim rsReport As New ADODB.Recordset
Dim Report As Object

Private Sub clearForm()
    Tgl1 = Format(isnow, "DD/MM/YYYY")
    Tgl2 = Format(isnow, "DD/MM/YYYY")
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
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
    cmdCetak.Visible = False
End Sub

Private Sub Form_Load()
    Call clearForm
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = 300 'asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height + 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > 400 Then
        CRViewer1.Width = ScaleWidth - 400 'asxPowerBanner1.Width - 100
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
   ' asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call clearForm
End Sub

Private Sub createRS()
          
1        On Error GoTo createRS_Error
2        If rsReport.State Then rsReport.Close
3         With rsReport
4             .Fields.Append "NoChassis", adVarChar, 40
5             .Fields.Append "NoPolisi", adVarChar, 15
6             .Fields.Append "Kelas", adVarChar, 35
7             .Fields.Append "Fmdl", adVarChar, 35
8             .Fields.Append "Mdl", adVarChar, 35
9             .Fields.Append "Nama", adVarChar, 35
10            .Fields.Append "Alamat", adVarChar, 100
11            .Fields.Append "Telp", adVarChar, 25
12            .Fields.Append "TanggalMasuk", adDate
13            .Fields.Append "TanggalDeliv", adDate
14            .Fields.Append "SA", adVarChar, 10
15            .Fields.Append "Mech", adVarChar, 10
16            .Fields.Append "JobOrder", adVarChar, 200
17            .Fields.Append "Minta", adVarChar, 200
18            .Fields.Append "Keluh", adVarChar, 200
19            .Fields.Append "HasilFI", adVarChar, 200
20            .Fields.Append "HasilTD", adVarChar, 200
21            .Fields.Append "CR7", adVarChar, 100
22            .Fields.Append "Urgent", adVarChar, 30
23            .Open
24        End With
          
25       On Error GoTo 0
26       Exit Sub

createRS_Error:

27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRS of Form frmRepoTestDrive")
End Sub

Private Sub tblSimpan_Click()
              
1        On Error GoTo tblSimpan_Click_Error
2         Screen.MousePointer = vbHourglass
3         Set Report = New rptRepoTestDrive
          Dim oRs As ADODB.Recordset
          Dim rsTemp As ADODB.Recordset
          
          Dim konter As Integer
          Dim isTambah As Boolean
          
4         Call createRS
5         sSQL = "select distinct j.no_wo,j.NO_CHASSIS,j.NO_POLISI,j.MODEL,j.MDL,j.FMDL,j.CUSTOMER,j.CUST_ADRES,"
6         sSQL = sSQL & " j.TELPANCUST,j.DATE,j.INVDATE,j.FRONT,j.MECH,REPLACE(REPLACE(REPLACE(j.REP1 + '; ' + j.REP2 + '; ' + j.REP3 + '; ' + j.REP4 + '; ' + j.REP5 + '; ' + j.REP6 + '; ' + j.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder,"
7         sSQL = sSQL & " j.REQ1 + '; ' + j.REQ2 + '; ' + j.REQ3 + '; ' + j.REQ4 + '; ' + j.REQ5 as Minta,"
8         sSQL = sSQL & " REPLACE(REPLACE(REPLACE(D.CMPL1 + '; ' + D.CMPL2 + '; ' + D.CMPL3 + '; ' + d.CMPL4 + '; ' + d.CMPL5 ,'; ;',''),'; ;',''),'; ;','') as Keluh,"
9         sSQL = sSQL & " isnull(i.HasilFI,'') HasilFI,isnull(i.HasilTestDrive,'') as HasilTD,isnull(i.SNJ1,'') snj1,isnull(i.isurgent1,'') isUrgent1,"
10        sSQL = sSQL & " isnull(i.SNJ2,'') snj2,isnull(i.isurgent2,'') isUrgent2,isnull(i.SNJ3,'') snj3,isnull(i.isurgent3,'') isUrgent3,"
11        sSQL = sSQL & " isnull(i.SNJ4,'') snj4,isnull(i.isurgent4,'') isUrgent4,isnull(i.SNJ5,'') snj5,isnull(i.isurgent5,'') isUrgent5,"
12        sSQL = sSQL & " isnull(i.SNJ6,'') snj6,isnull(i.isurgent6,'') isUrgent6,isnull(i.SNJ7,'') snj7,isnull(i.isurgent7,'') isUrgent7"
13        sSQL = sSQL & " from JOBDTA j with (nolock)"
          '' 20180515REV02'INSPECTION
          ''sSQL = sSQL & " left outer join INSPECTION i with (nolock) on i.NoWo = j.NO_WO"
14        sSQL = sSQL & " left outer join INSPECTION i with (nolock) on i.NoWo = j.NO_WO and i.TglEntry = (select top 1 TglEntry from INSPECTION with (nolock) where NoWo = j.NO_WO order by TglEntry desc )"
15        sSQL = sSQL & " left outer join JOBDTL d with (nolock) on d.NO_WO = j.NO_WO"
16        sSQL = sSQL & " where j.NOINV <> '' and deliv <> 'BTL' and "
17        sSQL = sSQL & " J.INVDATE Between '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND FS_LOG_DEL = '' AND FRONT like '%" & Trim(Text(5)) & "%' ORDER BY INVDATE ASC"
          
18        lblAmbilData.Visible = True
19        Set oRs = objO.FetchData(sSQL)
20        lblAmbilData.Visible = False
21        konter = 1
22        If Not oRs.EOF Then
23            oRs.MoveFirst
24            ProgressBar1.Visible = True
25            ProgressBar1.Max = oRs.RecordCount + 1
26            ProgressBar1.Min = 0
27            Do While Not oRs.EOF
28                isTambah = False
29                If Trim(oRs("snj1")) <> "" Then
30                    With rsReport
31                        .AddNew
32                        !noChassis = Trim(oRs("no_Chassis"))
33                        !NoPolisi = Trim(oRs("no_polisi"))
34                        !Kelas = Trim(oRs("model"))
35                        !Fmdl = Trim(oRs("fmdl"))
36                        !MDL = Trim(oRs("mdl"))
37                        !Nama = Trim(oRs("customer"))
38                        !Alamat = Trim(oRs("cust_adres"))
39                        !telp = Trim(oRs("telpancust"))
40                        !TanggalMasuk = Format(Trim(oRs("date")), "DD/MM/YYYY")
41                        !TanggalDeliv = Format(Trim(oRs("invdate")), "DD/MM/YYYY")
42                        !sa = Trim(oRs("front"))
43                        !Mech = Trim(oRs("mech"))
44                        !jobOrder = Left(Trim(oRs("joborder")), 200)
45                        !Minta = Left(Trim(oRs("minta")), 200)
46                        !Keluh = Left(Trim(oRs("keluh")), 200)
47                        !HasilFI = Left(Trim(oRs("hasilfi")), 200)
48                        !HasilTD = Left(Trim(oRs("hasiltd")), 200)
49                        !CR7 = Trim(oRs("snj1"))
50                        If oRs("isurgent1") = "1" Then
51                            !urgent = "Urgent"
52                        ElseIf oRs("isurgent1") = "0" Then
53                            !urgent = "Tidak Urgent"
54                        Else
55                            !urgent = ""
56                        End If
57                        .Update
58                    End With
59                    isTambah = True
60                End If
61                If Trim(oRs("snj2")) <> "" Then
62                    With rsReport
63                        .AddNew
64                        !noChassis = Trim(oRs("no_Chassis"))
65                        !NoPolisi = Trim(oRs("no_polisi"))
66                        !Kelas = Trim(oRs("model"))
67                        !Fmdl = Trim(oRs("fmdl"))
68                        !MDL = Trim(oRs("mdl"))
69                        !Nama = Trim(oRs("customer"))
70                        !Alamat = Trim(oRs("cust_adres"))
71                        !telp = Trim(oRs("telpancust"))
72                        !TanggalMasuk = Format(Trim(oRs("date")), "DD/MM/YYYY")
73                        !TanggalDeliv = Format(Trim(oRs("invdate")), "DD/MM/YYYY")
74                        !sa = Trim(oRs("front"))
75                        !Mech = Trim(oRs("mech"))
76                        !jobOrder = Left(Trim(oRs("joborder")), 200)
77                        !Minta = Left(Trim(oRs("minta")), 200)
78                        !Keluh = Left(Trim(oRs("keluh")), 200)
79                        !HasilFI = Left(Trim(oRs("hasilfi")), 200)
80                        !HasilTD = Left(Trim(oRs("hasiltd")), 200)
81                        !CR7 = Trim(oRs("snj2"))
82                        If oRs("isurgent2") = "1" Then
83                            !urgent = "Urgent"
84                        ElseIf oRs("isurgent2") = "0" Then
85                            !urgent = "Tidak Urgent"
86                        Else
87                            !urgent = ""
88                        End If
89                        .Update
90                    End With
91                    isTambah = True
92                End If
93                If Trim(oRs("snj3")) <> "" Then
94                    With rsReport
95                        .AddNew
96                        !noChassis = Trim(oRs("no_Chassis"))
97                        !NoPolisi = Trim(oRs("no_polisi"))
98                        !Kelas = Trim(oRs("model"))
99                        !Fmdl = Trim(oRs("fmdl"))
100                       !MDL = Trim(oRs("mdl"))
101                       !Nama = Trim(oRs("customer"))
102                       !Alamat = Trim(oRs("cust_adres"))
103                       !telp = Trim(oRs("telpancust"))
104                       !TanggalMasuk = Format(Trim(oRs("date")), "DD/MM/YYYY")
105                       !TanggalDeliv = Format(Trim(oRs("invdate")), "DD/MM/YYYY")
106                       !sa = Trim(oRs("front"))
107                       !Mech = Trim(oRs("mech"))
108                       !jobOrder = Left(Trim(oRs("joborder")), 200)
109                       !Minta = Left(Trim(oRs("minta")), 200)
110                       !Keluh = Left(Trim(oRs("keluh")), 200)
111                       !HasilFI = Left(Trim(oRs("hasilfi")), 200)
112                       !HasilTD = Left(Trim(oRs("hasiltd")), 200)
113                       !CR7 = Trim(oRs("snj3"))
114                       If oRs("isurgent3") = "1" Then
115                           !urgent = "Urgent"
116                       ElseIf oRs("isurgent3") = "0" Then
117                           !urgent = "Tidak Urgent"
118                       Else
119                           !urgent = ""
120                       End If
121                       .Update
122                   End With
123                   isTambah = True
124               End If
125               If Trim(oRs("snj4")) <> "" Then
126                   With rsReport
127                       .AddNew
128                       !noChassis = Trim(oRs("no_Chassis"))
129                       !NoPolisi = Trim(oRs("no_polisi"))
130                       !Kelas = Trim(oRs("model"))
131                       !Fmdl = Trim(oRs("fmdl"))
132                       !MDL = Trim(oRs("mdl"))
133                       !Nama = Trim(oRs("customer"))
134                       !Alamat = Trim(oRs("cust_adres"))
135                       !telp = Trim(oRs("telpancust"))
136                       !TanggalMasuk = Format(Trim(oRs("date")), "DD/MM/YYYY")
137                       !TanggalDeliv = Format(Trim(oRs("invdate")), "DD/MM/YYYY")
138                       !sa = Trim(oRs("front"))
139                       !Mech = Trim(oRs("mech"))
140                       !jobOrder = Left(Trim(oRs("joborder")), 200)
141                       !Minta = Left(Trim(oRs("minta")), 200)
142                       !Keluh = Left(Trim(oRs("keluh")), 200)
143                       !HasilFI = Left(Trim(oRs("hasilfi")), 200)
144                       !HasilTD = Left(Trim(oRs("hasiltd")), 200)
145                       !CR7 = Trim(oRs("snj4"))
146                       If oRs("isurgent4") = "1" Then
147                           !urgent = "Urgent"
148                       ElseIf oRs("isurgent4") = "0" Then
149                           !urgent = "Tidak Urgent"
150                       Else
151                           !urgent = ""
152                       End If
153                       .Update
154                   End With
155                   isTambah = True
156               End If
157               If Trim(oRs("snj5")) <> "" Then
158                   With rsReport
159                       .AddNew
160                       !noChassis = Trim(oRs("no_Chassis"))
161                       !NoPolisi = Trim(oRs("no_polisi"))
162                       !Kelas = Trim(oRs("model"))
163                       !Fmdl = Trim(oRs("fmdl"))
164                       !MDL = Trim(oRs("mdl"))
165                       !Nama = Trim(oRs("customer"))
166                       !Alamat = Trim(oRs("cust_adres"))
167                       !telp = Trim(oRs("telpancust"))
168                       !TanggalMasuk = Format(Trim(oRs("date")), "DD/MM/YYYY")
169                       !TanggalDeliv = Format(Trim(oRs("invdate")), "DD/MM/YYYY")
170                       !sa = Trim(oRs("front"))
171                       !Mech = Trim(oRs("mech"))
172                       !jobOrder = Left(Trim(oRs("joborder")), 200)
173                       !Minta = Left(Trim(oRs("minta")), 200)
174                       !Keluh = Left(Trim(oRs("keluh")), 200)
175                       !HasilFI = Left(Trim(oRs("hasilfi")), 200)
176                       !HasilTD = Left(Trim(oRs("hasiltd")), 200)
177                       !CR7 = Trim(oRs("snj5"))
178                       If oRs("isurgent5") = "1" Then
179                           !urgent = "Urgent"
180                       ElseIf oRs("isurgent5") = "0" Then
181                           !urgent = "Tidak Urgent"
182                       Else
183                           !urgent = ""
184                       End If
185                       .Update
186                   End With
187                   isTambah = True
188               End If
189               If Trim(oRs("snj6")) <> "" Then
190                   With rsReport
191                       .AddNew
192                       !noChassis = Trim(oRs("no_Chassis"))
193                       !NoPolisi = Trim(oRs("no_polisi"))
194                       !Kelas = Trim(oRs("model"))
195                       !Fmdl = Trim(oRs("fmdl"))
196                       !MDL = Trim(oRs("mdl"))
197                       !Nama = Trim(oRs("customer"))
198                       !Alamat = Trim(oRs("cust_adres"))
199                       !telp = Trim(oRs("telpancust"))
200                       !TanggalMasuk = Format(Trim(oRs("date")), "DD/MM/YYYY")
201                       !TanggalDeliv = Format(Trim(oRs("invdate")), "DD/MM/YYYY")
202                       !sa = Trim(oRs("front"))
203                       !Mech = Trim(oRs("mech"))
204                       !jobOrder = Left(Trim(oRs("joborder")), 200)
205                       !Minta = Left(Trim(oRs("minta")), 200)
206                       !Keluh = Left(Trim(oRs("keluh")), 200)
207                       !HasilFI = Left(Trim(oRs("hasilfi")), 200)
208                       !HasilTD = Left(Trim(oRs("hasiltd")), 200)
209                       !CR7 = Trim(oRs("snj6"))
210                       If oRs("isurgent6") = "1" Then
211                           !urgent = "Urgent"
212                       ElseIf oRs("isurgent6") = "0" Then
213                           !urgent = "Tidak Urgent"
214                       Else
215                           !urgent = ""
216                       End If
217                       .Update
218                   End With
219                   isTambah = True
220               End If
221               If Trim(oRs("snj7")) <> "" Then
222                   With rsReport
223                       .AddNew
224                       !noChassis = Trim(oRs("no_Chassis"))
225                       !NoPolisi = Trim(oRs("no_polisi"))
226                       !Kelas = Trim(oRs("model"))
227                       !Fmdl = Trim(oRs("fmdl"))
228                       !MDL = Trim(oRs("mdl"))
229                       !Nama = Trim(oRs("customer"))
230                       !Alamat = Trim(oRs("cust_adres"))
231                       !telp = Trim(oRs("telpancust"))
232                       !TanggalMasuk = Format(Trim(oRs("date")), "DD/MM/YYYY")
233                       !TanggalDeliv = Format(Trim(oRs("invdate")), "DD/MM/YYYY")
234                       !sa = Trim(oRs("front"))
235                       !Mech = Trim(oRs("mech"))
236                       !jobOrder = Left(Trim(oRs("joborder")), 200)
237                       !Minta = Left(Trim(oRs("minta")), 200)
238                       !Keluh = Left(Trim(oRs("keluh")), 200)
239                       !HasilFI = Left(Trim(oRs("hasilfi")), 200)
240                       !HasilTD = Left(Trim(oRs("hasiltd")), 200)
241                       !CR7 = Trim(oRs("snj7"))
242                       If oRs("isurgent7") = "1" Then
243                           !urgent = "Urgent"
244                       ElseIf oRs("isurgent7") = "0" Then
245                           !urgent = "Tidak Urgent"
246                       Else
247                           !urgent = ""
248                       End If
249                       .Update
250                   End With
251                   isTambah = True
252               End If
                  
253               If isTambah = False Then
254                   With rsReport
255                       .AddNew
256                       !noChassis = Trim(oRs("no_Chassis"))
257                       !NoPolisi = Trim(oRs("no_polisi"))
258                       !Kelas = Trim(oRs("model"))
259                       !Fmdl = Trim(oRs("fmdl"))
260                       !MDL = Trim(oRs("mdl"))
261                       !Nama = Trim(oRs("customer"))
262                       !Alamat = Trim(oRs("cust_adres"))
263                       !telp = Trim(oRs("telpancust"))
264                       !TanggalMasuk = Format(Trim(oRs("date")), "DD/MM/YYYY")
265                       !TanggalDeliv = Format(Trim(oRs("invdate")), "DD/MM/YYYY")
266                       !sa = Trim(oRs("front"))
267                       !Mech = Trim(oRs("mech"))
268                       !jobOrder = Left(Trim(oRs("joborder")), 200)
269                       !Minta = Left(Trim(oRs("minta")), 200)
270                       !Keluh = Left(Trim(oRs("keluh")), 200)
271                       !HasilFI = Left(Trim(oRs("hasilfi")), 200)
272                       !HasilTD = Left(Trim(oRs("hasiltd")), 200)
273                       !CR7 = ""
274                       If oRs("isurgent7") = "1" Then
275                           !urgent = "Urgent"
276                       ElseIf oRs("isurgent7") = "0" Then
277                           !urgent = "Tidak Urgent"
278                       Else
279                           !urgent = ""
280                       End If
281                       .Update
282                   End With
283               End If
284               ProgressBar1.value = konter
285               konter = konter + 1
286               oRs.MoveNext
287           Loop
288       End If
289       Report.database.SetDataSource rsReport
290       Report.txtPerusahaan.SetText gPerush
291       Report.txtAlamat.SetText gAlamat
292       Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
          
293       With CRViewer1
294           .ReportSource = Report
295           .ViewReport
296           .Zoom 100
297           .Visible = True
298       End With
299       cmdTutup.Visible = True
300       cmdSettingPrinter.Visible = True
301       ProgressBar1.Visible = False
          
302       Screen.MousePointer = vbDefault
303       Call DisableExport
304       Call setLokasiTombol
          
305      On Error GoTo 0
306      Exit Sub

tblSimpan_Click_Error:

307       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmRepoTestDrive")
End Sub

Private Sub cetakLaporan()

1        On Error GoTo cetakLaporan_Error
          Dim oRs As ADODB.Recordset
2         Screen.MousePointer = vbHourglass
          
3         Set Report = New rptRepoPenyerahan
4         sSQL = "SELECT a.NO_WO, NO_POLISI, FS_RC, TPSVC, PDS, FRONT, MECH, RUBAH, TGLJANJI, TGLRUBAH FROM JOBSERAH a  WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO " & _
                 "where DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND FRONT like '%" & Trim(Text(5)) & "%'"
              
5         Set oRs = objO.FetchData(sSQL)
6         Report.database.SetDataSource oRs
7         Report.txtPerusahaan.SetText gPerush
8         Report.txtAlamat.SetText gAlamat
9         Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
          
10        With CRViewer1
11            .ReportSource = Report
12            .ViewReport
13            .Zoom 100
14            .Visible = True
15        End With
          
16        cmdTutup.Visible = True
17        cmdSettingPrinter.Visible = True
18        Screen.MousePointer = vbDefault
          

19       On Error GoTo 0
20       Exit Sub

cetakLaporan_Error:

21        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakLaporan of Form frmRepoTestDrive")
End Sub

Private Sub cetakPerubahan()
Dim oRs As ADODB.Recordset
Screen.MousePointer = vbHourglass
    
    Set Report = New rptRepoPenyerahan
    sSQL = "SELECT a.NO_WO, NO_POLISI, FS_RC, TPSVC, PDS, FRONT, MECH, RUBAH, TGLJANJI, TGLRUBAH FROM JOBSERAH a  WITH(NOLOCK) INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO " & _
           "where DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND FRONT like '%" & Trim(Text(5)) & "%'"
        
    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cetakOTD(namaSA As String)
Dim oRs As ADODB.Recordset
Screen.MousePointer = vbHourglass
    
    Set Report = New RptRepoOTD
    
    'sSQL = "SELECT NO_POLISI, FRONT, A.NO_WO, A.NOINV, CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH AS JANJISERAH, ISNULL((SELECT TOP 1 CONVERT( varchar(10), INVDATE,101)+' '+INVTIME FROM T_INVOICE WHERE NO_WO = A.NO_WO AND LEFT(NOINV,8) = A.NOINV),'') AS PENYERAHAN, " & _
           "CASE WHEN datediff(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), ISNULL((SELECT TOP 1 CONVERT( varchar(10), INVDATE,101)+' '+INVTIME FROM T_INVOICE WHERE NO_WO = A.NO_WO AND LEFT(NOINV,8) = A.NOINV),'')) > 0 THEN 0 ELSE 1 END AS OTD " & _
           "FROM JOBDTA A INNER JOIN JOBDTL B ON A.NO_WO = B.NO_WO WHERE NOINV <> '' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND FRONT LIKE '%" & Trim(namaSA) & "%' AND DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' " & _
           "ORDER BY A.NO_WO"
           
    sSQL = "SELECT NO_POLISI, FRONT, A.NO_WO, A.NOINV, convert(datetime,CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH) AS JANJISERAH, " & _
           "Convert(DATETIME,TGLPANGGIL) AS Penyerahan, " & _
           "CASE WHEN DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), " & _
           "ISNULL(TGLPANGGIL,'')) <= 0 AND DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), ISNULL(TGLPANGGIL,'')) >= -30 THEN 1 ELSE 0 END AS SLISIH " & _
           "FROM JOBDTA A WITH(NOLOCK) " & _
           "INNER JOIN JOBDTL B WITH(NOLOCK) ON A.NO_WO = B.NO_WO " & _
           "LEFT JOIN " & Log.LogQueue & "..FINISHSERVICE C WITH(NOLOCK) ON A.NO_POLISI = C.NOPOL AND A.INVDATE = C.TGLFINISH AND PANGGIL <> 0 " & _
           "WHERE NOINV <> '' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND FRONT LIKE '%" & namaSA & "%' AND PDS = 0 AND THS = 0 AND TPSVC <> 'W' AND " & _
           "DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' " & _
           "ORDER BY A.NO_WO"
    
    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtPeriode.SetText "PERIODE : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
    Report.txtSa.SetText "SA : " & IIf(Trim(namaSA) = "", "Semua SA", Trim(namaSA))
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Text_GotFocus(Index As Integer)
        Dim DiTab As Boolean
        DiTab = False
        
        Select Case Index
        End Select
        
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
            Sendkeys ("{TAB}")
        Else
            Sendkeys ("{HOME}+{END}")
        End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 5
            gMainKeyCari = Text(Index)
            gSearchOption = "SA"
            frmCari.Show (vbModal)
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 5
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & Text(Index) & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SA"
                    frmCari.Show (vbModal)
                Else
                    Text(Index + 1) = oRs("FS_NM_PETUGAS")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
    End Select
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
    
    cmdCetak.Top = CRViewer1.Top + 345
    cmdCetak.Left = cboPrinter.Left + cboPrinter.Width + 30
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
        cmdCetak.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdCetak_Click()
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

