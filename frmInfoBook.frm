VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfoBook 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Informasi Booking"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14925
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H00DCFCFB&
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
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H00DCFCFB&
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
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   10440
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
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
   Begin VB.OptionButton Option2 
      BackColor       =   &H00BFE8EA&
      Caption         =   "Tanggal Booking Order"
      Height          =   225
      Left            =   1035
      TabIndex        =   15
      Top             =   1695
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00BFE8EA&
      Caption         =   "Tanggal Booking"
      Height          =   225
      Left            =   1020
      TabIndex        =   12
      Top             =   375
      Width           =   3015
   End
   Begin VB.CommandButton TblConvert 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Simpan Text"
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
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Export Text"
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
      Left            =   4155
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Cetak &Layar"
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
      Left            =   6045
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Cetak &Printer"
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
      Left            =   7935
      Style           =   1  'Graphical
      TabIndex        =   5
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
      TabIndex        =   6
      Top             =   7065
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   2130
      TabIndex        =   0
      Top             =   675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5475
      Top             =   2715
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   600
      TabIndex        =   11
      Top             =   3270
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmInfoBook.frx":0000
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   2145
      TabIndex        =   2
      Top             =   1995
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   2
      Left            =   2130
      TabIndex        =   1
      Top             =   990
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   3
      Left            =   2145
      TabIndex        =   3
      Top             =   2310
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7815
      Left            =   0
      Picture         =   "frmInfoBook.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   255
      Left            =   1065
      TabIndex        =   19
      Top             =   2340
      Width           =   960
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3420
      TabIndex        =   18
      Top             =   2340
      Width           =   1425
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   255
      Left            =   1050
      TabIndex        =   17
      Top             =   1020
      Width           =   960
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3405
      TabIndex        =   16
      Top             =   1020
      Width           =   1425
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   1065
      TabIndex        =   14
      Top             =   2025
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3420
      TabIndex        =   13
      Top             =   2025
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   3405
      TabIndex        =   8
      Top             =   705
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   1050
      TabIndex        =   7
      Top             =   705
      Width           =   960
   End
End
Attribute VB_Name = "frmInfoBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As Object

Private Sub hapusForm()
    Text(0) = ">>->>->>>>"
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub Command1_Click()
    'Call Cetak(1)
    Call Cetak2
End Sub

Private Sub Form_Load()
    Call hapusForm
    Option1 = True
End Sub

Private Sub Cetak2()
          
          Dim rsRepo As ADODB.Recordset
          Dim nTanggal As Date
          
1         Set Report = New RptInfoBook2
2         If Option1.value = True Then
3             If Text(0) = ">>->>->>>>" Or Text(2) = ">>->>->>>>" Then
4                 MsgBox "Masukkan Tanggal Booking ! "
5                 Exit Sub
6             Else
      '            sSQL = "SELECT A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.GSM, A.CDMA, A.FRONT, A.FMDL, A.DELIV, A.TIME, ISNULL(B.DATE,'') AS TGLSVC, ISNULL(NO_WO,'') AS NO_WO, " _
      '                 & "(CASE A.DELIV WHEN 'BOK' THEN 1 ELSE 0 END) AS BOOKING, (CASE A.DELIV WHEN 'OKE' THEN 1 ELSE 0 END) AS ONORDER, (CASE A.DELIV WHEN 'BTL' THEN 1 ELSE 0 END) AS BATAL, " _
      '                 & "(CASE LEFT(A.DELIV,3) WHEN 'RES' THEN 1 ELSE 0 END) AS RECH, DATE_BOOK, TIME_D " _
      '                 & "FROM BOOK A WITH(NOLOCK) LEFT JOIN JOBDTA B WITH(NOLOCK) ON A.NOFAK = B.NOFAK WHERE A.DATE BETWEEN '" & TglYYYYMMDD(Text(0)) & "' AND '" & TglYYYYMMDD(Text(2)) & "' ORDER BY A.NO_POLISI"
                                   
                  ' ==================================================================================
                  ' 17 FEB 2017 | THANIA | PENAMBAHAN INFO ASAL BOOKING DARI BOOKING ONLINE
                  ' ==================================================================================
      '            sSQL = "SELECT A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.GSM, A.CDMA, A.FRONT, A.FMDL, A.DELIV, A.TIME, ISNULL(B.DATE,'') AS TGLSVC, ISNULL(NO_WO,'') AS NO_WO, " _
      '                 & "(CASE A.DELIV WHEN 'BOK' THEN 1 ELSE 0 END) AS BOOKING, (CASE A.DELIV WHEN 'OKE' THEN 1 ELSE 0 END) AS ONORDER, (CASE A.DELIV WHEN 'BTL' THEN 1 ELSE 0 END) AS BATAL, " _
      '                 & "(CASE LEFT(A.DELIV,3) WHEN 'RES' THEN 1 ELSE 0 END) AS RECH, DATE_BOOK, TIME_D, CONVERT(VARCHAR(12),C.TGLENTRY,101) AS TglProsesOnline " _
      '                 & "FROM BOOK A WITH(NOLOCK) LEFT JOIN JOBDTA B WITH(NOLOCK) ON A.NOFAK = B.NOFAK LEFT JOIN BOOKINGONLINE C WITH (NOLOCK) ON A.NOFAK = C.NOFAK WHERE A.DATE BETWEEN '" & TglYYYYMMDD(text(0)) & "' AND '" & TglYYYYMMDD(text(2)) & "' ORDER BY A.NO_POLISI"
                  ' ==================================================================================
                  ' ==================================================================================
                  ' 20181112 | REV | MB THANIA | PENAMBAHAN NORANGKA PADA LAPORAN
                  ' ==================================================================================

7                 sSQL = "SELECT case A.NO_CHASSIS WHEN '' then case ISNULL(B.NO_CHASSIS,'') WHEN '' THEN '' ELSE B.NO_CHASSIS END else A.NO_CHASSIS end as NO_CHASSIS ,A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.GSM, A.CDMA, A.FRONT, A.FMDL, A.DELIV, A.TIME, ISNULL(B.DATE,'') AS TGLSVC, ISNULL(NO_WO,'') AS NO_WO, " _
                       & "(CASE A.DELIV WHEN 'BOK' THEN 1 ELSE 0 END) AS BOOKING, (CASE A.DELIV WHEN 'OKE' THEN 1 ELSE 0 END) AS ONORDER, (CASE A.DELIV WHEN 'BTL' THEN 1 ELSE 0 END) AS BATAL, " _
                       & "(CASE LEFT(A.DELIV,3) WHEN 'RES' THEN 1 ELSE 0 END) AS RECH, DATE_BOOK, TIME_D, CONVERT(VARCHAR(12),C.TGLENTRY,101) AS TglProsesOnline " _
                       & "FROM BOOK A WITH(NOLOCK) LEFT JOIN JOBDTA B WITH(NOLOCK) ON A.NOFAK = B.NOFAK LEFT JOIN BOOKINGONLINE C WITH (NOLOCK) ON A.NOFAK = C.NOFAK WHERE A.DATE BETWEEN '" & TglYYYYMMDD(Text(0)) & "' AND '" & TglYYYYMMDD(Text(2)) & "' ORDER BY A.NO_POLISI"
                  ' ==================================================================================
                  
8             End If
9             nTanggal = Format(Trim(Text(0)), "DD/MM/YYYY")
10        Else
11            If Text(1) = ">>->>->>>>" Or Text(3) = ">>->>->>>>" Then
12                MsgBox "Masukkan Tanggal Booking Order ! "
13                Exit Sub
14            Else
      '            sSQL = "SELECT A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.GSM, A.CDMA, A.FRONT, A.FMDL, A.DELIV, A.TIME, ISNULL(B.DATE,'') AS TGLSVC, ISNULL(NO_WO,'') AS NO_WO, " _
      '                 & "(CASE A.DELIV WHEN 'BOK' THEN 1 ELSE 0 END) AS BOOKING, (CASE A.DELIV WHEN 'OKE' THEN 1 ELSE 0 END) AS ONORDER, (CASE A.DELIV WHEN 'BTL' THEN 1 ELSE 0 END) AS BATAL, " _
      '                 & "(CASE LEFT(A.DELIV,3) WHEN 'RES' THEN 1 ELSE 0 END) AS RECH, DATE_BOOK, TIME_D " _
      '                 & "FROM BOOK A WITH(NOLOCK) LEFT JOIN JOBDTA B WITH(NOLOCK) ON A.NOFAK = B.NOFAK WHERE DATE_BOOK BETWEEN '" & TglYYYYMMDD(Text(0)) & "' AND '" & TglYYYYMMDD(Text(2)) & "' ORDER BY A.NO_POLISI"
                  
                  ' ==================================================================================
                  ' 17 FEB 2017 | THANIA | PENAMBAHAN INFO ASAL BOOKING DARI BOOKING ONLINE
                  ' ==================================================================================
      '            sSQL = "SELECT A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.GSM, A.CDMA, A.FRONT, A.FMDL, A.DELIV, A.TIME, ISNULL(B.DATE,'') AS TGLSVC, ISNULL(NO_WO,'') AS NO_WO, " _
      '                 & "(CASE A.DELIV WHEN 'BOK' THEN 1 ELSE 0 END) AS BOOKING, (CASE A.DELIV WHEN 'OKE' THEN 1 ELSE 0 END) AS ONORDER, (CASE A.DELIV WHEN 'BTL' THEN 1 ELSE 0 END) AS BATAL, " _
      '                 & "(CASE LEFT(A.DELIV,3) WHEN 'RES' THEN 1 ELSE 0 END) AS RECH, DATE_BOOK, TIME_D, CONVERT(VARCHAR(12),C.TGLENTRY,101) AS TglProsesOnline " _
      '                 & "FROM BOOK A WITH(NOLOCK) LEFT JOIN JOBDTA B WITH(NOLOCK) ON A.NOFAK = B.NOFAK LEFT JOIN BOOKINGONLINE C WITH (NOLOCK) ON A.NOFAK = C.NOFAK WHERE DATE_BOOK BETWEEN '" & TglYYYYMMDD(text(0)) & "' AND '" & TglYYYYMMDD(text(2)) & "' ORDER BY A.NO_POLISI"
                  ' ==================================================================================
                  ' ==================================================================================
                  ' 20181112 | REV | MB THANIA | PENAMBAHAN NORANGKA PADA LAPORAN
                  ' ==================================================================================
15                sSQL = "SELECT case A.NO_CHASSIS WHEN '' then case ISNULL(B.NO_CHASSIS,'') WHEN '' THEN '' ELSE B.NO_CHASSIS END else A.NO_CHASSIS end as NO_CHASSIS ,A.NO_POLISI, A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.GSM, A.CDMA, A.FRONT, A.FMDL, A.DELIV, A.TIME, ISNULL(B.DATE,'') AS TGLSVC, ISNULL(NO_WO,'') AS NO_WO, " _
                       & "(CASE A.DELIV WHEN 'BOK' THEN 1 ELSE 0 END) AS BOOKING, (CASE A.DELIV WHEN 'OKE' THEN 1 ELSE 0 END) AS ONORDER, (CASE A.DELIV WHEN 'BTL' THEN 1 ELSE 0 END) AS BATAL, " _
                       & "(CASE LEFT(A.DELIV,3) WHEN 'RES' THEN 1 ELSE 0 END) AS RECH, DATE_BOOK, TIME_D, CONVERT(VARCHAR(12),C.TGLENTRY,101) AS TglProsesOnline " _
                       & "FROM BOOK A WITH(NOLOCK) LEFT JOIN JOBDTA B WITH(NOLOCK) ON A.NOFAK = B.NOFAK LEFT JOIN BOOKINGONLINE C WITH (NOLOCK) ON A.NOFAK = C.NOFAK WHERE DATE_BOOK BETWEEN '" & TglYYYYMMDD(Text(1)) & "' AND '" & TglYYYYMMDD(Text(3)) & "' ORDER BY A.NO_POLISI"
                  ' ==================================================================================
                  
16            End If
17            nTanggal = Format(Trim(Text(1)), "DD/MM/YYYY")
18        End If
19        Set rsRepo = objO.FetchData(sSQL)
20        Report.database.SetDataSource rsRepo
21        Report.txtTanggal.SetText TglDDMMMMYYYY(Format(nTanggal, "dd/MM/yyyy"))

22        With CRViewer1
23            .ReportSource = Report
24            .ViewReport
25            .Visible = True
              
26            .Left = asxPowerBanner1.Left + asxPowerBanner1.Width
27            .Top = 0
28            .Width = Me.Width - asxPowerBanner1.Width - 100
29            .Height = Me.Height - 500
              
30            cmdTutup.Visible = True
31            cmdTutup.Left = .Left
32            cmdTutup.Top = .Top + 520
              
33            cmdSettingPrinter.Visible = True
34            cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
35            cmdSettingPrinter.Top = .Top + 520
              
36            .Zoom 100
37        End With
38        Screen.MousePointer = vbDefault
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptInfoBook
    
    gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
    
    If Option1.value = True Then
        If Text(0) = ">>->>->>>>" Or Text(2) = ">>->>->>>>" Then
            MsgBox "Masukkan Tanggal Booking ! "
            Exit Sub
        Else
            gKetTemp = "Tgl Booking"
            gTglTemp = Text(0) & " s/d " & Text(2)
            sSQL = "select DELIV, CUSTOMER, CUST_ADRES, KOTA, TELPON, NO_POLISI, FMDL, FRONT, TIME from book WITH(NOLOCK) " _
                 & " where date >= '" & TglYYYYMMDD(Text(0)) & "'" _
                 & " and date <= '" & TglYYYYMMDD(Text(2)) & "'" _
                 & " order by no_polisi"
        End If
    Else
        If Text(1) = ">>->>->>>>" Or Text(3) = ">>->>->>>>" Then
            MsgBox "Masukkan Tanggal Booking Order ! "
            Exit Sub
        Else
            gKetTemp = "Tgl Booking Order"
            gTglTemp = Text(1) & " s/d " & Text(3)
            sSQL = "select DELIV, CUSTOMER, CUST_ADRES, KOTA, TELPON, NO_POLISI, FMDL, FRONT, TIME from book WITH(NOLOCK) " _
                 & " where date_book >= '" & TglYYYYMMDD(Text(1)) & "'" _
                 & " and date_book <= '" & TglYYYYMMDD(Text(3)) & "'" _
                 & " order by no_polisi"
        End If
    End If
    
    rpt.DataControl1.ConnectionString = gReportConn
    rpt.DataControl1.Source = sSQL
    rpt.Printer.Orientation = ddOLandscape
    rpt.Printer.PaperSize = 14
    Select Case Layar
        Case 1
            rpt.Show
        Case 2
            ARViewer1.ReportSource = rpt
        Case Else
            rpt.PrintReport False
            Unload rpt
    End Select
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
    If Me.Height > (TblConvert.Height + 1000) Then
        TblConvert.Top = Me.Height - TblConvert.Height - 1000
    End If
    If Me.Height > (TblCetak.Height + 1000) Then
        TblCetak.Top = Me.Height - TblCetak.Height - 1000
    End If
    If Me.Height > (Command2.Height + 1000) Then
        Command2.Top = Me.Height - Command2.Height - 1000
    End If
    If Me.Height > (Command1.Height + 1000) Then
        Command1.Top = Me.Height - Command1.Height - 1000
    End If
    If Me.Height > (TblKeluar.Height + 1000) Then
        TblKeluar.Top = Me.Height - TblKeluar.Height - 1000
    End If
End Sub

Private Sub TblCetak_Click()
    Call Cetak(0)
End Sub

Private Sub TblConvert_Click()
    Dim Txt As New ActiveReportsTextExport.ARExportText

    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.ShowSave
    
    Txt.FileName = CommonDialog1.FileName
    Txt.Export ARViewer1.Pages
End Sub

Private Sub Command2_Click()
    Call Cetak(2)
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
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
    Call CekTombol(KeyCode)
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
    Select Case Index
        Case 0, 1, 2, 3
            If CekTgl(Text(Index)) <> True Then
                Text(Index) = gTanggal
            End If
    End Select
End Sub

