VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmRepoSupportTLHO 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form Support TL"
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
         Format          =   64225281
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
         Format          =   64225281
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
      FormatString    =   "   Laporan Support TL"
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
Attribute VB_Name = "frmRepoSupportTLHO"
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
          
2         If KodeDealer = "" Then
3             MsgBox "Silahkan pilih dealer terlebih dahulu.", vbInformation
4             Exit Sub
5         End If
          Dim Txt As String '
          
6         Screen.MousePointer = vbHourglass
7         cmdCetak.Enabled = False
8     Set Report = New RptRepoSupportTLHO
                Dim rsRepo As New ADODB.Recordset
9               Txt = "select '' as kategoryProblem , '' as kategoryKomponen,"
10              Txt = Txt & " n.keterangan as problem, j.FMDL , j.NO_POLISI,"
11              Txt = Txt & " case j.FN_RETURN when 1 then 'RTJ' else 'Tidak' end as RTJ,"
12              Txt = Txt & " j.DATE, '' as tglPBT,'' as kirimPBT,''as tglKirimPBT,'' as tglResponPBT,"
13              Txt = Txt & " j.DATE as tglSelesaiPerbaikan,'' as tglPesanPart,'' as tglPartTersedia,"
14              Txt = Txt & " n.infoTL as PenyebabMasalah,n.statusCase as PerbaikanDilakukan,"
15             Txt = Txt & " 'FIR' as fir,'0' as LTRSA from NeedSupportTL n with (nolock)"
16             Txt = Txt & " left join JOBDTA j with (nolock) on j.NO_WO = n.no_wo"
17             Txt = Txt & " where n.kodeDealer = '" & Trim(KodeDealer) & "' and n.tglEntry between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'"
                
18             Set rsRepo = objO.FetchDataHO(Txt)
19             Report.database.SetDataSource rsRepo
20             Report.txtPerusahaan.SetText namaDealer
21             Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
                 
22        cmdCetak.Enabled = True
23        With CRViewer1
24            .ReportSource = Report
25            .ViewReport
26            .Zoom 100
27            .Visible = True
28            .Left = 300
29            .Height = Me.Height - 400
30            .Width = Me.Width - .Left
31        End With
         
32        cmdTutup.Visible = True
33        cmdTutup.Left = CRViewer1.Left + 20
34        cmdTutup.Top = CRViewer1.Top + 400
                  
          'ProgressBar2.Visible = False
          'Frame2.Visible = False
35        Screen.MousePointer = vbDefault
          
36       Exit Sub

cmdCetak_Click_Error:
37        Screen.MousePointer = vbDefault
38        cmdCetak.Enabled = True
39        ProgressBar1.Visible = True
40        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCetak_Click of Form frmRepoSupportTLHO")
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

