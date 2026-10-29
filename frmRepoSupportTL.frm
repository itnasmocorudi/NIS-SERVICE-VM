VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmRepoSupportTL 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Monitoring Support TL"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   13875
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
      TabIndex        =   6
      Top             =   4200
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
      Height          =   270
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   360
      TabIndex        =   7
      Top             =   1920
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
      EnableSearchControl=   0   'False
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   6975
      Begin TDBDate6Ctl.TDBDate Tgl1 
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   165
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoSupportTL.frx":0000
         Caption         =   "frmRepoSupportTL.frx":012C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoSupportTL.frx":0198
         Keys            =   "frmRepoSupportTL.frx":01B6
         Spin            =   "frmRepoSupportTL.frx":0214
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
      Begin TDBDate6Ctl.TDBDate Tgl2 
         Height          =   330
         Left            =   3480
         TabIndex        =   2
         Top             =   172
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoSupportTL.frx":023C
         Caption         =   "frmRepoSupportTL.frx":0368
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoSupportTL.frx":03D4
         Keys            =   "frmRepoSupportTL.frx":03F2
         Spin            =   "frmRepoSupportTL.frx":0450
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
         TabIndex        =   4
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal PBT Internal"
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1500
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmRepoSupportTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmRepoSupportTL
' DateTime  : 15/05/2015 09:55
' Author    : RUDI KRISTANTO
' Purpose   :
'---------------------------------------------------------------------------------------


Private Sub hapusForm()
    Tgl1 = isnow
    Tgl2 = isnow
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

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = 400
    CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    CRViewer1.Width = ScaleWidth - 40
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
    'asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call hapusForm
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean

    boleh = True
    If Nul(Trim(Tgl2)) = "" Then
        boleh = False
        MsgBox "Tanggal Akhir tidak boleh Kosong", vbInformation + vbOKOnly
    End If
    
    If boleh Then
        Call Cetak
    End If
End Sub

Private Sub Cetak()
      Dim rsRepo As ADODB.Recordset
      Dim txt As String
1        On Error GoTo Cetak_Error

2     Screen.MousePointer = vbHourglass
          
3         Set Report = New RptRepoSupportTL
          
4         txt = "select '' as kategoryProblem , '' as kategoryKomponen,"
5         txt = txt & " n.keterangan as problem, j.FMDL , j.NO_POLISI,"
6         txt = txt & " case j.FN_RETURN when 1 then 'RTJ' else 'Tidak' end as RTJ,"
7         txt = txt & " j.DATE, '' as tglPBT,'' as kirimPBT,''as tglKirimPBT,'' as tglResponPBT,"
8         txt = txt & " j.DATE as tglSelesaiPerbaikan,'' as tglPesanPart,'' as tglPartTersedia,"
9         txt = txt & " n.infoTL as PenyebabMasalah,n.statusCase as PerbaikanDilakukan,"
10        txt = txt & " 'FIR' as fir,'0' as LTRSA from NeedSupportTL n with (nolock)"
11        txt = txt & " left join JOBDTA j with (nolock) on j.NO_WO = n.no_wo"
12        txt = txt & " where n.tglEntry between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'"
          
13        Set rsRepo = objO.FetchData(txt)
          
14        Report.Database.SetDataSource rsRepo
15        Report.txtPerusahaan.SetText gPerush
16        Report.txtAlamat.SetText gAlamat
17        Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
          
18        With CRViewer1
19            .ReportSource = Report
20            .ViewReport
21            .Zoom 100
22            .Visible = True
23        End With
24        cmdTutup.Visible = True
25        cmdSettingPrinter.Visible = True
26        Screen.MousePointer = vbDefault

27       On Error GoTo 0
28       Exit Sub

Cetak_Error:

29        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Cetak of Form frmRepoSupportTL")
End Sub

