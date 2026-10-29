VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPiutangAR 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Piutang"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   12300
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   555
      Visible         =   0   'False
      Width           =   1050
   End
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
      Height          =   300
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   555
      Visible         =   0   'False
      Width           =   1395
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   7200
      TabIndex        =   2
      Top             =   120
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
   Begin VB.CommandButton cmdCetak 
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3660
      Width           =   1635
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3660
      Width           =   1755
   End
   Begin TDBDate6Ctl.TDBDate DtAwal 
      Height          =   315
      Left            =   2070
      TabIndex        =   5
      Top             =   360
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Calendar        =   "frmPiutangAR.frx":0000
      Caption         =   "frmPiutangAR.frx":012C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPiutangAR.frx":0198
      Keys            =   "frmPiutangAR.frx":01B6
      Spin            =   "frmPiutangAR.frx":0214
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
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
   Begin TDBDate6Ctl.TDBDate DtAkhir 
      Height          =   315
      Left            =   3990
      TabIndex        =   6
      Top             =   360
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Calendar        =   "frmPiutangAR.frx":023C
      Caption         =   "frmPiutangAR.frx":0368
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPiutangAR.frx":03D4
      Keys            =   "frmPiutangAR.frx":03F2
      Spin            =   "frmPiutangAR.frx":0450
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2730
      TabIndex        =   7
      Top             =   735
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
      Index           =   0
      Left            =   2070
      TabIndex        =   8
      Top             =   735
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
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
      Index           =   1
      Left            =   2880
      TabIndex        =   9
      Top             =   735
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   2
      Left            =   2070
      TabIndex        =   10
      Top             =   1140
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   3
      Left            =   3990
      TabIndex        =   11
      Top             =   1140
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   4
      Left            =   2055
      TabIndex        =   17
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   5
      Left            =   3975
      TabIndex        =   18
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   270
      Left            =   3600
      TabIndex        =   22
      Top             =   1935
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
      Left            =   2040
      TabIndex        =   23
      Top             =   1935
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   24
      Top             =   2280
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.CheckBox chDisplayNol 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tampilkan hanya yang belum Lunas"
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   2760
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.CheckBox chDisplayOv 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tampilkan hanya yang Overdue"
      Height          =   375
      Left            =   2040
      TabIndex        =   26
      Top             =   3120
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Pelanggan"
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
      Left            =   720
      TabIndex        =   25
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
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
      Left            =   3615
      TabIndex        =   20
      Top             =   1605
      Width           =   225
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Faktur Service"
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
      Left            =   615
      TabIndex        =   19
      Top             =   1605
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   1350
      TabIndex        =   16
      Top             =   360
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
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
      Left            =   3630
      TabIndex        =   15
      Top             =   420
      Width           =   225
   End
   Begin VB.Label Label2 
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
      Left            =   1710
      TabIndex        =   14
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Faktur Part"
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
      Left            =   855
      TabIndex        =   13
      Top             =   1185
      Width           =   1050
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
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
      Left            =   3630
      TabIndex        =   12
      Top             =   1185
      Width           =   225
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8295
      Left            =   0
      Picture         =   "frmPiutangAR.frx":0478
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmPiutangAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim objBilling As New BClsBilling
Dim Report As Object
Dim bolehCetak As Boolean

Private Sub cmdCetak_Click()
    Set Report = New rptLaporanPiutang
    Dim rsRepo As New ADODB.Recordset
    Dim vNamaTabel As String
    vNamaTabel = "PIUTANGAR"
            
    '' TANPA TEMP TABEL - LANGSUNG T_INVOICE
    sSQL = "SELECT INVDATE, NO_POLISI, NO_WO, '" & gKodeFaktur & "'+NOINV AS NOINV, NOFAKTUR, ISNULL((CASE WHEN NOFAKTUR = '' THEN 0 ELSE (SELECT isnull(SUM(ISNULL(JUMLAHTRN,0)),0) FROM [" & NisDb & "]..ARTRN WITH(NOLOCK) WHERE NOBUKTI = NOFAKTUR AND NOFAKTUR = A.NOFAKTUR GROUP BY NOFAKTUR) END),0) AS PART, " & _
           "ISNULL((SELECT ISNULL(SUM(ISNULL(JUMLAHTRN,0)),0) FROM [" & NisDb & "]..ARTRN WITH(NOLOCK) WHERE NOBUKTI = NOFAKTUR AND NOFAKTUR = '" & gKodeFaktur & "'+A.NOINV GROUP BY NOFAKTUR),0) AS JASA, " & _
           "ISNULL((CASE WHEN NOFAKTUR = '' THEN 0 ELSE (SELECT isnull(SUM(ISNULL(JUMLAHTRN,0)),0) * -1 FROM [" & NisDb & "]..ARTRN WITH(NOLOCK) WHERE NOBUKTI <> NOFAKTUR AND NOFAKTUR = A.NOFAKTUR GROUP BY NOFAKTUR) END),0) AS ARBYRPART, " & _
           "ISNULL((SELECT ISNULL(SUM(ISNULL(JUMLAHTRN,0)),0) * -1 FROM [" & NisDb & "]..ARTRN WITH(NOLOCK) WHERE NOBUKTI <> NOFAKTUR AND NOFAKTUR = '" & gKodeFaktur & "'+A.NOINV GROUP BY NOFAKTUR),0) AS ARBYRJASA, " & _
           "FRONT, FS_KD_PLG, CUSTOMER, PAID, NO_CHASSIS, " & _
           " (SELECT TOP 1 TGLJTHTMP FROM [" & NisDb & "]..ARTRN WITH(NOLOCK) WHERE NOBUKTI = NOFAKTUR AND NOFAKTUR = '" & gKodeFaktur & "'+A.NOINV) AS OVERDUE " & _
           "FROM " & gDb & "..[T_INVOICE] A WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND INVDATE BETWEEN '" & Format(DtAwal, "MM/DD/YYYY") & "' AND '" & Format(DtAkhir, "MM/DD/YYYY") & "' " & _
           "AND FRONT LIKE '%" & Trim(text(0)) & "%' AND NOFAKTUR BETWEEN '" & Trim(text(2)) & "' AND '" & Trim(text(3)) & "' AND '" & gKodeFaktur & "'+A.NOINV BETWEEN '" & Trim(text(4)) & "' AND '" & Trim(text(5)) & "' "

    If Trim(text(6)) <> "" Then
        sSQL = sSQL & " and FS_KD_PLG = '" & Trim(text(6)) & "' "
    End If
    
    If chDisplayOv = Checked Then
        sSQL = sSQL & " AND (SELECT TOP 1 TGLJTHTMP FROM [" & NisDb & "]..ARTRN WITH(NOLOCK) WHERE NOBUKTI = NOFAKTUR AND NOFAKTUR = '" & gKodeFaktur & "'+A.NOINV oRDER By TGLJTHTMP) < '" & Format(isnow, "MM/DD/YYYY") & "' "
    End If
    
    If chDisplayNol = Checked Then
        sSQL = sSQL & " AND ((ISNULL((SELECT ISNULL(SUM(ISNULL(JUMLAHTRN,0)),0) FROM [" & NisDb & "]..ARTRN WITH(NOLOCK) WHERE NOBUKTI = NOFAKTUR AND NOFAKTUR = '" & gKodeFaktur & "'+A.NOINV GROUP BY NOFAKTUR),0) - ISNULL((SELECT ISNULL(SUM(ISNULL(JUMLAHTRN,0)),0) * -1 FROM [" & NisDb & "]..ARTRN WHERE NOBUKTI <> NOFAKTUR AND NOFAKTUR = '" & gKodeFaktur & "'+A.NOINV GROUP BY NOFAKTUR),0)) <> 0 " & _
                      " OR (ISNULL((CASE WHEN NOFAKTUR = '' THEN 0 ELSE (SELECT isnull(SUM(ISNULL(JUMLAHTRN,0)),0) FROM [" & NisDb & "]..ARTRN WHERE NOBUKTI = NOFAKTUR AND NOFAKTUR = A.NOFAKTUR GROUP BY NOFAKTUR) END),0) - ISNULL((CASE WHEN NOFAKTUR = '' THEN 0 ELSE (SELECT isnull(SUM(ISNULL(JUMLAHTRN,0)),0) * -1 FROM [" & NisDb & "]..ARTRN WHERE NOBUKTI <> NOFAKTUR AND NOFAKTUR = A.NOFAKTUR GROUP BY NOFAKTUR) END),0)) <> 0) "
    End If
    
    sSQL = sSQL & " ORDER BY A.NOINV"

    Set rsRepo = objO.FetchData(sSQL)
    If Not rsRepo.EOF Then
        Report.database.SetDataSource rsRepo
        Report.txtTanggal.SetText "Tanggal : " & Format(DtAwal, "DD - MMMM - YYYY") & " sampai dengan " & Format(DtAkhir, "DD - MMMM - YYYY")
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 100
            
            .Top = 0
            .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
            .Width = Me.Width - 600
            .Height = asxPowerBanner1.Height - 100
            
            cmdTutup.Top = .Top + 430
            cmdTutup.Left = .Left
            cmdTutup.Visible = True
            
            cmdSettingPrinter.Top = cmdTutup.Top
            cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
            cmdSettingPrinter.Visible = True
        
            .Visible = True
        End With
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Belum ada data !!"
    End If
End Sub


Private Sub BuatTabelRepoAR(vNamaTabel As String)
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    
    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
         & "    [SA][char](5) NOT NULL DEFAULT '', " _
         & "    [NOINV][varchar](15) NOT NULL DEFAULT '', " _
         & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
         & "    [NO_POLISI][varchar](12) NOT NULL DEFAULT '', " _
         & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
         & "    [JOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [JRTE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [RTJ][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [THS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [GRP][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TWC][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [LUB][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [PDI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [WAS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [BOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [MATERIAL][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [PART][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [OLI][decimal](18,2) NOT NULL DEFAULT 0, "
         
    sSQL = sSQL + "[SUBLET][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SUBTOT][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [METERAI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TOTAL][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TIPEINV][varchar](20) NOT NULL DEFAULT '', " _
         & "    [CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
         & "    [LABOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [NOFAKTUR][varchar](30) NOT NULL DEFAULT ''," _
         & "    [HPP][decimal](18,2) NOT NULL DEFAULT 0," _
         & "    [INVDATE][datetime] NOT NULL DEFAULT '01/01/3000', " _
         & "    [KODEPLG][varchar](50) NOT NULL DEFAULT '', " _
         & "    [PAID][varchar](5) NOT NULL DEFAULT '') ON [PRIMARY]"
    Call objO.UpdateData(sSQL)
End Sub


Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset
    createTemp = False
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelRepoAR(vNamaTabel)
        createTemp = True
    End If
End Function

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

Private Function kodePelanggan(rs As ADODB.Recordset) As String
    If Not rs.EOF Then
        kodePelanggan = Trim(rs("KDPLG"))
    End If
End Function

Private Function cekDuo(hrg As Double, nINV As String, Optional stat As Integer)
    If Right(Trim(nINV), 2) = "-B" Then
        cekDuo = 0
        If stat = 1 Then
            cekDuo = CDbl(hrg)
        End If
    Else
        cekDuo = CDbl(hrg)
        If stat = 2 Then
            cekDuo = 0
        End If
    End If
End Function

Private Sub IsiRepo(vNamaTabel As String)
Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
Dim hlp5 As String
Dim hlp6 As String
Dim hlp7 As String
Dim rsHlp As ADODB.Recordset
Dim rsAdd As ADODB.Recordset
Dim nNilai As Double
Dim nDuoInv As Boolean
Dim subletOnly As Boolean
Dim noInv As String
Dim nmCust As String
Dim nStatus As String
    
    If Log.RPLHistori = 0 Then
        Set rsTemp = objO.FetchData("SELECT CASE DELIV WHEN 'BTL' THEN CASE WHEN FD_RETUR BETWEEN '" & Format(Tgl1, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(Tgl2, "MM/DD/YYYY 23:59:59") & "' THEN 0 ELSE 1 END ELSE 1 END AS HIT_UNIT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                             & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                             & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP, PAID, FD_RETUR " _
                             & "FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C WITH(NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                             & "WHERE DELIV <> 'BTL' AND INVDATE BETWEEN '" & Format(DtAwal, "MM/DD/YYYY") & "' AND '" & Format(DtAkhir, "MM/DD/YYYY") & "' " & Ket & " ORDER BY NOINV")
    Else
        Set rsTemp = objO.FetchData("SELECT CASE DELIV WHEN 'BTL' THEN CASE WHEN FD_RETUR BETWEEN '" & Format(Tgl1, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(Tgl2, "MM/DD/YYYY 23:59:59") & "' THEN 0 ELSE 1 END ELSE 1 END AS HIT_UNIT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                             & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                             & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP, PAID, FD_RETUR " _
                             & "FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM C WITH(NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                             & "WHERE DELIV <> 'BTL' AND INVDATE BETWEEN '" & Format(DtAwal, "MM/DD/YYYY") & "' AND '" & Format(DtAkhir, "MM/DD/YYYY") & "' " & Ket & " ORDER BY NOINV")
    End If
                             '--:: 23 Agustus 2010 ; Menampilkan nilai retur ; DELIV <> 'BTL' AND ::--
    If Not rsTemp.EOF Then
        Call BuatTabelRepoAR(vNamaTabel)
        jmlUnit = 0
        Do While Not rsTemp.EOF
            jmlUnit = jmlUnit + rsTemp("HIT_UNIT") 'IIf(rsTemp("DELIV") <> "BTL" Or (rsTemp("DELIV") = "BTL" And (Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") >= Format(Tgl1, "MM/DD/YYYY") And Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") <= Format(Tgl2, "MM/DD/YYYY"))), 1, 0)
            subletOnly = True
            For i = 1 To 8
                hlp = "REP" & i
                If rsTemp(hlp) <> "" Then
                    subletOnly = False
                    hlp2 = "KLPKERJA" & i
                    hlp3 = "RTE" & i    '"ACTRATE" & i
                    hlp4 = "MEKANIK" & i
                    hlp5 = "OFF" & i
                    hlp6 = "DISC" & i
                    hlp7 = "FREEC" & i
                    If createTemp(vNamaTabel) Then
                        nDuoInv = False
                        If rsTemp("NDUOINV") = 1 Then
                            nDuoInv = True
                        End If
                        '#########################################################'
                        '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                        '#########################################################'
                        nStatus = rsTemp("STATUS")
                        nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100))
                        noInv = gKodeFaktur & Trim(rsTemp("NOINV"))
                        nmCust = Trim(rsTemp("CUSTOMER"))
                        nKodePlg = Trim(rsTemp("FS_KD_PLG"))
                        If nStatus = "AFFNONPPN" Then
                            nNilai = nNilai * (100 / 110)
                        End If
                        
                        Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
                        If Not rsHlp.EOF Then
                            Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                         & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "BOR = '" & setKoma(cekDuo((CDbl(rsTemp("BODY"))), noInv)) & "', MATERIAL = '" & setKoma(cekDuo((CDbl(rsTemp("MTRL"))), noInv)) & "', PART = '" & cekDuo(setKoma(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noInv) & "', " _
                                         & "OLI = '" & cekDuo(setKoma(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noInv) & "', SUBLET = '" & cekDuo(setKoma(CDbl(rsTemp("SBLT"))), noInv) & "', SUBTOT = '" & cekDuo(setKoma(CDbl(rsTemp("SUBTOT"))), noInv) & "', " _
                                         & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(Right(noInv, 2) = "-B", "", Trim(rsTemp("NOFAKTUR"))) & "', HPP = '" & cekDuo(setKoma(CDbl(rsTemp("HPP"))), noInv) & "', PAID = '" & Trim(rsTemp("PAID")) & "' " _
                                         & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
                        Else
                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                         & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "'" & IIf(rsTemp(hlp2) = "SBE" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "','" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "'" & IIf(rsTemp(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "'" & cekDuo(setKoma(CDbl(rsTemp("BODY"))), noInv) & "', '" & cekDuo(setKoma(CDbl(rsTemp("MTRL"))), noInv) & "', '" & cekDuo(setKoma(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noInv) & "', '" & cekDuo(setKoma(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noInv) & "', '" & cekDuo(setKoma(CDbl(rsTemp("SBLT"))), noInv) & "', '" & cekDuo(setKoma(CDbl(rsTemp("SUBTOT"))), noInv) & "', " _
                                         & " '" & cekDuo(setKoma(CDbl(rsTemp("MET"))), noInv, 1) & "', '" & setKoma(CDbl(rsTemp("TOTALL"))) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & setKoma(CDbl(rsTemp("LABOR"))) & "', '" & IIf(Right(noInv, 2) = "-B", "", Trim(rsTemp("NOFAKTUR"))) & "', '" & cekDuo(setKoma(CDbl(rsTemp("HPP"))), noInv) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(nKodePlg) & "', '" & Trim(rsTemp("PAID")) & "')")
                        End If
                        If rsHlp.State <> adStateClosed Then rsHlp.Close
                        
                    End If
                End If
            Next
            If subletOnly Then
                If rsTemp("NDUOINV") = 1 Then nDuoInv = True
                '#########################################################'
                '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                '#########################################################'
                noInv = gKodeFaktur & Trim(rsTemp("NOINV"))
                
                Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
                If Not rsHlp.EOF Then
                    Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                 & "JOR = '0', TIPEINV = '" & Trim(rsTemp("STATUS")) & "', CUSTOMER = '" & Trim(rsTemp("CUSTOMER")) & "', BOR = '" & setKoma(CDbl(rsTemp("BODY"))) & "', MATERIAL = '" & setKoma(CDbl(rsTemp("MTRL"))) & "', " _
                                 & "PART = '" & setKoma(CDbl(rsTemp("PART_G") + rsTemp("PART_NG"))) & "', OLI = '" & setKoma(CDbl(rsTemp("OIL_G") + rsTemp("OIL_NG"))) & "', SUBLET = '" & setKoma(CDbl(rsTemp("SBLT"))) & "', SUBTOT = '" & setKoma(CDbl(rsTemp("SUBTOT"))) & "', " _
                                 & "LABOR = '" & setKoma(CDbl(rsTemp("LABOR"))) & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "', HPP = '" & setKoma(CDbl(rsTemp("HPP"))) & "', PAID = '" & Trim(rsTemp("PAID")) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
                Else
                    Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                 & "'0','0', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                 & "'" & setKoma(CDbl(rsTemp("BODY"))) & "', '" & setKoma(CDbl(rsTemp("MTRL"))) & "', '" & setKoma(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))) & "', '" & setKoma(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))) & "', '" & setKoma(CDbl(rsTemp("SBLT"))) & "', '" & setKoma(CDbl(rsTemp("SUBTOT"))) & "', " _
                                 & " '" & setKoma(CDbl(rsTemp("MET"))) & "', '" & setKoma(CDbl(rsTemp("TOTALL"))) & "', '" & Trim(rsTemp("STATUS")) & "', '" & Trim(rsTemp("CUSTOMER")) & "','" & setKoma(CDbl(rsTemp("LABOR"))) & "', '" & Trim(rsTemp("NOFAKTUR")) & "','" & setKoma(CDbl(rsTemp("HPP"))) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(nKodePlg) & "','" & Trim(rsTemp("PAID")) & "')")
                End If
            End If
            
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
                    '#########################################################'
                    '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                    '#########################################################'
                    nStatus = rsTemp("STATUS")
                    nNilai = CDbl(rsAdd(hlp5)) * (1 - (IIf(Right(Trim(rsTemp("NOINV")), 2) = "-B", IIf(rsAdd("FREEC"), 100, 0), rsAdd("DISC")) / 100))
                    noInv = gKodeFaktur & Trim(rsTemp("NOINV"))
                    nmCust = Trim(rsTemp("CUSTOMER"))
                    nKodePlg = Trim(rsTemp("FS_KD_PLG"))
                    
                    If Trim(nStatus) = "AFFNONPPN" Then
                        nNilai = nNilai * (100 / 110)
                    End If
                    
                    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
                    If Not rsHlp.EOF Then
                        Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                     & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                     & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                     & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                     & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                     & "BOR = '" & cekDuo(setKoma(CDbl(rsTemp("BODY"))), noInv) & "', MATERIAL = '" & cekDuo(setKoma(CDbl(rsTemp("MTRL"))), noInv) & "', PART = '" & cekDuo(setKoma(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noInv) & "', " _
                                     & "OLI = '" & cekDuo(setKoma(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noInv) & "', SUBLET = '" & cekDuo(setKoma(CDbl(rsTemp("SBLT"))), noInv) & "', SUBTOT = '" & setKoma(CDbl(rsTemp("SUBTOT"))) & "', HPP = '" & cekDuo(setKoma(CDbl(rsTemp("HPP"))), noInv) & "', " _
                                     & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & setKoma(CDbl(rsTemp("LABOR"))) & "', NOFAKTUR = '" & IIf(Right(noInv, 2) = "-B", "", Trim(rsTemp("NOFAKTUR"))) & "', PAID = '" & Trim(rsTemp("PAID")) & "' " _
                                     & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
                    Else
                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                     & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "SBE" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "LUB" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & cekDuo(setKoma(CDbl(rsTemp("BODY"))), noInv) & "', '" & cekDuo(setKoma(CDbl(rsTemp("MTRL"))), noInv) & "', '" & cekDuo(setKoma(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))), noInv) & "', '" & cekDuo(setKoma(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))), noInv) & "', '" & cekDuo(setKoma(CDbl(rsTemp("SBLT"))), noInv) & "', '" & CDbl(setKoma(rsTemp("SUBTOT"))) & "', " _
                                     & "'" & cekDuo(setKoma(CDbl(rsTemp("MET"))), noInv, 1) & "', '" & setKoma(CDbl(rsTemp("TOTALL"))) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & setKoma(CDbl(rsTemp("LABOR"))) & "', '" & IIf(Right(noInv, 2) = "-B", "", Trim(rsTemp("NOFAKTUR"))) & "','" & cekDuo(setKoma(CDbl(rsTemp("HPP"))), noInv) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(nKodePlg) & "','" & Trim(rsTemp("PAID")) & "')")
                    End If
                     
                    rsAdd.MoveNext
                Loop
            End If
            If rsAdd.State <> adStateClosed Then rsAdd.Close
            rsTemp.MoveNext
        Loop
        bolehCetak = True
    Else
        bolehCetak = False
        Exit Sub
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    Call hapusForm
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub hapusForm()
    DtAwal.value = Format(isnow, "DD/MM/YYYY")
    DtAkhir.value = Format(isnow, "DD/MM/YYYY")
    For i = 0 To 5
        Select Case i
            Case Else
                text(i) = ""
        End Select
    Next
    text(2) = ""
    text(3) = "Dzzz"
    text(4) = "E"
    text(5) = "Ezzz"
    bolehCetak = False
    chDisplayNol = Checked
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
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
                gSearchOption = "SA"
                frmCari.Show (vbModal)
            Case 6
                gMainKeyCari = text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As New ADODB.Recordset

     
    Select Case Index
        Case 0
            If text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "SA"
                    frmCari.Show (vbModal)
                Else
                    text(Index + 1) = rsSearch("FS_NM_PETUGAS")
                End If
                 
            Else
                text(Index + 1) = ""
            End If
        Case 6
            If text(Index) <> "" Then
                Set rsSearch = objO.FetchDataACC("SELECT * FROM ARMST WITH(NOLOCK) WHERE KodeLgn = '" & text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "KODEPLG"
                    frmCariAcess.Show (vbModal)
                Else
                    text(Index + 1) = Left(rsSearch("NamaLgn"), 60)
                End If
            Else
                text(Index + 1) = ""
            End If
    End Select
     
End Sub
