VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoHarianTech 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Harian Mekanik"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin ucXPButton.XPButton XPButton1 
      Height          =   270
      Left            =   3870
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   476
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
      MICON           =   "frmRepoHarianTech.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   4920
      TabIndex        =   30
      Top             =   30
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   8040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   0
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
      Left            =   2595
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
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
      Left            =   1110
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   0
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   0
      Width           =   1050
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoHarianTech.frx":001C
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   21
      Top             =   0
      Width           =   445
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   5
      Top             =   4800
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
      EnableRefreshButton=   -1  'True
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
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   7455
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   450
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   794
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.OptionButton opCetak 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton opCetak 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rekap"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton opCetak 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   17
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox cboStall 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   97189889
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   2
         Top             =   1080
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2400
         TabIndex        =   7
         Top             =   1080
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
         Index           =   2
         Left            =   1530
         TabIndex        =   3
         Top             =   1440
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   2400
         TabIndex        =   8
         Top             =   1440
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
         Left            =   2520
         TabIndex        =   9
         Top             =   1080
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   10
         Top             =   1440
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
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   97255425
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   1530
         TabIndex        =   4
         Top             =   1800
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   2655
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stall"
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
         Left            =   1080
         TabIndex        =   15
         Top             =   1845
         Width           =   300
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Foreman"
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
         Left            =   810
         TabIndex        =   14
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mechanic"
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
         Left            =   780
         TabIndex        =   13
         Top             =   1485
         Width           =   660
      End
      Begin VB.Label Label10 
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
         Left            =   870
         TabIndex        =   12
         Top             =   240
         Width           =   570
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
         TabIndex        =   11
         Top             =   240
         Width           =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         X1              =   0
         X2              =   7440
         Y1              =   600
         Y2              =   600
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
   End
   Begin ucXPButton.XPButton tblSimpan 
      Height          =   495
      Left            =   5400
      TabIndex        =   23
      Top             =   3240
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&p"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepoHarianTech.frx":240E
      PICN            =   "frmRepoHarianTech.frx":242A
      PICH            =   "frmRepoHarianTech.frx":2881
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton tblKeluar 
      Height          =   495
      Left            =   6720
      TabIndex        =   24
      Top             =   3240
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&x"
      ENAB            =   -1  'True
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepoHarianTech.frx":2D6B
      PICN            =   "frmRepoHarianTech.frx":2D87
      PICH            =   "frmRepoHarianTech.frx":325B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   480
      X2              =   7920
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label LblKnow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE WAIT  (^_^) ..."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   600
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label prosen 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "     %"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   3600
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8295
      Left            =   0
      Picture         =   "frmRepoHarianTech.frx":37A6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoHarianTech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 03 November 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsTemp As ADODB.Recordset
Dim Report As Object

Dim doubleJob As Boolean
Dim resetJob As Boolean
Dim resetRate As Boolean
Dim vNamaTabel As String
Dim objRepo As New clsRepo
Dim rsGroup    As New ADODB.Recordset

Private Sub hapusForm()
    For i = 0 To 3
        Text(i) = ""
    Next
    Tgl1 = isnow
    Tgl2 = isnow
    LblKnow.Visible = False
    ProgressBar1.Visible = False
    prosen.Visible = False
End Sub

Private Sub isiCombo()
Dim hlp As String
Dim rsHlp As ADODB.Recordset
    
    Set rsHlp = objO.FetchData("SELECT FN_JML_STALL FROM T_PROFIL WITH(NOLOCK)")
    If Not rsHlp.EOF Then
        cboStall.Clear
        For i = 1 To rsHlp("FN_JML_STALL")
            hlp = "Stall " & i
            cboStall.AddItem hlp
        Next
    End If
    rsHlp.Close
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset
    createTemp = False
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)

    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ( " _
         & "[NO][DECIMAL](19) NOT NULL DEFAULT '', " _
         & "[NOWO][VARCHAR](10) NOT NULL DEFAULT '', " _
         & "[TANGGAL][DATETIME] NOT NULL DEFAULT '01/01/3000', " _
         & "[NOSTALL][CHAR](10) NOT NULL DEFAULT '', " _
         & "[START][CHAR](5) NOT NULL DEFAULT '', " _
         & "[STATUS][CHAR](1) NOT NULL DEFAULT '', " _
         & "[SELESAI][CHAR](5) NOT NULL DEFAULT '', " _
         & "[MECH][CHAR](5) NOT NULL DEFAULT '', " _
         & "[MECHNAME][CHAR](30) NOT NULL DEFAULT '', " _
         & "[FOREMAN][CHAR](5) NOT NULL DEFAULT '', " _
         & "[MENIT][DECIMAL](18,2) NOT NULL DEFAULT 0, " _
         & "[NOPOLISI][VARCHAR](15) NOT NULL DEFAULT '', " _
         & "[NORANGKA][VARCHAR](25) NOT NULL DEFAULT '', " _
         & "[RATE1][DECIMAL](18,2) NOT NULL DEFAULT 0, " _
         & "[RATE2][DECIMAL](18,2) NOT NULL DEFAULT 0, " _
         & "[GROUP][CHAR](5) NOT NULL DEFAULT '', " _
         & "[JOB][CHAR](100) NOT NULL DEFAULT '', " _
         & "[REASON][CHAR](500) NOT NULL DEFAULT '') ON [PRIMARY]"
    Call objO.UpdateData(sSQL)
    
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    End If
End Function

Private Sub IsiTempRepo(Ket As String, vNamaTabel As String, txt As String)
      Dim rsDouble As ADODB.Recordset
      Dim btkTabel As Boolean
      Dim woLama As String
      Dim startLama As String
      Dim stallLama As String
      Dim jamKrj As Double
      Dim rate1 As Double
      Dim rate2 As Double
      Dim nCount As Integer
      Dim nProg As Integer
      Dim strSql As String
          
1         woLama = ""
2         startLama = ""
          Dim rteJMech As Double
          
3         If txt <> "" Then
4             txt = " AND " & txt
5         End If
          
6         LblKnow.Visible = True
7         DoEvents
8         Set rsTemp = objRepo.searchWorkTech(Format(Tgl1, "MM/DD/YYYY"), Format(Tgl2, "MM/DD/YYYY"), Ket, txt)
9         If Not rsTemp.EOF Then
10            nCount = 0
11            nProg = 0
12            btkTabel = createTemp(vNamaTabel)
13            Do While Not rsTemp.EOF
14                LblKnow.Visible = False
15                ProgressBar1.Visible = True
16                ProgressBar1.Min = 0
17                ProgressBar1.Max = rsTemp.RecordCount
18                nProg = nProg + 1
19                ProgressBar1.value = nProg
20                prosen.Caption = str(Round((nProg / rsTemp.RecordCount) * 100, 0)) + " %"
21                prosen.Visible = True
22                DoEvents
23                If btkTabel Then
24                    If woLama = Trim(rsTemp("NOORDER")) Then
25                        rate1 = 0
26                        rate2 = 0
27                    Else
28                        rate1 = CDbl(rsTemp("RATE1"))
29                        rate2 = CDbl(rsTemp("RATE2"))
30                    End If
                      
31                    If Log.repoHarianMech = 1 Then
                          Dim start As String
                          ''2016-09-16 || validasi telah finish untuk unit tersebut , jika telah finish dihitung rate, jika belum tidak dihitung rate
32                        Set rsDouble = objO.FetchData("select * from tempProgress where NOORDER = '" & Trim(rsTemp!noOrder) & "' and flagp <> 'M' order by TANGGAL DESC, START desc,FLAGP")
33                        If Not rsDouble.EOF Then
34                             rsDouble.MoveFirst
35                             If UCase(Trim(rsDouble!flagp)) <> "F" Then
                                  ''jika unit belum Finish
36                                rate1 = 0
37                                rate2 = 0
38                             Else
                                  ''jika telah finish , jika unit EM cari rate / jumlah mekanik yang kerja
                                  ''rsdouble posisi di flag finish CARI selain finish setelah
39                                If Trim(rsTemp!CODE) = "1" Then
40                                   start = "00:00"
41                                   totalPekerja = 0
          '                           If Trim(rsTemp!noOrder) = "200452" And Trim(rsTemp!mech) = "ALF" Then
          '                              totalPekerja = 0
          '                           End If
42                                   Do While Not rsDouble.EOF
43                                      If Trim(rsDouble!flagp) <> "F" And cariJumlahMenit(rsDouble!start) >= cariJumlahMenit(start) Then
44                                         totalPekerja = totalPekerja + 1
45                                         start = Trim(rsDouble!start)
46                                      End If
47                                      rsDouble.MoveNext
48                                   Loop
49                                   rate1 = Round((rate1 / totalPekerja), 2)
50                                   rate2 = Round((rate2 / totalPekerja), 2)
51                                End If
52                             End If
53                        End If
54                    End If
                      
55                    If startLama = Trim(rsTemp("START")) And stallLama = Trim(rsTemp("NOSTALL")) Then
56                        jamKrj = 0
57                    Else
58                        jamKrj = CDbl(Round(CDbl(rsTemp("JAMKERJA")) / 60, 2))
59                    End If
60                    resetJob = False
61                    Set rsDouble = objO.FetchData("SELECT NOWO FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NOWO = '" & Trim(rsTemp("NOORDER")) & "' AND JOB = '" & Trim(rsTemp("JOB")) & "' AND MECH = '" & Trim(rsTemp("MECH")) & "' AND TANGGAL = '" & Format(rsTemp("TANGGAL"), "MM/DD/YYYY") & "'")
62                    If Not rsDouble.EOF Then
63                        resetJob = True
64                    End If
65                    If rsDouble.State <> adStateClosed Then rsDouble.Close
                      
66                    If Log.repoHarianMech = 1 Then
67                        resetRate = True
68                        Set rsDouble = objO.FetchData("SELECT TOP " & IIf(totalPekerja = 0, 1, totalPekerja) & " MECH, DATEDIFF(n,START,PAUSE) FROM TEMPPROGRESS WHERE NOORDER = '" & Trim(rsTemp("NOORDER")) & "' AND FLAGP <> 'M' GROUP BY MECH, DATEDIFF(n,START,PAUSE) ORDER BY COUNT(*), DATEDIFF(n,START,PAUSE) DESC")
69                    Else
70                        resetRate = False
71                        Set rsDouble = objO.FetchData("SELECT TOP 1 MECH, DATEDIFF(n,START,PAUSE) FROM TEMPPROGRESS WHERE NOORDER = '" & Trim(rsTemp("NOORDER")) & "' AND FLAGP <> 'M' GROUP BY MECH, DATEDIFF(n,START,PAUSE) ORDER BY COUNT(*), DATEDIFF(n,START,PAUSE) DESC")
72                    End If
                      
73                    If Not rsDouble.EOF Then
74                        If Log.repoHarianMech = 1 Then
75                            rsDouble.MoveFirst
76                            Do While Not rsDouble.EOF
77                                If Trim(rsDouble("MECH")) = Trim(rsTemp("MECH")) Then
78                                    resetRate = False
79                                    Exit Do
80                                End If
81                                rsDouble.MoveNext
82                            Loop
83                        Else
84                            If Trim(rsDouble("MECH")) <> Trim(rsTemp("MECH")) Then
85                                resetRate = True
86                            End If
87                        End If
                          
88                    End If
89                    If rsDouble.State <> adStateClosed Then rsDouble.Close
                      
90                    doubleJob = False
                      'Set rsDouble = objO.FetchData("SELECT NOORDER FROM TEMPPROGRESS WHERE NOORDER = '" & Trim(rsTemp("NOORDER")) & "' AND TANGGAL < '" & Format(Tgl1, "MM/DD/YYYY") & "' AND FLAGP <> 'M'")
91                    Set rsDouble = objO.FetchData("SELECT NOORDER FROM TEMPPROGRESS WITH(NOLOCK) WHERE NOORDER = '" & Trim(rsTemp("NOORDER")) & "' AND TANGGAL < '" & Format(Tgl1, "MM/DD/YYYY") & "' AND FLAGP not in ('M','P')")
92                    If Not rsDouble.EOF Then
93                        doubleJob = True
94                    End If
95                    If rsDouble.State <> adStateClosed Then rsDouble.Close
                      
                      '' ALASAN JOB STOPPED
                      Dim rsStop As ADODB.Recordset
                      Dim ketStop As String
96                    ketStop = ""
97                    If Trim(rsTemp("PAUSE")) <> "" And Trim(rsTemp("FLAGP")) = "P" Then
98                        Set rsStop = objO.FetchData("SELECT * FROM PROGRESS WITH(NOLOCK) WHERE NOORDER = '" & Trim(rsTemp("NOORDER")) & "' AND FLAG = 'P' AND JAM = '" & Trim(rsTemp("PAUSE")) & "'")
99                        If Not rsStop.EOF Then
100                           ketStop = Trim(Mid(rsStop("KETERANGAN"), InStr(rsStop("KETERANGAN"), "|") + 1, Len(rsStop("KETERANGAN")) - InStr(rsStop("KETERANGAN"), "|")))
101                       End If
102                       If rsStop.State <> adStateClosed Then rsStop.Close
103                   End If
                      
104                   If Not resetJob Or jamKrj > 0 Then
105                       If mekanik(Trim(rsTemp("MECH"))) Then
106                           If jamKrj > 0 Then nCount = nCount + 1
107                           Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (" & CDbl(nCount) & ",'" & Trim(rsTemp("NOORDER")) & "', '" & Format(rsTemp("TANGGAL"), "MM/DD/YYYY") & "', '" & Trim(rsTemp("NOSTALL")) & "', '" & Trim(rsTemp("START")) & "', " _
                                           & "'" & Trim(rsTemp("FLAGP")) & "', '" & Trim(rsTemp("PAUSE")) & "', '" & Trim(rsTemp("MECH")) & "', '" & Trim(rsTemp("NAME")) & "', '" & Trim(rsTemp("FOREMAN")) & "', " & RepRp(Replace(jamKrj, ",", ".")) & ", '" & Trim(rsTemp("NO_POLISI")) & "', " _
                                           & "'" & Trim(rsTemp("NO_RANGKA")) & "', " & RepRp(IIf(resetRate, 0, IIf(doubleJob, 0, IIf(resetJob, 0, Replace(CDbl(rate1), ",", "."))))) & "," & RepRp(IIf(resetRate, 0, IIf(doubleJob, 0, IIf(resetJob, 0, Replace(CDbl(rate2), ",", "."))))) & ", " _
                                           & "'" & IIf(resetJob, "", Trim(rsTemp("GRP"))) & "','" & IIf(resetJob, "", Trim(rsTemp("JOB"))) & "','" & Trim(ketStop) & "')")
108                       End If
109                   End If
110                   woLama = Trim(rsTemp("NOORDER"))
111                   startLama = Trim(rsTemp("START"))
112                   stallLama = Trim(rsTemp("NOSTALL"))
113               End If
114               rsTemp.MoveNext
115           Loop
116           Call Cetak(vNamaTabel, txt)
117           Call DisableExport
118           Call setLokasiTombol
119       Else
120           MsgBox "Tidak ada data !!"
121           Exit Sub
122       End If
End Sub

Private Function mekanik(namaMekanik As String) As Boolean
    Dim rsSearch As ADODB.Recordset
    mekanik = False
    Set rsSearch = objO.FetchData("SELECT MECH FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(namaMekanik) & "'")
    If Not rsSearch.EOF Then
        mekanik = True
    End If
End Function


Private Sub Cetak(vNamaTabel As String, txt As String)
Dim oRs As Object
Set Report = New rptRepoHarianTech1
Screen.MousePointer = vbHourglass
    
    Set oRs = objRepo.printRepo(vNamaTabel)
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
    prosen.Visible = False
    ProgressBar1.Visible = False
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
    f.show (vbModal)
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
            .show vbModal
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

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

Private Sub Form_Resize()
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

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub Form_Load()
    onTimer = False
    Call hapusForm
    Call isiCombo
    vNamaTabel = "HarianTech" & DriveSerialNumber("c:")
End Sub

Private Sub TblKeluar_Click()
    onTimer = True
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    Unload Me
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean
Dim txtTemp As String
Dim txt As String
    
    boleh = False
    txtTemp = ""
    txt = ""
    ProgressBar1.Visible = False
    prosen.Visible = False
    If Text(0) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "FOREMAN = '" & Text(0) & "' "
    End If
    If Text(2) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "MECH = '" & Text(2) & "' "
    End If
    
    If Trim(Text(4)) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "NOSTALL = 'STALL " & Trim(Text(4)) & "' "
    End If

    If opCetak(0).value Then
        Call IsiTempRepo(txtTemp, vNamaTabel, txt)
        
    ElseIf opCetak(1).value Then
        Call harianMech
    ElseIf opCetak(2).value Then
        Call groupMech
    End If
End Sub

Private Sub harianMech()
      'Dim Report As New rptRepoHarianTech2
1         Set Report = New rptRepoHarianTech2
2         LblKnow.Visible = True
3         DoEvents
4         ProgressBar1.Visible = True
5         ProgressBar1.value = 0
6         ProgressBar1.Max = DateDiff("d", Tgl1, Tgl2) + 3
          Dim Rate         As Double
          Dim start        As String
          Dim totalPekerja As Integer
          Dim totalJT      As Double
          Dim totalJA      As Double
          Dim JF_JA        As Double
          Dim JF_JT        As Double
          
7         Call objO.UpdateData("Delete from rekapMech")
          Dim nTgl As Date
8         nTgl = Format(Tgl1, "DD/MM/YYYY")
          
9             LblKnow.Caption = "Collect Data ... " & Format(nTgl, "DD/MM/YYYY")
10        DoEvents
          Dim i As Integer
11        i = 1
12        Do Until Format(nTgl, "DD/MM/YYYY") = Format(DateAdd("d", 1, Tgl2), "DD/MM/YYYY")
13            ProgressBar1.value = i
14            LblKnow.Caption = "Collect Data ... " & Format(nTgl, "DD/MM/YYYY")
15            DoEvents
16            Call objO.InsertData("Insert into RekapMech " & "EXEC rekapHarianMech '" & Format(nTgl, "MM/DD/YYYY") & "', '" & Format(nTgl, "MM/DD/YYYY") & "', '" & Trim(Text(2)) & "', '" & Format(isnow, "HH:NN") & "'")
              
17            If Log.repoHarianMech = 1 Then
                   ''sSQL = "SELECT DISTINCT NOORDER ,RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7+RTE8+(SELECT ISNULL(SUM(RTE),0) FROM JOBADD2 WHERE NO_WO = t.NOORDER) AS JRATE, j.CODE FROM TEMPPROGRESS t WITH (NOLOCK) LEFT JOIN JOBDTA j WITH (NOLOCK) ON j.NO_WO = t.NOORDER WHERE t.MECH = '" & Trim(text(2)) & "' AND t.TANGGAL = '" & Format(nTgl, "MM/DD/YYYY") & "'"
                   ''Set rsTemp = objO.FetchData("SELECT DISTINCT NOORDER FROM TEMPPROGRESS WITH (NOLOCK) WHERE MECH = '" & Trim(Text(2)) & "' AND TANGGAL = '" & Format(nTgl, "MM/DD/YYYY") & "' ")
                   
                   '' ----------------------------------------------
                   '' 20190314 | REV | ORDER PREDIAGNOSE JUGA DIHITUNG
                   '' --------------------------------------------------------------
18                 sSQL = "SELECT DISTINCT NOORDER ,"
19                 sSQL = sSQL & " CASE WHEN isnull(j.CODE,'') <> '' then"
20                 sSQL = sSQL & " j.RTE1+j.RTE2+j.RTE3+j.RTE4+j.RTE5+j.RTE6+j.RTE7+j.RTE8+"
21                 sSQL = sSQL & " (SELECT ISNULL(SUM(RTE),0) FROM JOBADD2 WHERE NO_WO = t.NOORDER)"
22                 sSQL = sSQL & " Else"
23                 sSQL = sSQL & " D.RTE1+D.RTE2+D.RTE3+D.RTE4+D.RTE5+D.RTE6+D.RTE7+D.RTE8+"
24                 sSQL = sSQL & " (SELECT ISNULL(SUM(RTE),0) FROM JOBADD2 WHERE NO_WO = t.NOORDER)"
25                 sSQL = sSQL & " End"
26                 sSQL = sSQL & " AS JRATE,"
27                 sSQL = sSQL & " CASE WHEN isnull(j.CODE,'') <> '' then"
28                 sSQL = sSQL & " j.CODE"
29                 sSQL = sSQL & " Else"
30                 sSQL = sSQL & " d.CODE"
31                 sSQL = sSQL & " end As CODE"
32                 sSQL = sSQL & " FROM TEMPPROGRESS t WITH (NOLOCK)"
33                 sSQL = sSQL & " LEFT JOIN JOBDTA j WITH (NOLOCK) ON j.NO_WO = t.NOORDER"
34                 sSQL = sSQL & " LEFT JOIN DIAGDTA D WITH (NOLOCK) ON D.NO_DIAG = T.NOORDER"
      '             sSQL = sSQL & " WHERE t.MECH = '" & Trim(Text(2)) & "'"
35                 sSQL = sSQL & " WHERE t.MECH = '" & Trim(Text(2)) & "' AND t.flagp <> 'M' "
36                 sSQL = sSQL & " AND t.TANGGAL = '" & Format(nTgl, "MM/DD/YYYY") & "'"
                   
37                 Set rsTemp = objO.FetchData(sSQL)
38                 totalFlate = 0
39                 If Not rsTemp.EOF Then
40                     rsTemp.MoveFirst
41                     Do While Not rsTemp.EOF
                          ''Set rsHelp = objO.FetchData("select * from tempProgress WITH (NOLOCK) where NOORDER = '" & Trim(rsTemp!noOrder) & "' and flagp <> 'M' order by START desc,FLAGP")
                          ''20201123 req mb thania
42                        Set rsHelp = objO.FetchData("select * from tempProgress WITH (NOLOCK) where NOORDER = '" & Trim(rsTemp!noOrder) & "' and flagp <> 'M' order by tanggal desc, START desc,FLAGP")
43                        start = "00:00"
44                        Rate = CDbl(rsTemp!Jrate)
45                        totalPekerja = 0
                          
46                        If Not rsHelp.EOF Then
47                           rsHelp.MoveFirst
48                           If UCase(Trim(rsHelp!flagp)) <> "F" Then
                                ''jika unit belum Finish
49                              Rate = 0
50                           Else
                                ''jika telah finish , jika unit EM cari rate / jumlah mekanik yang kerja
                                ''rsdouble posisi di flag finish CARI selain finish setelah
51                              If Trim(rsTemp!CODE) = "1" Then
52                                 start = "00:00"
53                                 totalPekerja = 0
              '                           If Trim(rsTemp!noOrder) = "200452" And Trim(rsTemp!mech) = "ALF" Then
              '                              totalPekerja = 0
              '                           End If
54                                 rsHelp.MoveFirst
55                                 Do While Not rsHelp.EOF
56                                    If Trim(rsHelp!flagp) <> "F" And cariJumlahMenit(rsHelp!start) >= cariJumlahMenit(start) Then
57                                       totalPekerja = totalPekerja + 1
58                                       start = Trim(rsHelp!start)
59                                    End If
60                                    rsHelp.MoveNext
61                                 Loop
62                                 Rate = Round((Rate / totalPekerja), 2)
                                   
63                              End If
64                           End If
65                        End If
66                        totalFlate = totalFlate + Rate
                          
67                        rsTemp.MoveNext
68                     Loop
69                 End If
                   
                   ''mengambil nilai jt
70                 totalJT = objO.FetchData_Double("select ROUND(isnull((SELECT (CAST(ISNULL(MENITKERJA,0) AS FLOAT)/60) FROM T_MEKANIK WHERE TANGGAL = '" & Format(nTgl, "DD/MM/YYYY") & "' AND MECH = '" & Trim(Text(2)) & "'),0),2) jt")
                   ''mengambil nilai ja
71                 totalJA = objO.FetchData_Double("select ROUND(isnull((SELECT (CAST(ISNULL(SUM(DATEDIFF(n,START,CASE WHEN PAUSE = '' THEN '" & Format(isnow, "HH:NN") & "' ELSE PAUSE END)),0) AS FLOAT)/60) FROM TEMPPROGRESS WHERE TANGGAL = '" & Format(nTgl, "MM/DD/YYYY") & "' AND MECH = '" & Trim(Text(2)) & "'),0),2) AS JA")
                   
72                 If totalJT <> 0 Then
73                     JF_JT = CDbl(totalFlate / totalJT)
74                 Else
75                     JF_JT = 0
76                 End If
77                 If totalJA <> 0 Then
78                     JF_JA = CDbl(totalFlate / totalJA)
79                 Else
80                     JF_JA = 0
81                 End If
                   
                   ''update data rekapMech untuk mekanik dan tanggal dengan rate yang dihitung dari em stall
82                 Call objO.InsertData("UPDATE RekapMech SET JF = " & Replace(CDbl(totalFlate), ",", ".") & ",JF_JT = " & Replace(JF_JT, ",", ".") & ",JF_JA = " & Replace(JF_JA, ",", ".") & " WHERE TANGGAL = '" & Format(nTgl, "MM/DD/YYYY") & "' AND MECH = '" & Trim(Text(2)) & "'")
                   
83            End If ''XXXXXXX End log repo harian mech XXXXXXXXXXXXXXX
              
              
84            nTgl = DateAdd("d", i, Tgl1)
85            i = i + 1
86        Loop

87        Set oRs = objO.FetchData("Select * from rekapMech WITH(NOLOCK)")

88        ProgressBar1.Visible = False
89        DoEvents
          'sSQL = "EXEC rekapHarianMech '" & Format(Tgl1, "MM/DD/YYYY") & "', '" & Format(Tgl2, "MM/DD/YYYY") & "', '" & Trim(Text(2)) & "', '" & Format(isnow, "HH:NN") & "'"
          
          'Set oRs = objO.FetchData(sSQL)
90        Report.database.SetDataSource oRs
              
          
          Dim rsFO As ADODB.Recordset
91        Set rsFO = oRs.Clone
92        If Not rsFO.EOF Then
93            rsFO.Find "FO <> ''", , adSearchForward, 1
94            On Error GoTo nulFO
95            Report.txtFO.SetText Trim(rsFO("FO"))
nulFO:
96        Else
97            Report.txtFO.SetText ""
98        End If
99        Report.txtPerusahaan.SetText gPerush
100       Report.txtAlamat.SetText gAlamat
101       Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
         
102       With CRViewer1
103           .ReportSource = Report
104           .ViewReport
105           .Zoom 100
106           .Visible = True
107       End With
108       Call DisableExport
109       Call setLokasiTombol
110       cmdTutup.Visible = True
111       cmdSettingPrinter.Visible = True
112       Screen.MousePointer = vbDefault
113       LblKnow.Visible = False
End Sub

Private Sub createRepoGroup()
1        On Error GoTo createRepoGroup_Error

2         If rsGroup.State Then rsGroup.Close
3         Set rsGroup = Nothing
4         With rsGroup
5             .Fields.Append "DATE", adDate 'adVarChar, 25
6             .Fields.Append "NOINV", adVarChar, 25
7             .Fields.Append "NO_POLISI", adVarChar, 25
8             .Fields.Append "NO_WO", adVarChar, 25
9             .Fields.Append "GRP", adVarChar, 25
10            .Fields.Append "MECH", adVarChar, 25
11            .Fields.Append "NAME", adVarChar, 100
12            .Fields.Append "NO_POLISI2", adVarChar, 25
13            .Fields.Append "JRATE", adDouble
14            .Fields.Append "ARATE", adDouble
15            .Fields.Append "NOMINAL", adDouble
              
16            .Open
17        End With

18       On Error GoTo 0
19       Exit Sub

createRepoGroup_Error:

20       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRepoGroup of Form frmRepoHarianTech")
         
           
End Sub
Private Sub groupMech()
      'Dim Report As New rptRepoGroupTech
1         Set Report = New rptRepoGroupTech
          Dim start            As String
          Dim rate1            As Double
          Dim totalPekerja     As Integer
          Dim mekanik(10)      As String
2         On Error Resume Next
3         LblKnow.Visible = True
4         DoEvents
          'sSQL = "EXEC rekapGroupMech '" & Format(Tgl1, "MM/DD/YYYY") & "', '" & Format(Tgl2, "MM/DD/YYYY") & "', '" & Trim(Text(2)) & "'"
5         If Log.repoHarianMech = 0 Then
6             sSQL = "EXEC rekapGroupMech2 '" & Format(Tgl1, "MM/DD/YYYY") & "', '" & Format(Tgl2, "MM/DD/YYYY") & "', '" & Trim(Text(2)) & "', '" & Trim(Text(0)) & "'"
7             Set oRs = objO.FetchData(sSQL)
8             Report.database.SetDataSource oRs
9         Else
10            sSQL = "EXEC rekapGroupMechEM '" & Format(Tgl1, "MM/DD/YYYY") & "', '" & Format(Tgl2, "MM/DD/YYYY") & "', '" & Trim(Text(2)) & "', '" & Trim(Text(0)) & "'"
11            Set oRs = objO.FetchData(sSQL)
12            If Not oRs.EOF Then
                ''2016-09-19 | penambahan validasi untuk wo em , dicek untuk mekanik yang mengerjakan
13              oRs.MoveFirst
14              Call createRepoGroup
15              Do While Not oRs.EOF
16                If Trim(oRs!CODE) = "1" Then
17                    For i = 0 To 10
18                       mekanik(i) = ""
19                    Next
20                    start = "00:00"
21                    totalPekerja = 0
22                    rate1 = CDbl(Trim(oRs!Jrate))
23                    Set rsTemp = objO.FetchData("select * from tempProgress where NOORDER = '" & Trim(oRs!no_wo) & "' and flagp <> 'M' order by START desc,FLAGP")
24                    If Not rsTemp.EOF Then
25                       rsTemp.MoveFirst
26                       If Trim(rsTemp!flagp) <> "F" Then
                            ''jika unit belum finish
27                          rate1 = 0
28                          With rsGroup
29                             .AddNew
30                             !Date = Format(Trim(oRs!Date), "mm/dd/yyyy")
31                             !noInv = Trim(oRs!noInv)
32                             !NO_POLISI = Trim(oRs!NO_POLISI)
33                             !no_wo = Trim(oRs!no_wo)
34                             !grp = Trim(oRs!grp)
35                             !Mech = Trim(oRs!Mech)
36                             !Name = Trim(oRs!Name)
37                             !no_polisi2 = Trim(oRs!NO_POLISI)
38                             !Jrate = CDbl(Trim(rate1))
39                             !arate = CDbl(Trim(oRs!arate))
40                             !nominal = CDbl(Trim(oRs!nominal))
                               
41                             .Update
42                          End With
43                       Else
44                          Do While Not rsTemp.EOF
45                             If Trim(rsTemp!flagp) <> "F" And cariJumlahMenit(rsTemp!start) >= cariJumlahMenit(start) Then
46                                mekanik(totalPekerja) = Trim(rsTemp!Mech)
47                                totalPekerja = totalPekerja + 1
                                  
48                                start = Trim(rsTemp!start)
49                             End If
50                             rsTemp.MoveNext
51                          Loop
52                          rate1 = Round((rate1 / totalPekerja), 2)
53                          For i = 0 To 10
54                             If mekanik(i) <> "" Then
55                                With rsGroup
56                                   .AddNew
57                                   !Date = Format(Trim(oRs!Date), "mm/dd/yyyy")
58                                   !noInv = Trim(oRs!noInv)
59                                   !NO_POLISI = Trim(oRs!NO_POLISI)
60                                   !no_wo = Trim(oRs!no_wo)
61                                   !grp = Trim(oRs!grp)
62                                   !Mech = Trim(mekanik(i))
63                                   !Name = objO.FetchData_String("select TOP 1 name from MEKANIK WITH (NOLOCK) where MECH = '" & Trim(mekanik(i)) & "'")
64                                   !no_polisi2 = Trim(oRs!NO_POLISI)
65                                   !Jrate = CDbl(Trim(rate1))
66                                   !arate = CDbl(Trim(oRs!arate))
67                                   !nominal = CDbl(Trim(oRs!nominal))
                                     
68                                   .Update
69                                End With
70                             Else
71                                Exit For
72                             End If
73                          Next
                            
74                       End If
                         
75                    End If
                      
76                Else
77                    With rsGroup
78                       .AddNew
79                       !Date = Format(Trim(oRs!Date), "mm/dd/yyyy")
80                       !noInv = Trim(oRs!noInv)
81                       !NO_POLISI = Trim(oRs!NO_POLISI)
82                       !no_wo = Trim(oRs!no_wo)
83                       !grp = Trim(oRs!grp)
84                       !Mech = Trim(oRs!Mech)
85                       !Name = Trim(oRs!Name)
86                       !no_polisi2 = Trim(oRs!NO_POLISI)
87                       !Jrate = CDbl(Trim(oRs!Jrate))
88                       !arate = CDbl(Trim(oRs!arate))
89                       !nominal = CDbl(Trim(oRs!nominal))
                         
90                       .Update
91                    End With
92                End If
93                oRs.MoveNext
94              Loop
95            End If
96            Report.database.SetDataSource rsGroup
97        End If
98        Report.txtPerusahaan.SetText gPerush
99        Report.txtAlamat.SetText gAlamat
100       Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
         
101       With CRViewer1
102           .ReportSource = Report
103           .ViewReport
104           .Zoom 100
105           .Visible = True
106       End With
107       Call DisableExport
108           Call setLokasiTombol
109       cmdTutup.Visible = True
110       cmdSettingPrinter.Visible = True
111       Screen.MousePointer = vbDefault
112       LblKnow.Visible = False
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
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "FOREMAN"
                frmCari.show (vbModal)
            Case 2
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ALL"
                gBantuan4 = " WHERE [GROUP] = '" & Trim(Text(0)) & "'"
                frmCari.show (vbModal)
            Case 4
                gMainKeyCari = Text(Index)
                gSearchOption = "STALL"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT DISTINCT FS_KD_PETUGAS, FS_NM_PETUGAS FROM T_PETUGAS WITH(NOLOCK) WHERE FS_KD_PETUGAS = '" & Trim(Text(Index)) & "' ")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "FOREMAN"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = oRs("FS_NM_PETUGAS")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
        Case 2
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Text(Index) & "' ")
                If Not oRs.EOF Then
                    Text(Index + 1) = oRs("NAME")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
    End Select
End Sub

Private Function cariJumlahMenit(jam As String) As Double
1     On Error GoTo vb_error
          Dim jJam As Integer
          Dim jMenit As Integer
2         jam = Trim(jam)
3         jJam = CInt(Left(jam, InStr(1, jam, ":") - 1))
4         jMenit = CInt(Right(jam, InStr(1, jam, ":") - 1))
          
5         cariJumlahMenit = CDbl(((jJam * 60) + (jMenit)))
6         Exit Function
vb_error:
7         MsgBox "Line Number " & err.Number & vbCrLf & " descripsi : " & err.Description & vbCrLf & "Fungsi cari Jumlah Menit", vbInformation
End Function
