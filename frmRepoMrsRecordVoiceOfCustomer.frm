VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoMrsRecordVoiceOfCustomer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan MRS untuk Record Voice Of Customer"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   12735
   Begin ComctlLib.ProgressBar progBar 
      Height          =   585
      Left            =   615
      TabIndex        =   26
      Top             =   5400
      Visible         =   0   'False
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   1032
      _Version        =   327682
      Appearance      =   1
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
      Left            =   9465
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2355
      Visible         =   0   'False
      Width           =   1185
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
      Left            =   1455
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2325
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
      Left            =   2565
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2325
      Width           =   1440
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
      Left            =   4050
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2325
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   6375
      TabIndex        =   9
      Top             =   2355
      Visible         =   0   'False
      Width           =   3015
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   270
      Left            =   5325
      TabIndex        =   10
      Top             =   2325
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
      MICON           =   "frmRepoMrsRecordVoiceOfCustomer.frx":0000
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
      Height          =   8055
      Left            =   8595
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   4620
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
   Begin TDBDate6Ctl.TDBDate DtDOAwal 
      Height          =   330
      Left            =   1785
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmRepoMrsRecordVoiceOfCustomer.frx":001C
      Caption         =   "frmRepoMrsRecordVoiceOfCustomer.frx":0148
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoMrsRecordVoiceOfCustomer.frx":01B4
      Keys            =   "frmRepoMrsRecordVoiceOfCustomer.frx":01D2
      Spin            =   "frmRepoMrsRecordVoiceOfCustomer.frx":0230
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
   Begin TDBDate6Ctl.TDBDate DtDOAkhir 
      Height          =   330
      Left            =   3825
      TabIndex        =   2
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmRepoMrsRecordVoiceOfCustomer.frx":0258
      Caption         =   "frmRepoMrsRecordVoiceOfCustomer.frx":0384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoMrsRecordVoiceOfCustomer.frx":03F0
      Keys            =   "frmRepoMrsRecordVoiceOfCustomer.frx":040E
      Spin            =   "frmRepoMrsRecordVoiceOfCustomer.frx":046C
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
   Begin ucXPButton.XPButton cmdCetak 
      Height          =   495
      Left            =   5850
      TabIndex        =   6
      Top             =   1200
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
      MICON           =   "frmRepoMrsRecordVoiceOfCustomer.frx":0494
      PICN            =   "frmRepoMrsRecordVoiceOfCustomer.frx":04B0
      PICH            =   "frmRepoMrsRecordVoiceOfCustomer.frx":0907
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdKeluar 
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   5355
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
      MICON           =   "frmRepoMrsRecordVoiceOfCustomer.frx":0DF1
      PICN            =   "frmRepoMrsRecordVoiceOfCustomer.frx":0E0D
      PICH            =   "frmRepoMrsRecordVoiceOfCustomer.frx":12E1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBDate6Ctl.TDBDate recMappAwal 
      Height          =   330
      Left            =   1815
      TabIndex        =   14
      Top             =   2385
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmRepoMrsRecordVoiceOfCustomer.frx":182C
      Caption         =   "frmRepoMrsRecordVoiceOfCustomer.frx":1958
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoMrsRecordVoiceOfCustomer.frx":19C4
      Keys            =   "frmRepoMrsRecordVoiceOfCustomer.frx":19E2
      Spin            =   "frmRepoMrsRecordVoiceOfCustomer.frx":1A40
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
   Begin TDBDate6Ctl.TDBDate recMappAkhir 
      Height          =   330
      Left            =   3855
      TabIndex        =   15
      Top             =   2385
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmRepoMrsRecordVoiceOfCustomer.frx":1A68
      Caption         =   "frmRepoMrsRecordVoiceOfCustomer.frx":1B94
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoMrsRecordVoiceOfCustomer.frx":1C00
      Keys            =   "frmRepoMrsRecordVoiceOfCustomer.frx":1C1E
      Spin            =   "frmRepoMrsRecordVoiceOfCustomer.frx":1C7C
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
   Begin ucXPButton.XPButton cmdPrintMapping 
      Height          =   495
      Left            =   5790
      TabIndex        =   16
      Top             =   2970
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
      MICON           =   "frmRepoMrsRecordVoiceOfCustomer.frx":1CA4
      PICN            =   "frmRepoMrsRecordVoiceOfCustomer.frx":1CC0
      PICH            =   "frmRepoMrsRecordVoiceOfCustomer.frx":2117
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBDate6Ctl.TDBDate tglBoard1 
      Height          =   330
      Left            =   1875
      TabIndex        =   21
      Top             =   4110
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmRepoMrsRecordVoiceOfCustomer.frx":2601
      Caption         =   "frmRepoMrsRecordVoiceOfCustomer.frx":272D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoMrsRecordVoiceOfCustomer.frx":2799
      Keys            =   "frmRepoMrsRecordVoiceOfCustomer.frx":27B7
      Spin            =   "frmRepoMrsRecordVoiceOfCustomer.frx":2815
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
   Begin TDBDate6Ctl.TDBDate tglBoard2 
      Height          =   330
      Left            =   3915
      TabIndex        =   22
      Top             =   4110
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmRepoMrsRecordVoiceOfCustomer.frx":283D
      Caption         =   "frmRepoMrsRecordVoiceOfCustomer.frx":2969
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoMrsRecordVoiceOfCustomer.frx":29D5
      Keys            =   "frmRepoMrsRecordVoiceOfCustomer.frx":29F3
      Spin            =   "frmRepoMrsRecordVoiceOfCustomer.frx":2A51
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
   Begin ucXPButton.XPButton cmdPrintBoard 
      Height          =   495
      Left            =   5760
      TabIndex        =   23
      Top             =   4680
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
      MICON           =   "frmRepoMrsRecordVoiceOfCustomer.frx":2A79
      PICN            =   "frmRepoMrsRecordVoiceOfCustomer.frx":2A95
      PICH            =   "frmRepoMrsRecordVoiceOfCustomer.frx":2EEC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: RECORD VOICE OF CUSTOMER BOARD MRA::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1290
      TabIndex        =   24
      Top             =   3690
      Width           =   5085
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      X1              =   7035
      X2              =   540
      Y1              =   3525
      Y2              =   3525
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000080FF&
      Height          =   975
      Left            =   570
      Top             =   3630
      Width           =   6420
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal  :"
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
      Left            =   810
      TabIndex        =   25
      Top             =   4110
      Width           =   945
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   570
      Shape           =   4  'Rounded Rectangle
      Top             =   3630
      Width           =   6435
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: RECORD VOICE OF CUSTOMER RECOVERY MAPPING::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   765
      TabIndex        =   20
      Top             =   1965
      Width           =   6015
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   1905
      Width           =   6435
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal  :"
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
      Left            =   750
      TabIndex        =   19
      Top             =   2385
      Width           =   945
   End
   Begin VB.Label Label5 
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
      Left            =   3375
      TabIndex        =   18
      Top             =   2430
      Width           =   345
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: RECORD VOICE OF CUSTOMER::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   585
      TabIndex        =   17
      Top             =   1950
      Width           =   4875
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      Height          =   975
      Left            =   510
      Top             =   1905
      Width           =   6420
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   6990
      X2              =   585
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   6975
      X2              =   480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   975
      Left            =   480
      Top             =   120
      Width           =   6555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: RECORD VOICE OF CUSTOMER::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   555
      TabIndex        =   5
      Top             =   165
      Width           =   4875
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
      Left            =   3345
      TabIndex        =   4
      Top             =   645
      Width           =   345
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal  :"
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
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   945
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6540
   End
End
Attribute VB_Name = "frmRepoMrsRecordVoiceOfCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As Object
Dim rsLaporan As New ADODB.Recordset

Private Sub cmdCetak_Click()
      Dim rsRepo As ADODB.Recordset
1     Set Report = New rptMRSVoiceOfCustomer
          
2         If Trim(DtDOAwal) <> "" And Trim(DtDOAkhir) <> "" Then
3             sSQL = "SELECT DATE, " & _
                     "(SELECT COUNT(DISTINCT NO_CHASSIS) FROM JOBDTA A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC INNER JOIN M_ESTIMASI C WITH(NOLOCK) ON A.NO_CHASSIS = C.NORANGKA AND C.SBENEXT = B.FS_NILAI " & _
                     "WHERE ISNULL(JT_BLN,'') <> '' AND ISNULL(JT_TGL,'')<> ''  AND A.DATE = Q.DATE AND (DATEDIFF(d, (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), DATE) > 15 OR (KM-KMNEXT) > 1000)) AS early, " & _
                     "(SELECT COUNT(DISTINCT NO_CHASSIS) FROM JOBDTA A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC INNER JOIN M_ESTIMASI C WITH(NOLOCK) ON A.NO_CHASSIS = C.NORANGKA AND C.SBENEXT = B.FS_NILAI " & _
                     "WHERE ISNULL(JT_BLN,'') <> '' AND ISNULL(JT_TGL,'')<> ''  AND A.DATE = Q.DATE AND (ABS(DATEDIFF(d, (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), DATE)) <= 15 AND ABS(KMNEXT-KM) <= 1000)) AS akurat, " & _
                     "(SELECT COUNT(DISTINCT NO_CHASSIS) FROM JOBDTA A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC INNER JOIN M_ESTIMASI C WITH(NOLOCK) ON A.NO_CHASSIS = C.NORANGKA AND C.SBENEXT = B.FS_NILAI " & _
                     "WHERE ISNULL(JT_BLN,'') <> '' AND ISNULL(JT_TGL,'')<> ''  AND A.DATE = Q.DATE AND (DATEDIFF(d, (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END), DATE) < 0 OR (KMNEXT-KM) > 1000)) AS late " & _
                     "FROM JOBDTA Q WITH(NOLOCK) WHERE DATE BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' GROUP BY DATE ORDER BY DATE"

      '        sSQL = "select b.keteranganTipe as judul,b.pesan ,COUNT(a.LastNoWO) AS total"
      '        sSQL = sSQL & " from M_ListReminderCall a"
      '        sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
      '        sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "'"
      '        sSQL = sSQL & " group by b.keteranganTipe,b.pesan"
              
              ''2017 revisi untuk nilai 0 tetap ditampilkan
              ''-------------------------------------------------------
4             sSQL = "select a.keteranganTipe,a.pesan,isnull(b.total,0) total"
5             sSQL = sSQL & " from M_VoiceOfCustomer a"
6             sSQL = sSQL & " left join ("
7             sSQL = sSQL & " select b.keteranganTipe as judul,b.pesan ,COUNT(a.LastNoWO) AS total"
8             sSQL = sSQL & " from M_ListReminderCall a"
9             sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
10            sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "'"
11            sSQL = sSQL & " and a.NoRangka NOT in (select NoRangka from listReminderBoard where NoRangka = a.NoRangka and NoWo = a.LastNoWO)"
12            sSQL = sSQL & " group by b.keteranganTipe,b.pesan"
13            sSQL = sSQL & " ) b on b.judul = a.keteranganTipe and b.pesan = a.pesan"
14            sSQL = sSQL & " Union"
15            sSQL = sSQL & " select 'KPI','Total Data Tidak Terhubung' ,COUNT(a.LastNoWO) AS total"
16            sSQL = sSQL & " from M_ListReminderCall a"
17            sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
18            sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' and tipe = 0"
19            sSQL = sSQL & " and a.NoRangka NOT in (select NoRangka from listReminderBoard where NoRangka = a.NoRangka and NoWo = a.LastNoWO)"
20            sSQL = sSQL & " Union"
21            sSQL = sSQL & " select 'KPI','Total Data Terhubung' ,COUNT(a.LastNoWO) AS total"
22            sSQL = sSQL & " from M_ListReminderCall a"
23            sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
24            sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "' and tipe = 1"
25            sSQL = sSQL & " and a.NoRangka NOT in (select NoRangka from listReminderBoard where NoRangka = a.NoRangka and NoWo = a.LastNoWO)"
26            sSQL = sSQL & " Union"
27            sSQL = sSQL & " select 'KPI','Total Data' ,COUNT(a.LastNoWO) AS total"
28            sSQL = sSQL & " from M_ListReminderCall a"
29            sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
30            sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "'"
31            sSQL = sSQL & " and a.NoRangka NOT in (select NoRangka from listReminderBoard where NoRangka = a.NoRangka and NoWo = a.LastNoWO)"
              
32            Set rsRepo = objO.FetchData(sSQL)
33            Report.database.SetDataSource rsRepo
34            Report.txtPerusahaan.SetText gPerush
35            Report.TxtAlamat.SetText gAlamat
36            Report.txtTanggal.SetText "dari Tanggal " & TglDDMMMMYYYY(Format(DtDOAwal, "DD/MM/YYYY"), 1) & "    ~    " & TglDDMMMMYYYY(Format(DtDOAkhir, "DD/MM/YYYY"), 1)
              
37            With CRViewer1
38                .ReportSource = Report
39                .ViewReport
40                .Zoom 100
41                .Top = 0
42                .Left = 300 'Image1.Left + Image1.Width + 20
43                .Width = Me.Width - 500
44                .Height = Me.Height - 300 ' Image1.Height
                  
45                cmdTutup.Top = .Top + 430
46                cmdTutup.Left = .Left
47                cmdTutup.Visible = True
                  
48                cmdSettingPrinter.Top = cmdTutup.Top
49                cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 20
50                cmdSettingPrinter.Visible = True
51                .Visible = True
52            End With
53            cmdTutup.Visible = True
54            Screen.MousePointer = vbDefault
55            Call DisableExport
56            Call setLokasiTombol
57        End If
End Sub

Private Sub cmdPrintBoard_Click()
          Dim rsRepo As ADODB.Recordset
1     Set Report = New rptMRSVoiceOfCustomer
          
2         If Trim(tglBoard1) <> "" And Trim(tglBoard2) <> "" Then
               
              
              ''2017 revisi untuk nilai 0 tetap ditampilkan
              ''-------------------------------------------------------
3             sSQL = "select a.keteranganTipe,a.pesan,isnull(b.total,0) total"
4             sSQL = sSQL & " from M_VoiceOfCustomer a"
5             sSQL = sSQL & " left join ("
6             sSQL = sSQL & " select b.keteranganTipe as judul,b.pesan ,COUNT(a.LastNoWO) AS total"
7             sSQL = sSQL & " from   a"
8             sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
9             sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(tglBoard1, "MM/DD/YYYY") & "' AND '" & Format(tglBoard2, "MM/DD/YYYY") & "'"
10            sSQL = sSQL & " and a.NoRangka in (select NoRangka from listReminderBoard where NoRangka = a.NoRangka and NoWo = a.LastNoWO)"
11            sSQL = sSQL & " group by b.keteranganTipe,b.pesan"
12            sSQL = sSQL & " ) b on b.judul = a.keteranganTipe and b.pesan = a.pesan"
13            sSQL = sSQL & " Union"
14            sSQL = sSQL & " select 'KPI','Total Data Tidak Terhubung' ,COUNT(a.LastNoWO) AS total"
15            sSQL = sSQL & " from M_ListReminderCall a"
16            sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
17            sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(tglBoard1, "MM/DD/YYYY") & "' AND '" & Format(tglBoard2, "MM/DD/YYYY") & "' and tipe = 0"
18            sSQL = sSQL & " and a.NoRangka in (select NoRangka from listReminderBoard where NoRangka = a.NoRangka and NoWo = a.LastNoWO)"
19            sSQL = sSQL & " Union"
20            sSQL = sSQL & " select 'KPI','Total Data Terhubung' ,COUNT(a.LastNoWO) AS total"
21            sSQL = sSQL & " from M_ListReminderCall a"
22            sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
23            sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(tglBoard1, "MM/DD/YYYY") & "' AND '" & Format(tglBoard2, "MM/DD/YYYY") & "' and tipe = 1"
24            sSQL = sSQL & " and a.NoRangka in (select NoRangka from listReminderBoard where NoRangka = a.NoRangka and NoWo = a.LastNoWO)"
25            sSQL = sSQL & " Union"
26            sSQL = sSQL & " select 'KPI','Total Data' ,COUNT(a.LastNoWO) AS total"
27            sSQL = sSQL & " from M_ListReminderCall a"
28            sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
29            sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(tglBoard1, "MM/DD/YYYY") & "' AND '" & Format(tglBoard2, "MM/DD/YYYY") & "'"
30            sSQL = sSQL & " and a.NoRangka in (select NoRangka from listReminderBoard where NoRangka = a.NoRangka and NoWo = a.LastNoWO)"
              
              
31            Set rsRepo = objO.FetchData(sSQL)
              Dim rsCetak As New ADODB.Recordset
32            If rsCetak.State Then rsCetak.Close
33            With rsCetak
34                .Fields.Append "judul", adVarChar, 100
35                .Fields.Append "pesan", adVarChar, 200
36                .Fields.Append "total", adVarChar, 100
37                .Open
38            End With
              
39            If Not rsRepo.EOF Then
40                rsRepo.MoveFirst
41                progBar.Visible = True
42                progBar.Min = 1
43                progBar.Max = rsRepo.RecordCount + 6
                  Dim nilai As Integer
44                nilai = 1
45                DoEvents
46                Do While Not rsRepo.EOF
47                    With rsCetak
48                        .AddNew
49                        !judul = Nul(rsRepo.Fields(0).value)
50                        !pesan = Nul(rsRepo.Fields(1).value)
51                        !total = Nul(rsRepo.Fields(2).value)
52                        .Update
53                    End With
54                    progBar.value = nilai
55                    rsRepo.MoveNext
56                    nilai = nilai + 1
57                    DoEvents
58                Loop
                  
                  Dim rsFilter As New ADODB.Recordset
59                Set rsFilter = rsCetak.Clone
                  
                  Dim pembilang As Double
                  Dim penyebut As Double
                  
60                With rsCetak
61                    .AddNew
62                    pembilang = 0
63                    penyebut = 0
                      
64                    rsFilter.filter = ""
65                    rsFilter.filter = "pesan = 'Total Data Tidak Terhubung'"
66                    If Not rsFilter.EOF Then
67                        pembilang = rsFilter!total
68                    End If
                      
69                    rsFilter.filter = ""
70                    rsFilter.filter = "pesan = 'Total Data'"
71                    If Not rsFilter.EOF Then
72                        penyebut = rsFilter!total
73                    End If
74                    !judul = "KPI"
75                    !pesan = "Persentase Data Tidak Terhubung"
76                    If penyebut = 0 Then
77                        !total = 0
78                    Else
79                        !total = Round(((pembilang / penyebut) * 100), 0)
80                    End If
                      '!total = IIf(penyebut = 0, 0, Round(((pembilang / penyebut) * 100), 0))
81                    .Update
82                End With
                  
                  
83                With rsCetak
84                    .AddNew
85                    pembilang = 0
86                    penyebut = 0
                      
87                    rsFilter.filter = ""
88                    rsFilter.filter = "pesan = 'Total Data Terhubung'"
89                    If Not rsFilter.EOF Then
90                        pembilang = rsFilter!total
91                    End If
                      
92                    rsFilter.filter = ""
93                    rsFilter.filter = "pesan = 'Total Data'"
94                    If Not rsFilter.EOF Then
95                        penyebut = rsFilter!total
96                    End If
97                    !judul = "KPI"
98                    !pesan = "Persentase Data Terhubung"
99                    If penyebut = 0 Then
100                       !total = 0
101                   Else
102                       !total = Round(((pembilang / penyebut) * 100), 0)
103                   End If
                      '!total = IIf(penyebut = 0, 0, Round(((pembilang / penyebut) * 100), 0))
104                   .Update
105               End With
                  
                  
106               With rsCetak
107                   .AddNew
                      ''MENCARI DARI DAFTAR REMINDER YANG BIKIN WO DIBULAN BERIKUTNYA
108                   sSQL = "SELECT COUNT(b.NO_WO) as total"
109                   sSQL = sSQL & " FROM listReminderBoard a"
110                   sSQL = sSQL & " INNER JOIN JOBDTA b on a.NoRangka = b.NO_CHASSIS and b.FS_LOG_DEL = ''"
111                   sSQL = sSQL & " where CONVERT(varchar(20),a.tglentry,101) between '" & Format(tglBoard1, "mm/dd/yyyy") & "' and '" & Format(tglBoard1, "mm/dd/yyyy") & "'"
112                   sSQL = sSQL & " and ((YEAR(b.date) = '" & IIf(Month(tglBoard1) = "12", year(tglBoard1) + 1, year(tglBoard1)) & "' and MONTH(b.DATE) = '" & IIf(Month(tglBoard1) = "12", 1, Month(tglBoard1) + 1) & "') or (YEAR(b.date) = '" & IIf(Month(tglBoard2) = "12", year(tglBoard2) + 1, year(tglBoard2)) & "' and MONTH(b.DATE) = '" & IIf(Month(tglBoard2) = "12", 1, Month(tglBoard2) + 1) & "'))"
113                   !judul = "KPI"
114                   !pesan = "Come After Invited"
115                   !total = objO.FetchData_Double(sSQL)
116                   .Update
117               End With
                  
                  '''Nomor tidak aktif','Customer Deal & Booking on Call','tidak di angkat','Tidak bertemu pemilik','Operasional Luar Kota','KM belum sampai','Waktu Belum Pasti','Pelayanan Buruk','Waktu service lama'
118               With rsCetak
119                   .AddNew
                      
120                   pembilang = 0
121                   penyebut = 0
                      
122                   rsFilter.filter = ""
123                   sSQL = "'Nomor tidak aktif','Customer Deal & Booking on Call','tidak di angkat','Tidak bertemu pemilik','Operasional Luar Kota','KM belum sampai','Waktu Belum Pasti','Pelayanan Buruk','Waktu service lama'"
124                   rsFilter.filter = "pesan = 'Nomor tidak aktif' or pesan = 'Customer Deal & Booking on Call' or pesan = 'Tidak di angkat' or pesan = 'Tidak bertemu yang bersangkutan' or pesan = 'Operasional Luar Kota' or pesan = 'KM belum memenuhi untuk service' or pesan = 'Waktu Belum Pasti' or pesan = 'Pelayanan Buruk' or pesan = 'Waktu service lama'"
125                   If Not rsFilter.EOF Then
                          'pembilang = rsFilter!total
126                       rsFilter.MoveFirst
127                       Do While Not rsFilter.EOF
128                           pembilang = pembilang + rsFilter!total
129                           rsFilter.MoveNext
130                       Loop
131                   End If
                      
132                   rsFilter.filter = ""
133                   rsFilter.filter = "pesan = 'Total Data'"
134                   If Not rsFilter.EOF Then
135                       penyebut = rsFilter!total
136                   End If
                      
137                   !judul = "KPI"
138                   !pesan = "Outstanding Reminder"
139                   If penyebut = 0 Then
140                       !total = 0
141                   Else
142                       !total = Round(((pembilang / penyebut) * 100), 0)
143                   End If
                      '!total = Round((pembilang / penyebut) * 100, 0)
144                   .Update
145               End With
                  
146               With rsCetak
147                   .AddNew
                      
148                   pembilang = 0
149                   penyebut = 0
                      
150                   rsFilter.filter = ""
151                   sSQL = "'Nomor tidak aktif','Customer Deal & Booking on Call','tidak di angkat','Tidak bertemu pemilik','Operasional Luar Kota','KM belum sampai','Waktu Belum Pasti','Pelayanan Buruk','Waktu service lama'"
152                   rsFilter.filter = "pesan = 'Nomor tidak aktif' or pesan = 'Customer Deal & Booking on Call' or pesan = 'Tidak di angkat' or pesan = 'Tidak bertemu yang bersangkutan' or pesan = 'Operasional Luar Kota' or pesan = 'KM belum memenuhi untuk service' or pesan = 'Waktu Belum Pasti' or pesan = 'Pelayanan Buruk' or pesan = 'Waktu service lama'"
153                   If Not rsFilter.EOF Then
                          'pembilang = rsFilter!total
154                       rsFilter.MoveFirst
155                       Do While Not rsFilter.EOF
156                           pembilang = pembilang + rsFilter!total
157                           rsFilter.MoveNext
158                       Loop
159                   End If
                      
160                   rsFilter.filter = ""
161                   rsFilter.filter = "pesan = 'Total Data'"
162                   If Not rsFilter.EOF Then
163                       penyebut = rsFilter!total
164                   End If
                      
165                   !judul = "KPI"
166                   !pesan = "Potential Customer Booking On Call"
167                   If penyebut = 0 Then
168                       !total = 0
169                   Else
170                       !total = Round(((pembilang / penyebut) * 100), 0)
171                   End If
                      '!total = Round((pembilang / penyebut) * 100, 0)
172                   .Update
173               End With
                  
174               With rsCetak
175                   .AddNew
176                   pembilang = 0
177                   penyebut = 0
                      
178                   rsFilter.filter = ""
179                   rsFilter.filter = "pesan = 'Customer Deal & Booking on Call'"
180                   If Not rsFilter.EOF Then
181                       pembilang = rsFilter!total
182                   End If
                      
183                   rsFilter.filter = ""
184                   rsFilter.filter = "pesan = 'Total Data Terhubung'"
185                   If Not rsFilter.EOF Then
186                       penyebut = rsFilter!total
187                   End If
188                   !judul = "KPI"
189                   !pesan = "Appointment Rate"
190                   If penyebut = 0 Then
191                       !total = 0
192                   Else
193                       !total = Round(((pembilang / penyebut) * 100), 0)
194                   End If
                      '!total = IIf(penyebut = 0, 0, Round(((pembilang / penyebut) * 100), 0))
195                   .Update
196               End With
                  
197           End If
198           progBar.Visible = False
199           Report.database.SetDataSource rsCetak
200           Report.txtPerusahaan.SetText gPerush
201           Report.TxtAlamat.SetText gAlamat
202           Report.txtTanggal.SetText "dari Tanggal " & TglDDMMMMYYYY(Format(tglBoard1, "DD/MM/YYYY"), 1) & "    ~    " & TglDDMMMMYYYY(Format(tglBoard2, "DD/MM/YYYY"), 1)
203           Report.Text800.SetText "--o0o-- LAPORAN RECORD VOICE OF CUSTOMER BOARD MRA --o0o--"
204           With CRViewer1
205               .ReportSource = Report
206               .ViewReport
207               .Zoom 100
208               .Top = 0
209               .Left = 300 'Image1.Left + Image1.Width + 20
210               .Width = Me.Width - 500
211               .Height = Me.Height - 300 ' Image1.Height
                  
212               cmdTutup.Top = .Top + 430
213               cmdTutup.Left = .Left
214               cmdTutup.Visible = True
                  
215               cmdSettingPrinter.Top = cmdTutup.Top
216               cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 20
217               cmdSettingPrinter.Visible = True
218               .Visible = True
219           End With
220           cmdTutup.Visible = True
221           Screen.MousePointer = vbDefault
222           Call DisableExport
223           Call setLokasiTombol
224       End If
End Sub

Private Sub cmdPrintMapping_Click()
Dim rsRepo As ADODB.Recordset
Set Report = New rptMRSVoiceOfCustomer
    
    If Trim(recMappAwal) <> "" And Trim(recMappAkhir) <> "" Then
 
'        sSQL = "select b.keteranganTipe as judul,b.pesan ,COUNT(a.NoRangka) AS total"
'        sSQL = sSQL & " from M_ListReminderCallMapping a"
'        sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
'        sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(recMappAwal, "MM/DD/YYYY") & "' AND '" & Format(recMappAkhir, "MM/DD/YYYY") & "'"
'        sSQL = sSQL & " group by b.keteranganTipe,b.pesan"
        ''2017-10-10
        ''-------------------------------------------------------
        sSQL = "select a.keteranganTipe,a.pesan,isnull(b.total,0) total"
        sSQL = sSQL & " from M_VoiceOfCustomer a"
        sSQL = sSQL & " left join ("
        sSQL = sSQL & " select b.keteranganTipe as judul,b.pesan ,COUNT(a.NoRangka) AS total"
        sSQL = sSQL & " from M_ListReminderCallMapping a"
        sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
        sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(recMappAwal, "MM/DD/YYYY") & "' AND '" & Format(recMappAkhir, "MM/DD/YYYY") & "'"
        
        sSQL = sSQL & " group by b.keteranganTipe,b.pesan"
        sSQL = sSQL & " ) b on b.judul = a.keteranganTipe and b.pesan = a.pesan"
        
        sSQL = sSQL & " Union"
        sSQL = sSQL & " select 'KPI','Total Data Tidak Terhubung' ,COUNT(a.NoRangka) AS total"
        sSQL = sSQL & " from M_ListReminderCallMapping a"
        sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
        sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(recMappAwal, "MM/DD/YYYY") & "' AND '" & Format(recMappAkhir, "MM/DD/YYYY") & "' and tipe = 0"
        sSQL = sSQL & " Union"
        sSQL = sSQL & " select 'KPI','Total Data Terhubung' ,COUNT(a.NoRangka) AS total"
        sSQL = sSQL & " from M_ListReminderCallMapping a"
        sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
        sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(recMappAwal, "MM/DD/YYYY") & "' AND '" & Format(recMappAkhir, "MM/DD/YYYY") & "' and tipe = 1"
        sSQL = sSQL & " Union"
        sSQL = sSQL & " select 'KPI','Total Data' ,COUNT(a.NoRangka) AS total"
        sSQL = sSQL & " from M_ListReminderCallMapping a"
        sSQL = sSQL & " inner join M_voiceofcustomer b on a.Pesan_Telp = b.pesan and a.K_Telp = b.tipe"
        sSQL = sSQL & " where convert(VARCHAR(20),a.TglEntry,101) between '" & Format(recMappAwal, "MM/DD/YYYY") & "' AND '" & Format(recMappAkhir, "MM/DD/YYYY") & "'"
        
        Set rsRepo = objO.FetchData(sSQL)
        Report.database.SetDataSource rsRepo
        Report.txtPerusahaan.SetText gPerush
        Report.TxtAlamat.SetText gAlamat
        Report.txtTanggal.SetText "dari Tanggal " & TglDDMMMMYYYY(Format(recMappAwal, "DD/MM/YYYY"), 1) & "    ~    " & TglDDMMMMYYYY(Format(recMappAkhir, "DD/MM/YYYY"), 1)
        Report.Text800.SetText "--o0o-- LAPORAN RECORD VOICE OF CUSTOMER RECOVERY MAPPING --o0o--"
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 100
            .Top = 0
            .Left = 300 'Image1.Left + Image1.Width + 20
            .Width = Me.Width - 500
            .Height = Me.Height - 300 ' Image1.Height
            
            cmdTutup.Top = .Top + 430
            cmdTutup.Left = .Left
            cmdTutup.Visible = True
            
            cmdSettingPrinter.Top = cmdTutup.Top
            cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 20
            cmdSettingPrinter.Visible = True
            .Visible = True
        End With
        cmdTutup.Visible = True
        Screen.MousePointer = vbDefault
        Call DisableExport
        Call setLokasiTombol
    End If
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

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DtDOAwal = Format(isnow, "DD/MM/YYYY")
    DtDOAkhir = Format(isnow, "DD/MM/YYYY")
    
    tglBoard1 = Format(isnow, "DD/MM/YYYY")
    tglBoard2 = Format(isnow, "DD/MM/YYYY")
    
    recMappAkhir = Format(isnow, "DD/MM/YYYY")
    recMappAwal = Format(isnow, "DD/MM/YYYY")
    
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
End Sub

Private Sub Form_Resize()
   ' Image1.Height = Me.Height - 50
End Sub


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 445
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 445
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 445
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 445
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 445
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdPrintOut.Top = CRViewer1.Top + 445
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

Private Sub XPButton2_Click()

End Sub
