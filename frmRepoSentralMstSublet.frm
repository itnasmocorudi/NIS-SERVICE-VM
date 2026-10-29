VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmRepoSentralMstSublet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Display Master Sublet"
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
      TabIndex        =   17
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
      MICON           =   "frmRepoSentralMstSublet.frx":0000
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      Picture         =   "frmRepoSentralMstSublet.frx":001C
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   7
      Top             =   0
      Width           =   445
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   0
      Top             =   4920
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
      Height          =   3735
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   7455
      Begin TDBDate6Ctl.TDBDate tglCreate1 
         Height          =   285
         Left            =   1650
         TabIndex        =   34
         Top             =   1715
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   503
         Calendar        =   "frmRepoSentralMstSublet.frx":8E46
         Caption         =   "frmRepoSentralMstSublet.frx":8F72
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoSentralMstSublet.frx":8FDE
         Keys            =   "frmRepoSentralMstSublet.frx":8FFC
         Spin            =   "frmRepoSentralMstSublet.frx":905A
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
         ValueVT         =   1
         Value           =   3.24037959698105E-315
         CenturyMode     =   0
      End
      Begin VB.ComboBox cbStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmRepoSentralMstSublet.frx":9082
         Left            =   1650
         List            =   "frmRepoSentralMstSublet.frx":908F
         TabIndex        =   31
         Top             =   1390
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   255
         Left            =   1680
         TabIndex        =   27
         Top             =   200
         Width           =   2895
         Begin VB.OptionButton optCreator 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optCreator 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dealer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   630
            TabIndex        =   29
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optCreator 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1470
            TabIndex        =   28
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   450
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Visible         =   0   'False
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   794
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.OptionButton opCetak 
         BackColor       =   &H00FFFFFF&
         Caption         =   " f"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   5
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton opCetak 
         BackColor       =   &H00FFFFFF&
         Caption         =   " f"
         Height          =   255
         Index           =   0
         Left            =   1650
         TabIndex        =   4
         Top             =   2280
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   945
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbPromo 
         Height          =   285
         Left            =   5010
         TabIndex        =   18
         Top             =   2235
         Visible         =   0   'False
         Width           =   2415
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbPekerjaan 
         Height          =   285
         Left            =   1650
         TabIndex        =   19
         Top             =   1125
         Width           =   2415
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbSupplier 
         Height          =   285
         Left            =   1650
         TabIndex        =   22
         Top             =   840
         Width           =   2415
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   6000
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   6120
         Top             =   600
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbDealer 
         Height          =   285
         Left            =   1650
         TabIndex        =   25
         Top             =   530
         Width           =   2415
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   375
         Left            =   5880
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   6000
         Top             =   1680
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin TDBDate6Ctl.TDBDate tglCreate2 
         Height          =   285
         Left            =   2880
         TabIndex        =   35
         Top             =   1710
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   503
         Calendar        =   "frmRepoSentralMstSublet.frx":90A1
         Caption         =   "frmRepoSentralMstSublet.frx":91CD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoSentralMstSublet.frx":9239
         Keys            =   "frmRepoSentralMstSublet.frx":9257
         Spin            =   "frmRepoSentralMstSublet.frx":92B5
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
         ValueVT         =   1
         Value           =   3.24037959698105E-315
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate tglPeriode 
         Height          =   285
         Index           =   0
         Left            =   1650
         TabIndex        =   36
         Top             =   2040
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   503
         Calendar        =   "frmRepoSentralMstSublet.frx":92DD
         Caption         =   "frmRepoSentralMstSublet.frx":9409
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoSentralMstSublet.frx":9475
         Keys            =   "frmRepoSentralMstSublet.frx":9493
         Spin            =   "frmRepoSentralMstSublet.frx":94F1
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
         ValueVT         =   1
         Value           =   3.24037959698105E-315
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate tglPeriode 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   37
         Top             =   2040
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   503
         Calendar        =   "frmRepoSentralMstSublet.frx":9519
         Caption         =   "frmRepoSentralMstSublet.frx":9645
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoSentralMstSublet.frx":96B1
         Keys            =   "frmRepoSentralMstSublet.frx":96CF
         Spin            =   "frmRepoSentralMstSublet.frx":972D
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
         ValueVT         =   1
         Value           =   3.24037959698105E-315
         CenturyMode     =   0
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Create"
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
         Index           =   1
         Left            =   690
         TabIndex        =   33
         Top             =   1760
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Index           =   1
         Left            =   975
         TabIndex        =   32
         Top             =   1455
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
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
         Left            =   945
         TabIndex        =   26
         Top             =   575
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Creator Master"
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
         Left            =   315
         TabIndex        =   24
         Top             =   230
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Index           =   0
         Left            =   870
         TabIndex        =   23
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sublet"
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
         Left            =   990
         TabIndex        =   21
         Top             =   1125
         Width           =   450
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promo"
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
         Left            =   4350
         TabIndex        =   20
         Top             =   2235
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   2655
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Periode"
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
         Index           =   0
         Left            =   615
         TabIndex        =   3
         Top             =   2040
         Width           =   795
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
         TabIndex        =   2
         Top             =   240
         Width           =   60
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
      TabIndex        =   9
      Top             =   4200
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
      MICON           =   "frmRepoSentralMstSublet.frx":9755
      PICN            =   "frmRepoSentralMstSublet.frx":9771
      PICH            =   "frmRepoSentralMstSublet.frx":9BC8
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
      TabIndex        =   10
      Top             =   4200
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
      MICON           =   "frmRepoSentralMstSublet.frx":A0B2
      PICN            =   "frmRepoSentralMstSublet.frx":A0CE
      PICH            =   "frmRepoSentralMstSublet.frx":A5A2
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
      Y1              =   4800
      Y2              =   4800
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
      TabIndex        =   8
      Top             =   4320
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
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8295
      Left            =   0
      Picture         =   "frmRepoSentralMstSublet.frx":AAED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoSentralMstSublet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim rsTemp As ADODB.Recordset
Dim Report As Object

Dim doubleJob As Boolean
Dim resetJob As Boolean
Dim resetRate As Boolean
Dim vNamaTabel As String
Dim objRepo As New clsRepo
Dim rsGroup    As New ADODB.Recordset
Dim rsLaporan  As New ADODB.Recordset

Private Sub hapusForm()
    Tgl1 = isnow
    Tgl2 = isnow
    LblKnow.Visible = False
    ProgressBar1.Visible = False
    prosen.Visible = False
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

Private Sub cbSupplier_Change()
    Call refreshPekerjaan
End Sub
 

Private Sub cbSupplier_Click()
    Call cbSupplier_Change
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

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
'    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
'         & " drop table [tempdb]..[" & vNamaTabel & "]"
'    Call objO.UpdateData(sSQL)
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

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub Form_Load()
    onTimer = False
    
    If gUstat = 0 Then
        Label3.Visible = False
        cbDealer.Visible = False
         
        Label10(0).Top = Label10(1).Top
        tglPeriode(0).Top = tglCreate1.Top
        tglPeriode(1).Top = tglCreate2.Top
        
        Label10(1).Top = Label1(1).Top
        tglCreate1.Top = cbStatus.Top
        tglCreate2.Top = cbStatus.Top
        
        Label1(1).Top = Label4.Top
        cbStatus.Top = cbPekerjaan.Top
        
        Label4.Top = Label1(0).Top
        cbPekerjaan.Top = cbSupplier.Top
        
        Label1(0).Top = Label3.Top
        cbSupplier.Top = cbDealer.Top
        
         
    End If
    
    Call hapusForm
    Call refreshSupplier
'    Call refreshPekerjaan
'    Call refreshPromo
    Call refreshDealer
End Sub

 

Private Sub optCreator_Click(Index As Integer)
    Call refreshSupplier
End Sub

Private Sub TblKeluar_Click()
    onTimer = True
    Unload Me
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean
Dim txtTemp As String
Dim Txt As String
Dim dealer As String
    
    boleh = False
    txtTemp = ""
    Txt = ""
    ProgressBar1.Visible = False
    prosen.Visible = False
    
    If NulSID(cbPekerjaan.value) <> "" Then
        Txt = Txt & " and a.[KDSUBLET] = '" & NulSID(cbPekerjaan.value) & "'"
    End If
    
'    If NulSID(cbPromo.value) <> "" Then
'        Txt = Txt & " and e.[kdpromo] = '" & NulSID(cbPromo.value) & "'"
'    End If
    
    If NulSID(cbSupplier.value) <> "" Then
        Txt = Txt & " and b.KDSUPP = '" & NulSID(cbSupplier.value) & "'"
    End If
    
    If tglCreate1.value <> "" And tglCreate2.value <> "" Then
        Txt = Txt & " and (a.tglentry BETWEEN '" & Format(tglCreate1.value, "yyyy-mm-dd") & " 00:00' AND '" & Format(tglCreate2.value, "yyyy-mm-dd") & " 23:59')"
    End If
    
    If NulSID(tglPeriode(0).value) <> "" And NulSID(tglPeriode(1).value) <> "" Then
        Txt = Txt & " and (a.awalper = akhirper or ('" & Format(tglPeriode(0).value, "yyyy-mm-dd") & "' BETWEEN a.awalper and a.akhirper) or ('" & Format(tglPeriode(1).value, "yyyy-mm-dd") & "' BETWEEN a.awalper and a.akhirper)) "
    End If
    
    If optCreator(0).value Then
        Txt = Txt & " and SUBSTRING(a.KDSUBLET,1,4) = 'NR00'"
    ElseIf optCreator(1).value Then
        Txt = Txt & " and SUBSTRING(a.KDSUBLET,1,4) <> 'NR00'"
    End If
    
    If NulSID(cbDealer.value) <> "" Then
        Txt = Txt & " and (a.ALLCABANG = 1 or c.FS_KD_DEALER = '" & NulSID(cbDealer.value) & "')"
    ElseIf gUstat = 0 Then
        Txt = Txt & " and (a.ALLCABANG = 1 or c.FS_KD_DEALER = '" & Log.KodeDealer & "')"
    End If
    
    
    If NulSID(cbStatus.text) = NulSID("All") Or NulSID(cbStatus.text) = "" Then
        Txt = Txt
    ElseIf NulSID(cbStatus.text) = NulSID("On") Then
        Txt = Txt & " and a.FS_LOG_DEL = '' "
    ElseIf NulSID(cbStatus.text) = NulSID("Off") Then
        Txt = Txt & " and a.FS_LOG_DEL <> '' "
    End If
'    If opCetak(0).value Then
'       ' Call IsiTempRepo(txtTemp, vNamaTabel, Txt)
'        dealer = ",b.FS_NM_DEALER"
'    ElseIf opCetak(1).value Then
'       ' Call harianMech
'        dealer = ""
'    End If
    
    If gUstat = 1 Then
    
        sSQL = "select DISTINCT a.KDSUBLET,a.NAMASUBLET,a.KODESATUAN,a.KETERANGAN,a.HARGABELI,a.HARGAJUAL,b.KETERANGAN,"
        sSQL = sSQL & " case a.allcabang when '1' then 'All Cabang' when '0' then 'Cabang Tertentu' end as cabang,"
        sSQL = sSQL & " case a.ALLMODEL when '1' then 'All Model' when '0' then 'Model Tertentu' end as model,"
        sSQL = sSQL & " a.AWALPER , a.AKHIRPER, a.TGLENTRY"
        sSQL = sSQL & " from MSTSUBLET a with (nolock)"
        sSQL = sSQL & " inner join MSTSUPPLIER b with (nolock) on a.KDSUPP = b.KDSUPP"
        sSQL = sSQL & " LEFT JOIN MSTSUBLETDEALER c with (nolock) on c.KDSUBLET = a.KDSUBLET"
        sSQL = sSQL & " where b.FS_LOG_DEL = '' " & Txt
        ''sSQL = sSQL & " where 1 = 1 " & Txt
        'sSQL = sSQL & " group by a.NAMABARANG " & dealer
        
        Set rsLaporan = objO.FetchDataHO(sSQL)
    Else
'        sSQL = "select a.NAMABARANG,sum(a.QTY) as qty,sum(a.HARGA*a.qty) hargajual,sum(a.HARGABELI*a.qty) hargabeli,sum(a.HARGA*a.qty)-sum(a.HARGABELI*a.qty) as total " & IIf(dealer = "", ",'ALL DEALER' as FS_NM_DEALER", dealer)
'        sSQL = sSQL & " from [" & dbReplikasiHO & "]..HIST_SUBLETnew a with (nolock)"
'        sSQL = sSQL & " left join [" & dbReplikasiHO & "]..T_INVOICEnew b with (nolock) on a.no_wo = b.no_wo and a.noinv = b.noinv and a.CABANG = b.cabang"
'        sSQL = sSQL & " left join [" & dbReplikasiHO & "]..MSTSUBLET c with (nolock) on c.KDSUBLET = a.[KODEITEM]"
'        sSQL = sSQL & " left join [" & dbReplikasiHO & "]..MSTSUPPLIER d with (nolock) on d.KDSUPP = c.KDSUPP"
'        sSQL = sSQL & " left join [" & dbReplikasiHO & "]..MSTSUBLETPROMO e with (nolock) on e.kdpromo = a.kdpromo"
'        sSQL = sSQL & " where 1 = 1 " & Txt
'        sSQL = sSQL & " group by a.NAMABARANG " & dealer

        sSQL = "select DISTINCT a.KDSUBLET,a.NAMASUBLET,a.KODESATUAN,a.KETERANGAN,a.HARGABELI,a.HARGAJUAL,b.KETERANGAN,"
        sSQL = sSQL & " case a.allcabang when '1' then 'All Cabang' when '0' then 'Cabang Tertentu' end as cabang,"
        sSQL = sSQL & " case a.ALLMODEL when '1' then 'All Model' when '0' then 'Model Tertentu' end as model,"
        sSQL = sSQL & " a.AWALPER , a.AKHIRPER, a.TGLENTRY"
        sSQL = sSQL & " from [" & dbReplikasiHO & "]..MSTSUBLET a with (nolock)"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..MSTSUPPLIER b with (nolock) on a.KDSUPP = b.KDSUPP"
        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER c with (nolock) on c.KDSUBLET = a.KDSUBLET"
        sSQL = sSQL & " where b.FS_LOG_DEL = '' " & Txt
        
        Set rsLaporan = objO.FetchData(sSQL)
    End If
    If Not rsLaporan.EOF Then
        Dim rsRepo As ADODB.Recordset
        Screen.MousePointer = vbHourglass
            
        Set Report = New rptRepoSentralisasiMstSublet
        Report.database.SetDataSource rsLaporan
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
        Call DisableExport
        Call setLokasiTombol
    Else
        MsgBox "Tidak ada data yang ditampilkan.", vbInformation
    End If
End Sub
 
Private Sub createRSLaporan()
   On Error GoTo createRSLaporan_Error
    If rsLaporan.State Then rsLaporan.Close
    Set rsLaporan = Nothing
    With rsLaporan
        .Fields.Append "NO_POLISI", adVarChar, 12
        .Fields.Append "NO_WO", adVarChar, 8
        .Fields.Append "NOINV", adVarChar, 10
        .Fields.Append "CUSTOMER", adVarChar, 160
        .Fields.Append "NO_CHASSIS", adVarChar, 25
        .Fields.Append "FMDL", adVarChar, 50
        .Fields.Append "FS_NM_DEALER", adVarChar, 30
        .Fields.Append "FRONT", adVarChar, 5
        .Fields.Append "GRP1", adVarChar, 100
        .Fields.Append "RTE1", adDouble
        
        .Fields.Append "ACTRATE1", adDouble
        .Fields.Append "DISC1", adDouble
        .Fields.Append "discPartValid", adDouble
        .Fields.Append "discSublet", adDouble
        .Fields.Append "LABOR", adDouble
        .Fields.Append "part", adDouble
        .Fields.Append "SBLT", adDouble
        .Fields.Append "Total", adDouble
        .Fields.Append "KODETIPE", adDouble
        .Open
    End With
   On Error GoTo 0
   Exit Sub

createRSLaporan_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure createRSLaporan of Form frmRepoSentralMstSublet")
End Sub

Private Sub refreshPromo()
    sSQL = "SELECT DISTINCT A.KDPROMO,A.NAMAPROMO,A.KETERANGAN,A.HARGAJUAL,A.KDSUBLET "
    sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLETPROMO A WITH (NOLOCK)"
'    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLET D WITH (NOLOCK) ON A.KDSUBLET = D.KDSUBLET"
'    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMOMDL B WITH (NOLOCK) ON A.KDPROMO = B.KDPROMO"
'    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMODEALER C WITH (NOLOCK) ON A.KDPROMO = C.KDPROMO"
    sSQL = sSQL & " WHERE  A.FS_LOG_DEL = '' "

    Adodc1.ConnectionString = SQLCONBengkel
    Adodc1.RecordSource = sSQL
    cbPromo.DataSourceList = Adodc1
    cbPromo.DataFieldList = "KDPROMO"
    
    cbPromo.Columns(0).Width = 2000
    cbPromo.Columns(1).Width = 3500
    cbPromo.Columns(0).Caption = "Kode Promo"
    cbPromo.Columns(1).Caption = "Nama Promo"
    cbPromo.Columns(2).Visible = False
    cbPromo.Columns(3).Visible = False
    cbPromo.Columns(4).Visible = False
    cbPromo.text = ""
End Sub


Private Sub refreshPekerjaan()
    If gUstat = 1 Then
        sSQL = "SELECT DISTINCT D.KDSUBLET,D.NAMASUBLET,D.KETERANGAN,D.HARGAJUAL"
        sSQL = sSQL & " FROM MSTSUBLET D "
        sSQL = sSQL & " LEFT JOIN MSTSUBLETMDL B WITH (NOLOCK) ON D.KDSUBLET = B.KDSUBLET"
        sSQL = sSQL & " LEFT JOIN MSTSUBLETDEALER C WITH (NOLOCK) ON D.KDSUBLET = C.KDSUBLET"
        sSQL = sSQL & " WHERE D.FS_LOG_DEL = '' " & IIf(cbSupplier.value <> "", " and D.KDSUPP = '" & NulSID(cbSupplier.value) & "'", "")
        sSQL = sSQL & " and (D.ALLCABANG = '1' " & IIf(NulSID(cbDealer.value) <> "", " AND C.FS_KD_DEALER = '" & NulSID(cbDealer.value) & "'", "") & ")"
    Else
        sSQL = "SELECT DISTINCT D.KDSUBLET,D.NAMASUBLET,D.KETERANGAN,D.HARGAJUAL"
        sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLET D "
        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL B WITH (NOLOCK) ON D.KDSUBLET = B.KDSUBLET"
        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER C WITH (NOLOCK) ON D.KDSUBLET = C.KDSUBLET"
        sSQL = sSQL & " WHERE D.FS_LOG_DEL = '' " & IIf(cbSupplier.value <> "", " and D.KDSUPP = '" & NulSID(cbSupplier.value) & "'", "")
        sSQL = sSQL & " and (D.ALLCABANG = '1'  OR C.FS_KD_DEALER = '" & NulSID(Log.KodeDealer) & "')"
    End If
    Adodc2.ConnectionString = SQLCONBengkel
    Adodc2.RecordSource = sSQL
    Adodc2.Refresh
    cbPekerjaan.DataSourceList = Adodc2
    cbPekerjaan.DataFieldList = "KDSUBLET"
    
    cbPekerjaan.Columns(0).Width = 1700
    cbPekerjaan.Columns(1).Width = 4000
    cbPekerjaan.Columns(0).Caption = "Kode Sublet"
    cbPekerjaan.Columns(1).Caption = "Nama Sublet"
    cbPekerjaan.Columns(2).Visible = False
    cbPekerjaan.Columns(3).Visible = False
    cbPekerjaan.Columns(4).Visible = False
    cbPekerjaan.text = ""
End Sub

Private Sub refreshDealer()
    If gUstat = 1 Then
        sSQL = "select FS_KD_DEALER,FS_NM_DEALER,DBASECBG from T_DEALER where FS_LOG_DEL = '' and INTERNAL = 1 and FS_KD_DEALER <> 'NR 00' ORDER BY FS_KD_DEALER "
    Else
        sSQL = "select FS_KD_DEALER,FS_NM_DEALER,DBASECBG from [" & dbReplikasiHO & "]..T_DEALER where FS_LOG_DEL = '' and INTERNAL = 1 and FS_KD_DEALER <> 'NR 00' ORDER BY FS_KD_DEALER "
    
    End If
    Adodc3.ConnectionString = SQLCONBengkel
    Adodc3.RecordSource = sSQL
    cbDealer.DataSourceList = Adodc3
    cbDealer.DataFieldList = "FS_KD_DEALER"
    cbDealer.DataFieldToDisplay = "FS_NM_DEALER"
    
    cbDealer.Columns(0).Width = 1700
    cbDealer.Columns(1).Width = 4000
    cbDealer.Columns(0).Caption = "Kode Dealer"
    cbDealer.Columns(1).Caption = "Nama Dealer"
    cbDealer.Columns(2).Visible = False
    cbDealer.text = ""
End Sub


Private Sub refreshSupplier()
    If optCreator(2).value Then
        If gUstat = 1 Then
            sSQL = "select KDSUPP,SUPPLIER,ALAMAT1 from MSTSUPPLIER with (nolock) where FS_LOG_DEL = ''"
        Else
            sSQL = "select KDSUPP,SUPPLIER,ALAMAT1 from [" & dbReplikasiHO & "]..MSTSUPPLIER with (nolock) where FS_LOG_DEL = '' AND (SUBSTRING(KDSUPP,1,4) = 'NR00' OR CABANG = '" & NulSID(gDb) & "')"
        End If
    ElseIf optCreator(0).value Then
        If gUstat = 1 Then
            sSQL = "select KDSUPP,SUPPLIER,ALAMAT1 from MSTSUPPLIER with (nolock) where FS_LOG_DEL = '' and SUBSTRING(KDSUPP,1,4) = 'NR00' "
        Else
            sSQL = "select KDSUPP,SUPPLIER,ALAMAT1 from [" & dbReplikasiHO & "]..MSTSUPPLIER with (nolock) where FS_LOG_DEL = '' and SUBSTRING(KDSUPP,1,4) = 'NR00' "
        End If
        
    ElseIf optCreator(1).value Then
        If gUstat = 1 Then
            sSQL = "select KDSUPP,SUPPLIER,ALAMAT1 from MSTSUPPLIER with (nolock) where FS_LOG_DEL = '' and SUBSTRING(KDSUPP,1,4) <> 'NR00' "
        Else
            sSQL = "select KDSUPP,SUPPLIER,ALAMAT1 from [" & dbReplikasiHO & "]..MSTSUPPLIER with (nolock) where FS_LOG_DEL = '' and SUBSTRING(KDSUPP,1,4) <> 'NR00' AND CABANG = '" & NulSID(gDb) & "' "
        End If
        
    End If
    Adodc4.ConnectionString = SQLCONBengkel
    Adodc4.RecordSource = sSQL
    Adodc4.Refresh
    cbSupplier.DataSourceList = Adodc4
    cbSupplier.DataFieldList = "KDSUPP"
    cbSupplier.DataFieldToDisplay = "KDSUPP"
    
    cbSupplier.Columns(0).Width = 1700
    cbSupplier.Columns(1).Width = 4000
    cbSupplier.Columns(0).Caption = "Kode Supplier"
    cbSupplier.Columns(1).Caption = "Nama Supplier"
    cbSupplier.Columns(2).Caption = "Alamat Supplier"
    cbSupplier.Refresh
    cbSupplier.text = ""
End Sub
