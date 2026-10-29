VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoSOdalamProses 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan SO dalam Proses"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   13080
   ShowInTaskbar   =   0   'False
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
      Left            =   705
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5955
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
      Left            =   1815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5955
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
      Left            =   3300
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5955
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
      Left            =   8745
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5955
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5625
      TabIndex        =   52
      Top             =   5985
      Visible         =   0   'False
      Width           =   3015
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   270
      Left            =   4575
      TabIndex        =   57
      Top             =   5955
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
      MICON           =   "frmRepoSOdalamProses.frx":0000
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
      Left            =   600
      TabIndex        =   14
      Top             =   5640
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   480
      TabIndex        =   18
      Top             =   480
      Width           =   12375
      Begin VB.OptionButton opVehicle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1 Model"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   6
         Top             =   2280
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton opVehicle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Beberapa Model"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   7
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox chRTJ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Return Job"
         Height          =   255
         Left            =   10500
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CheckBox chTHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THS"
         Height          =   255
         Left            =   10500
         TabIndex        =   19
         Top             =   2160
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   2
         Top             =   840
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
         TabIndex        =   21
         Top             =   840
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
         Top             =   1200
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   5400
         TabIndex        =   22
         Top             =   1200
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
         Index           =   3
         Left            =   1530
         TabIndex        =   4
         Top             =   1560
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
         TabIndex        =   23
         Top             =   1560
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
         TabIndex        =   24
         Top             =   840
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
         Index           =   4
         Left            =   2520
         TabIndex        =   25
         Top             =   1560
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
         Index           =   5
         Left            =   1530
         TabIndex        =   5
         Top             =   1920
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
         Left            =   2400
         TabIndex        =   26
         Top             =   1920
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
         Index           =   7
         Left            =   1530
         TabIndex        =   8
         Top             =   2640
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   3480
         TabIndex        =   27
         Top             =   2640
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
         Left            =   2520
         TabIndex        =   28
         Top             =   1920
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
         Index           =   8
         Left            =   1530
         TabIndex        =   29
         Top             =   3000
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   7290
         TabIndex        =   9
         Top             =   840
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   8160
         TabIndex        =   30
         Top             =   847
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
         Index           =   10
         Left            =   8280
         TabIndex        =   31
         Top             =   840
         Width           =   3675
         _ExtentX        =   6482
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
         Index           =   11
         Left            =   7290
         TabIndex        =   10
         Top             =   1200
         Width           =   1755
         _ExtentX        =   3096
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
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Left            =   9135
         TabIndex        =   32
         Top             =   1207
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
         Index           =   12
         Left            =   7290
         TabIndex        =   11
         Top             =   1920
         Width           =   1755
         _ExtentX        =   3096
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
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   270
         Left            =   9135
         TabIndex        =   33
         Top             =   1920
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
         Index           =   13
         Left            =   7290
         TabIndex        =   12
         Top             =   2280
         Width           =   1755
         _ExtentX        =   3096
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
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   9135
         TabIndex        =   34
         Top             =   2280
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
         Index           =   14
         Left            =   2400
         TabIndex        =   35
         Top             =   2640
         Visible         =   0   'False
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   15
         Left            =   7290
         TabIndex        =   13
         Top             =   2640
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin TDBDate6Ctl.TDBDate Tgl1 
         Height          =   330
         Left            =   1560
         TabIndex        =   0
         Top             =   172
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoSOdalamProses.frx":001C
         Caption         =   "frmRepoSOdalamProses.frx":0148
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoSOdalamProses.frx":01B4
         Keys            =   "frmRepoSOdalamProses.frx":01D2
         Spin            =   "frmRepoSOdalamProses.frx":0230
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
         TabIndex        =   1
         Top             =   172
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmRepoSOdalamProses.frx":0258
         Caption         =   "frmRepoSOdalamProses.frx":0384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoSOdalamProses.frx":03F0
         Keys            =   "frmRepoSOdalamProses.frx":040E
         Spin            =   "frmRepoSOdalamProses.frx":046C
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   24
         Left            =   10455
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox MaskEdBox11 
         Height          =   270
         Left            =   11880
         TabIndex        =   49
         Top             =   1200
         Visible         =   0   'False
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
         Index           =   25
         Left            =   7290
         TabIndex        =   50
         Top             =   1560
         Visible         =   0   'False
         Width           =   4700
         _ExtentX        =   8281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Program"
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
         Left            =   9330
         TabIndex        =   51
         Top             =   1245
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Job Order"
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
         Left            =   240
         TabIndex        =   47
         Top             =   885
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Order"
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
         TabIndex        =   46
         Top             =   1245
         Width           =   720
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
         TabIndex        =   45
         Top             =   1605
         Width           =   660
      End
      Begin VB.Label Label3 
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
         Left            =   1245
         TabIndex        =   44
         Top             =   1965
         Width           =   195
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Kendaraan"
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
         Left            =   195
         TabIndex        =   43
         Top             =   2325
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Penjual"
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
         Left            =   6165
         TabIndex        =   42
         Top             =   885
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Service"
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
         Left            =   6330
         TabIndex        =   41
         Top             =   1245
         Width           =   870
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Order"
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
         Left            =   6435
         TabIndex        =   40
         Top             =   1965
         Width           =   765
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proses"
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
         Left            =   6720
         TabIndex        =   39
         Top             =   2325
         Width           =   480
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   240
         Width           =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   12360
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Th Kend."
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
         Left            =   6555
         TabIndex        =   36
         Top             =   2685
         Width           =   645
      End
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4200
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   1890
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8655
      Left            =   0
      Picture         =   "frmRepoSOdalamProses.frx":0494
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoSOdalamProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 28 Juni 2010
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsTemp      As ADODB.Recordset
Dim Report      As Object
Dim vNamaTabel  As String
Dim isAdaData   As Boolean

Private Sub hapusForm()
    For i = 0 To 15
        Select Case i
            Case 15
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    Tgl1 = isnow
    Tgl2 = isnow
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset

    createTemp = False
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelRepo(vNamaTabel)
        createTemp = True
    End If
End Function

Private Sub IsiTempRepo(Ket As String, vNamaTabel As String, txt As String)
Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
Dim hlp5 As String
    
    'DELIV <> 'BTL' AND
    If Nul(Trim(Tgl1)) = "" Then
        Set rsTemp = objO.FetchData("SELECT *, A.MECH AS MECH, A.CUSTOMER AS CUSTOMER, C.FS_KD_DEALER AS KDDLRJUAL, C.FS_NM_DEALER AS NMDLRJUAL,CASE D.TUNGGUSVC WHEN 0 THEN 'DITUNGGU' ELSE 'DITINGGAL' END AS TUNGGUSVC FROM JOBDTA A WITH(NOLOCK) INNER JOIN CUSTDAT C WITH(NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN STALL B WITH(NOLOCK) ON REPLACE(SUBSTRING(B.KETERANGAN, CHARINDEX('|',B.KETERANGAN)+1,8),'|','') = A.NO_WO INNER JOIN JOBDTL D WITH(NOLOCK) ON A.NO_WO = D.NO_WO WHERE [DATE] <= '" & Format(Tgl2, "MM/DD/YYYY") & "' AND FS_LOG_DEL = '' AND NOINV = '' AND DELIV <> 'BTL' AND RPR_COND <> 'K' " & Ket & " ORDER BY A.NO_WO")
    Else
        Set rsTemp = objO.FetchData("SELECT *, A.MECH AS MECH, A.CUSTOMER AS CUSTOMER, C.FS_KD_DEALER AS KDDLRJUAL, C.FS_NM_DEALER AS NMDLRJUAL,CASE D.TUNGGUSVC WHEN 0 THEN 'DITUNGGU' ELSE 'DITINGGAL' END AS TUNGGUSVC FROM JOBDTA A WITH(NOLOCK) INNER JOIN CUSTDAT C WITH(NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN STALL B WITH(NOLOCK) ON REPLACE(SUBSTRING(B.KETERANGAN, CHARINDEX('|',B.KETERANGAN)+1,8),'|','') = A.NO_WO INNER JOIN JOBDTL D WITH(NOLOCK) ON A.NO_WO = D.NO_WO WHERE [DATE] BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND FS_LOG_DEL = '' AND NOINV = '' AND DELIV <> 'BTL' AND RPR_COND <> 'K' " & Ket & " ORDER BY A.NO_WO")
    End If
    If Not rsTemp.EOF Then
        Call BuatTabelRepo(vNamaTabel)
        Do While Not rsTemp.EOF
            ' DARI JOBDTA
            For i = 1 To 8
                hlp = "REP" & i
                If rsTemp(hlp) <> "" Then
                    hlp2 = "KLPKERJA" & i
                    hlp3 = "ACTRATE" & i
                    hlp4 = "MECH" '"MEKANIK" & i
                    hlp5 = "RTE" & i
                    If createTemp(vNamaTabel) Then
                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("NO_CHASSIS") & "', '" & rsTemp("NO_POLISI") & "', '" & rsTemp("FS_KD_PLG") & "', '" & rsTemp("FS_KD_WIL") & "', " _
                                     & "'" & rsTemp("CUSTOMER") & "', '" & rsTemp("MDL") & "', '" & rsTemp("FMDL") & "', '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "', '" & rsTemp("NO_WO") & "','" & rsTemp("NOINV") & "', '" & rsTemp("FRONT") & "', '" & rsTemp(hlp2) & "', " _
                                     & "'" & rsTemp(hlp) & "', '" & Trim(rsTemp(hlp4)) & "','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 1, Replace(CDbl(rsTemp(hlp5)) * -1, ",", "."), Replace(CDbl(rsTemp(hlp5)), ",", ".")) & "', '0', '0', '" & Trim(rsTemp("TPSVC")) & "', '" & rsTemp("THS") & "', '" & rsTemp("FN_RETURN") & "', '" & rsTemp("FLAG") & "', " _
                                     & "'" & rsTemp("NOFAK") & "', '" & IIf(IsNull(rsTemp("NOSTALL")), "", rsTemp("NOSTALL")) & "', '" & rsTemp("KDDLRJUAL") & "', '" & rsTemp("NMDLRJUAL") & "','" & rsTemp("TIME") & "','00:00')")
                    End If
                End If
            Next
            
            ' DARI ADDITIONAL JOB
            Dim rsTSub As ADODB.Recordset
            Set rsTSub = objO.FetchData("SELECT * FROM JOBADD2 WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsTemp("NO_WO")) & "' AND NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "'")
            If Not rsTSub.EOF Then
                Do While Not rsTSub.EOF
                    If createTemp(vNamaTabel) Then
                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("NO_CHASSIS") & "', '" & rsTemp("NO_POLISI") & "', '" & rsTemp("FS_KD_PLG") & "', '" & rsTemp("FS_KD_WIL") & "', " _
                                     & "'" & rsTemp("CUSTOMER") & "', '" & rsTemp("MDL") & "', '" & rsTemp("FMDL") & "', '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "', '" & rsTemp("NO_WO") & "','" & rsTemp("NOINV") & "', '" & rsTemp("FRONT") & "', '" & rsTSub("KLPKERJA") & "', " _
                                     & "'" & rsTSub("REP") & "', '" & Trim(rsTemp(hlp4)) & "','" & Replace(CDbl(rsTSub("ACTRTE")), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 1, Replace((CDbl(rsTSub("RTE")) * -1), ",", "."), Replace(CDbl(rsTSub("RTE")), ",", ".")) & "', '0', '0', '" & Trim(rsTemp("TPSVC")) & "', '" & rsTemp("THS") & "', '" & rsTemp("FN_RETURN") & "', '" & rsTemp("FLAG") & "', " _
                                     & "'" & rsTemp("NOFAK") & "', '" & IIf(IsNull(rsTemp("NOSTALL")), "", rsTemp("NOSTALL")) & "', '" & rsTemp("KDDLRJUAL") & "', '" & rsTemp("NMDLRJUAL") & "','" & rsTemp("TIME") & "','00:00')")
                    End If
                    rsTSub.MoveNext
                Loop
            End If
            
            ' DARI SUBLET
            Set rsTSub = objO.FetchData("SELECT * FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsTemp("NO_WO")) & "' AND NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "'")
            If Not rsTSub.EOF Then
                Do While Not rsTSub.EOF
                    If createTemp(vNamaTabel) Then
                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("NO_CHASSIS") & "', '" & rsTemp("NO_POLISI") & "', '" & rsTemp("FS_KD_PLG") & "', '" & rsTemp("FS_KD_WIL") & "', " _
                                     & "'" & rsTemp("CUSTOMER") & "', '" & rsTemp("MDL") & "', '" & rsTemp("FMDL") & "', '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "', '" & rsTemp("NO_WO") & "','" & rsTemp("NOINV") & "', '" & rsTemp("FRONT") & "', 'SUB', " _
                                     & "'" & rsTSub("NAMABARANG") & "', '','0', '0', '0', '0', '" & Trim(rsTemp("TPSVC")) & "', '" & rsTemp("THS") & "', '" & rsTemp("FN_RETURN") & "', '" & rsTemp("FLAG") & "', " _
                                     & "'" & rsTemp("NOFAK") & "', '" & IIf(IsNull(rsTemp("NOSTALL")), "", rsTemp("NOSTALL")) & "', '" & rsTemp("KDDLRJUAL") & "', '" & rsTemp("NMDLRJUAL") & "','" & rsTemp("TIME") & "','00:00')")
                    End If
                    rsTSub.MoveNext
                Loop
            End If
            rsTemp.MoveNext
        Loop
        isAdaData = True
        Call Cetak(vNamaTabel, txt)
    Else
        isAdaData = False
        MsgBox "Tidak ada data !!", vbInformation + vbOKOnly
        Exit Sub
    End If
End Sub

Private Sub Cetak(vNamaTabel As String, txt As String)
Dim rsRepo As ADODB.Recordset
Screen.MousePointer = vbHourglass
    
    Set Report = New RptRepoSOdalamProses
    If txt <> "" Then
        txt = " WHERE NO_WO IN (SELECT DISTINCT NO_WO FROM [TEMPDB]..[" & vNamaTabel & "] WHERE " & txt & ")"
    End If
    Set rsRepo = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "]" & txt & " ORDER BY NO_WO, KLPKERJA")
    Report.database.SetDataSource rsRepo
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
    Call hapusForm
    vNamaTabel = "SOdalamProses" & DriveSerialNumber("c:")
End Sub

Private Sub opVehicle_Click(Index As Integer)
    If opVehicle(0) Then
        Text(7).Visible = True
        Text(8).Visible = True
        MaskEdBox4.Visible = True
        Text(14).Visible = False
    Else
        Text(7).Visible = False
        Text(8).Visible = False
        MaskEdBox4.Visible = False
        Text(14).Visible = True
        Text(14).Left = Text(7).Left
        Text(14).SelLength = Len(Text(14))
        Text(14).SetFocus
    End If
End Sub

Private Sub TblKeluar_Click()
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call hapusForm
End Sub

Private Sub tblSimpan_Click()
      Dim boleh As Boolean
      Dim txtTemp As String
      Dim txt As String
1         isAdaData = False
2         boleh = True
3         txtTemp = ""
4         txt = ""
5         If Nul(Trim(Tgl2)) = "" Then
6             boleh = False
7             MsgBox "Tanggal Akhir tidak boleh Kosong", vbInformation + vbOKOnly
8         End If
9         If Text(0) <> "" Then
10            If txt <> "" Then
11                txt = txt & "AND "
12            End If
13            txt = txt & "KLPKERJA = '" & Text(0) & "' "
14        End If
15        If Text(2) <> "" Then
16            If txt <> "" Then
17                txt = txt & "AND "
18            End If
19            txt = txt & "REPS LIKE '%" & Text(2) & "%' "
20        End If
21        If Text(3) <> "" Then
22            boleh = True
23            If txt <> "" Then
24                txt = txt & "AND "
25            End If
26            txt = txt & "MECHS = '" & Text(3) & "' "
27        End If
28        If Text(5) <> "" Then
29            txtTemp = txtTemp & "AND FRONT = '" & Text(5) & "' "
30        End If
31        If Text(7) <> "" Or Text(14) <> "" Then
32            If opVehicle(0) Then
33                txtTemp = txtTemp & "AND A.MDL = '" & Trim(Text(7)) & "' "
34            Else
35                txtTemp = txtTemp & "AND A.FMDL LIKE '%" & Trim(Text(14)) & "%' "
36            End If
37        End If
38        If Text(9) <> "" Then
39            txtTemp = txtTemp & "AND C.FS_KD_DEALER = '" & Text(9) & "' "
40        End If
41        If Text(11) <> "" Then
      '        txtTemp = txtTemp & "AND TPSVC = '" & IIf(Text(11) = "REGULAR", "", IIf(Text(11) = "WARRANTY", "W", "P")) & "' "
      '        If Trim(Text(11)) = "PROGRAM" Then
      '            If Trim(Text(24)) <> "" Then
      '                txtTemp = txtTemp & "AND KODETIPE = '" & Trim(Text(24)) & "' "
      '            End If
      '        End If
              
42            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
43                txtTemp = txtTemp & "AND TPSVC = '" & IIf(Text(11) = "WARRANTY", "W", "") & "' "
44            Else
45                txtTemp = txtTemp & "AND TPSVC = '" & IIf(Text(11) = "REGULAR", "", IIf(Text(11) = "WARRANTY", "W", "P")) & "' "
46            End If
47            If Trim(Text(11)) = "PROGRAM" Then
48                If Trim(Text(24)) <> "" Then
49                    If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
50                        txtTemp = txtTemp & "AND (KODETIPE = '" & Trim(Text(24)) & "' OR A.no_wo in (select no_wo from T_invoiceJobDetail where FS_LOG_DEL = '' and no_wo = A.no_wo and id_program = '" & Trim(Text(24)) & "' )) "
51                    Else
52                        txtTemp = txtTemp & "AND KODETIPE = '" & Trim(Text(24)) & "' "
53                    End If
54                End If
55            End If
56        End If
              
57        If Text(12) <> "" Then
58            If Text(12) = "BOOKING" Then
59                txtTemp = txtTemp & "AND NOFAK <> ''"
60            ElseIf Text(12) = "NON BOOKING" Then
61                txtTemp = txtTemp & "AND NOFAK = ''"
62            End If
63        End If
64        If Text(13) <> "" Then
65            If Text(13) = "PROSES" Then
66                txtTemp = txtTemp & "AND FLAG = 'S'" '"AND RPR_COND <> 'K' AND RPR_COND <> '' AND RPR_COND <> 'S' AND RPR_COND <> 'T'"
67            ElseIf Text(13) = "WAITING" Then
68                txtTemp = txtTemp & "AND FLAG = 'N'"
69            ElseIf Text(13) = "FINISH" Then
70                txtTemp = txtTemp & "AND FLAG = 'F'"
71            ElseIf Text(13) = "CLOSING" Then
72                txtTemp = txtTemp & "AND FLAG = 'K'"
73            ElseIf Text(13) = "PAUSED" Then
74                txtTemp = txtTemp & "AND FLAG = 'P'"
75            End If
76        End If
77        If Text(15) > 0 Then
78            txtTemp = txtTemp & "AND C.FS_TH_MOBIL = '" & CDbl(Text(15)) & "'"
79        End If
80        If chRTJ = Checked Then
81            txtTemp = txtTemp & "AND FN_RETURN = '1'"
82        End If
83        If chTHS = Checked Then
84            txtTemp = txtTemp & "AND THS = '1'"
85        End If
          
86        If boleh Then
87            Call IsiTempRepo(txtTemp, vNamaTabel, txt)
88            If isAdaData Then
89                Call DisableExport
90                Call setLokasiTombol
91            End If
92        End If
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
                gSearchOption = "JENIS_KERJA"
                frmCari.show (vbModal)
            Case 2
                If Text(0) <> "" Then
                    If Trim(Text(0)) <> "SUB" Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "JOB_ORDER"
                        gBantuan = Text(0)
                        frmCari.show (vbModal)
                    Else
                        gMainKeyCari = Text(Index)
                        gSearchOption = "SUBLET"
                        frmCari.show (vbModal)
                    End If
                Else
                    MsgBox "Group Job Order Harus diisi", vbInformation + vbOKOnly
                    Text(0).SetFocus
                    Exit Sub
                End If
            Case 3
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ALL"
                gBantuan4 = ""
                frmCari.show (vbModal)
            Case 5
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.show (vbModal)
            Case 7
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
            Case 9
                gMainKeyCari = Text(Index)
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCari.show (vbModal)
            Case 11
                gMainKeyCari = Text(Index)
                gSearchOption = "TPSVC"
                frmCari.show (vbModal)
            Case 12
                gMainKeyCari = Text(Index)
                gSearchOption = "TP_ORD"
                frmCari.show (vbModal)
            Case 13
                gMainKeyCari = Text(Index)
                gSearchOption = "PROSES"
                frmCari.show (vbModal)
            Case 24
                ' PROGRAM
                gMainKeyCari = Text(Index)
                gSearchOption = "PROGRAM_ALL" '"PROGRAM"
                gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
                frmCari.show (vbModal)
                 
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsHlp As ADODB.Recordset
1             Select Case Index
                  Case 0
2                     If Text(Index) <> "" Then
3                         Set oRs = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE JOB = '" & Text(Index) & "' and FS_LOG_DEL = ''")
4                         If oRs.EOF Then
5                             If Trim(Text(Index)) <> "SUB" Then
6                                 gMainKeyCari = Text(Index)
7                                 gSearchOption = "JENIS_KERJA"
8                                 frmCari.show (vbModal)
9                             Else
10                                Text(Index + 1) = "SUBLET"
11                            End If
12                        Else
13                            Text(Index + 1) = oRs("REMARK")
14                        End If
15                        oRs.Close
16                    Else
17                        Text(Index + 1) = ""
18                    End If
19                Case 2
20                    If Text(0) <> "" Then
21                        If Text(Index) <> "" Then
22                            Set oRs = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE JOB = '" & Text(0) & "' AND FS_ITEM = '" & Text(Index) & "' and FS_LOG_DEL = ''")
23                            If oRs.EOF Then
24                                If Trim(Text(0)) <> "SUB" Then
25                                    gMainKeyCari = Text(Index)
26                                    gSearchOption = "JOB_ORDER"
27                                    gBantuan = Text(0)
28                                    frmCari.show (vbModal)
29                                Else
                                       
30                                    Set rsHlp = objO.FetchData("SELECT DISTINCT NAMABARANG FROM JOBSUBLET WITH(NOLOCK) WHERE NAMABARANG = '" & Trim(Text(Index)) & "'")
31                                    If rsHlp.EOF Then
32                                        gMainKeyCari = Text(Index)
33                                        gSearchOption = "SUBLET"
34                                        frmCari.show (vbModal)
35                                    End If
36                                    rsHlp.Close
37                                End If
38                            End If
39                            oRs.Close
40                        End If
41                    Else
42                        MsgBox "Group Job Order Harus diisi", vbInformation + vbOKOnly
43                        Text(0).SetFocus
44                        Exit Sub
45                    End If
46                Case 3
47                    If Text(Index) <> "" Then
48                        Set oRs = objO.FetchData("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Text(Index) & "'")
49                        If oRs.EOF Then
50                            gMainKeyCari = Text(Index)
51                            gSearchOption = "MEKANIK"
52                            gBantuan = "ALL"
53                            gBantuan4 = ""
54                            frmCari.show (vbModal)
55                        Else
56                            Text(Index + 1) = oRs("NAME")
57                        End If
58                        oRs.Close
59                    Else
60                        Text(Index + 1) = ""
61                    End If
62                Case 5
63                    If Text(Index) <> "" Then
64                        Set oRs = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & Text(Index) & "'")
65                        If oRs.EOF Then
66                            gMainKeyCari = Text(Index)
67                            gSearchOption = "SA"
68                            frmCari.show (vbModal)
69                        Else
70                            Text(Index + 1) = oRs("FS_NM_PETUGAS")
71                        End If
72                        oRs.Close
73                    Else
74                        Text(Index + 1) = ""
75                    End If
76                Case 7
77                    If Text(Index) <> "" Then
                          
78                        If Log.RPLHistori = 0 Then
79                            Set oRs = objO.FetchData("SELECT * FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
80                        Else
81                            Set oRs = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
82                        End If
83                        If oRs.EOF Then
84                            gMainKeyCari = Text(Index)
85                            gSearchOption = "MDL"
86                            frmCari.show (vbModal)
87                        Else
88                            Text(Index + 1) = oRs("DESC")
89                        End If
90                        oRs.Close
91                    Else
92                        Text(Index + 1) = ""
93                    End If
94                Case 9
95                    If Text(Index) <> "" Then
96                        Set oRs = objO.FetchData("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Text(Index) & "' and FS_LOG_DEL = ''")
97                        If oRs.EOF Then
98                            gMainKeyCari = Text(Index)
99                            gBantuan = "ALL"
100                           gSearchOption = "DEALER"
101                           frmCari.show (vbModal)
102                       Else
103                           Text(Index + 1) = oRs("FS_NM_DEALER")
104                       End If
105                       oRs.Close
106                   Else
107                       Text(Index + 1) = ""
108                   End If
109               Case 11
110                   If Text(Index) <> "" Then
111                       Set oRs = objO.FetchData("SELECT TPSERVICE FROM T_REPO WITH(NOLOCK) WHERE TPSERVICE = '" & Text(Index) & "'")
112                       If oRs.EOF Then
113                           gMainKeyCari = Text(Index)
114                           gSearchOption = "TPSVC"
115                           frmCari.show (vbModal)
116                       Else
117                           Label19.Visible = False
118                           MaskEdBox11.Visible = False
119                           Text(25).Visible = False
120                           Text(24).Visible = False
121                           If Trim(Text(Index)) = "PROGRAM" Then
122                               Label19.Visible = True
123                               MaskEdBox11.Visible = True
124                               Text(25).Visible = True
125                               Text(24).Visible = True
126                           End If
127                       End If
128                       oRs.Close
129                   End If
130               Case 12
131                   If Text(Index) <> "" Then
132                       Set oRs = objO.FetchData("SELECT TPORDER FROM T_REPO WITH(NOLOCK) WHERE TPORDER = '" & Text(Index) & "' AND TPORDER <> ''")
133                       If oRs.EOF Then
134                           gMainKeyCari = Text(Index)
135                           gSearchOption = "TP_ORD"
136                           frmCari.show (vbModal)
137                       End If
138                       oRs.Close
139                   End If
140               Case 13
141                   If Text(Index) <> "" Then
142                       Set oRs = objO.FetchData("SELECT PROSES FROM T_REPO WITH(NOLOCK) WHERE PROSES = '" & Text(Index) & "'")
143                       If oRs.EOF Then
144                           gMainKeyCari = Text(Index)
145                           gSearchOption = "PROSES"
146                           frmCari.show (vbModal)
147                       End If
148                       oRs.Close
149                   End If
150               Case 24
151                   If Trim(Text(Index)) <> "" Then
                      
152                       If Log.RPLHistori = 0 Then
153                           Set rsOpsi = objO.FetchData("SELECT * FROM HO_program A WITH(NOLOCK) LEFT JOIN HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
154                       Else
155                           If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
156                               sSQL = "SELECT PROGRAM FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''"
157                               sSQL = sSQL & " UNION "
158                               sSQL = sSQL & " SELECT NAMAFP AS PROGRAM FROM [" & dbReplikasiHO & "]..HO_FIXPRICE A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_FIXPRICECBG B WITH(NOLOCK) ON A.id_FP = B.id_FP WHERE A.id_FP = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''"
159                               Set rsOpsi = objO.FetchData(sSQL)
160                           Else
161                               Set rsOpsi = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
162                           End If
                              ''Set rsOpsi = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
163                       End If
164                       If rsOpsi.EOF Then
165                           gMainKeyCari = Text(Index)
166                           gSearchOption = "PROGRAM_ALL"
167                           gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
168                           frmCari.show (vbModal)
169                       Else
170                           Text(25) = rsOpsi("PROGRAM")
171                       End If
172                   End If
173           End Select
      '    End If
End Sub
