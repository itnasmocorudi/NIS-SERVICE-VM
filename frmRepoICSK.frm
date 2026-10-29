VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoICSK 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Chasing Form"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13050
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5565
      TabIndex        =   91
      Top             =   5640
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
      Left            =   8685
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   5610
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
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   5610
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
      Left            =   1755
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   5610
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
      TabIndex        =   87
      Top             =   5640
      Width           =   1050
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   4455
      TabIndex        =   92
      Top             =   5625
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
      MICON           =   "frmRepoICSK.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoICSK.frx":001C
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   31
      Top             =   0
      Width           =   445
   End
   Begin VB.PictureBox PicBars 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   720
      ScaleHeight     =   615
      ScaleWidth      =   6975
      TabIndex        =   28
      Top             =   4260
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.ProgressBar nBars 
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nLabs 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   4935
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   600
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   9855
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
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   661
   End
   Begin TDBNumber6Ctl.TDBNumber txtPart 
      Height          =   315
      Left            =   7800
      TabIndex        =   12
      Top             =   1800
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      Calculator      =   "frmRepoICSK.frx":23EF
      Caption         =   "frmRepoICSK.frx":240F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoICSK.frx":247B
      Keys            =   "frmRepoICSK.frx":2499
      Spin            =   "frmRepoICSK.frx":24E3
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   2
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   1397882885
      MinValueVT      =   1112342533
   End
   Begin TDBNumber6Ctl.TDBNumber txtDiss 
      Height          =   315
      Left            =   7800
      TabIndex        =   13
      Top             =   2160
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      Calculator      =   "frmRepoICSK.frx":250B
      Caption         =   "frmRepoICSK.frx":252B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoICSK.frx":2597
      Keys            =   "frmRepoICSK.frx":25B5
      Spin            =   "frmRepoICSK.frx":25FF
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   2
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   1397882885
      MinValueVT      =   1112342533
   End
   Begin TDBNumber6Ctl.TDBNumber txtSats 
      Height          =   315
      Left            =   7800
      TabIndex        =   14
      Top             =   2520
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      Calculator      =   "frmRepoICSK.frx":2627
      Caption         =   "frmRepoICSK.frx":2647
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoICSK.frx":26B3
      Keys            =   "frmRepoICSK.frx":26D1
      Spin            =   "frmRepoICSK.frx":271B
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   2
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   1397882885
      MinValueVT      =   1112342533
   End
   Begin ucXPButton.XPButton cmdSave 
      Height          =   495
      Left            =   7440
      TabIndex        =   15
      Top             =   2880
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&s"
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
      MICON           =   "frmRepoICSK.frx":2743
      PICN            =   "frmRepoICSK.frx":275F
      PICH            =   "frmRepoICSK.frx":2C72
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdICSK 
      Height          =   330
      Left            =   9240
      TabIndex        =   16
      Top             =   1215
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmRepoICSK.frx":3214
      PICN            =   "frmRepoICSK.frx":3230
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdPrint 
      Height          =   495
      Left            =   7440
      TabIndex        =   17
      Top             =   3540
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&s"
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
      MICON           =   "frmRepoICSK.frx":34C4
      PICN            =   "frmRepoICSK.frx":34E0
      PICH            =   "frmRepoICSK.frx":3937
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker Tgl1 
      Height          =   375
      Left            =   960
      TabIndex        =   19
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   110886913
      CurrentDate     =   39852
   End
   Begin MSComCtl2.DTPicker Tgl2 
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   110886913
      CurrentDate     =   39852
   End
   Begin ucXPButton.XPButton TblSimpan 
      Height          =   495
      Left            =   960
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&s"
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
      MICON           =   "frmRepoICSK.frx":3E21
      PICN            =   "frmRepoICSK.frx":3E3D
      PICH            =   "frmRepoICSK.frx":4294
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboThn 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cboBln 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1200
      Width           =   855
   End
   Begin ucXPButton.XPButton cmdCetakChassingForm 
      Height          =   495
      Left            =   3720
      TabIndex        =   32
      Top             =   2160
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8454016
      BCOLO           =   8454016
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepoICSK.frx":477E
      PICN            =   "frmRepoICSK.frx":479A
      PICH            =   "frmRepoICSK.frx":47B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   10560
      TabIndex        =   44
      Top             =   480
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Frame Frame3 
         BackColor       =   &H000000FF&
         Caption         =   "MAKSIMAL"
         Height          =   3495
         Left            =   5040
         TabIndex        =   77
         Top             =   530
         Width           =   1455
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   17
            Left            =   240
            TabIndex        =   78
            Top             =   240
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   18
            Left            =   240
            TabIndex        =   79
            Top             =   600
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   19
            Left            =   240
            TabIndex        =   80
            Top             =   960
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   20
            Left            =   240
            TabIndex        =   81
            Top             =   1320
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   21
            Left            =   240
            TabIndex        =   82
            Top             =   1680
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   22
            Left            =   240
            TabIndex        =   83
            Top             =   2040
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   23
            Left            =   240
            TabIndex        =   84
            Top             =   2400
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   24
            Left            =   240
            TabIndex        =   85
            Top             =   2760
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   25
            Left            =   240
            TabIndex        =   86
            Top             =   3120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080FF80&
         Caption         =   "MINIMAL"
         Height          =   3495
         Left            =   3480
         TabIndex        =   67
         Top             =   530
         Width           =   1455
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   240
            TabIndex        =   68
            Top             =   240
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   9
            Left            =   240
            TabIndex        =   69
            Top             =   600
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   10
            Left            =   240
            TabIndex        =   70
            Top             =   960
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   240
            TabIndex        =   71
            Top             =   1320
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   240
            TabIndex        =   72
            Top             =   1680
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   13
            Left            =   240
            TabIndex        =   73
            Top             =   2040
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   240
            TabIndex        =   74
            Top             =   2400
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   15
            Left            =   240
            TabIndex        =   75
            Top             =   2760
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   16
            Left            =   240
            TabIndex        =   76
            Top             =   3120
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99:99"
            PromptChar      =   ">"
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   600
         TabIndex        =   66
         Top             =   3840
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "CETAK "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   4080
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   89
         Left            =   2400
         TabIndex        =   46
         Top             =   720
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   48
         Top             =   1080
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   50
         Top             =   1440
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   52
         Top             =   1800
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   54
         Top             =   2160
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   56
         Top             =   2520
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   58
         Top             =   2880
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   60
         Top             =   3240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   62
         Top             =   3600
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99:99"
         PromptChar      =   ">"
      End
      Begin VB.Label lblQuery 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   3960
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tgg.  Registrasi : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT. Billing : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " Tgg. Validasi : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " LT. Final Inspection : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tgg. Final Inspection : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT. Produksi : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT.  Tgg. Dikerjaan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT. Penerimaan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tgg. Penerimaan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Adjust Waktu Lead Time "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -120
         TabIndex        =   45
         Top             =   120
         Width           =   6855
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   0
         Top             =   120
         Width           =   8175
      End
   End
   Begin ucXPButton.XPButton TblKeluar 
      Height          =   495
      Left            =   13920
      TabIndex        =   26
      Top             =   4320
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
      MICON           =   "frmRepoICSK.frx":4CA0
      PICN            =   "frmRepoICSK.frx":4CBC
      PICH            =   "frmRepoICSK.frx":5190
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton tblRefresh 
      Height          =   495
      Left            =   12720
      TabIndex        =   24
      Top             =   4320
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&r"
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
      MICON           =   "frmRepoICSK.frx":56DB
      PICN            =   "frmRepoICSK.frx":56F7
      PICH            =   "frmRepoICSK.frx":5C73
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cbSublet 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12960
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ComboBox cbMasuk 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12960
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ComboBox cbJenis 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmRepoICSK.frx":62AA
      Left            =   12960
      List            =   "frmRepoICSK.frx":62AC
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ComboBox cboBlnLead 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12960
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox cboThnLead 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13920
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   1200
      Width           =   1095
   End
   Begin ucXPButton.XPButton btnCariLead 
      Height          =   495
      Left            =   13800
      TabIndex        =   33
      Top             =   3240
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&s"
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
      MICON           =   "frmRepoICSK.frx":62AE
      PICN            =   "frmRepoICSK.frx":62CA
      PICH            =   "frmRepoICSK.frx":6721
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   11160
      X2              =   15120
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Service  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11040
      TabIndex        =   41
      Top             =   1800
      Width           =   1845
   End
   Begin VB.Label Label14 
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
      Left            =   13740
      TabIndex        =   35
      Top             =   1350
      Width           =   60
   End
   Begin VB.Label nWait 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ... (^_^)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   27
      Top             =   3600
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000002&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   600
      Top             =   4200
      Width           =   14655
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      X1              =   960
      X2              =   4920
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      X1              =   960
      X2              =   4920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   25
      Top             =   1680
      Width           =   345
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
      Left            =   3780
      TabIndex        =   22
      Top             =   1350
      Width           =   60
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Service  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   21
      Top             =   1200
      Width           =   1785
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000002&
      FillColor       =   &H00FF0000&
      Height          =   3015
      Left            =   600
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  CHASSING FORM  ::."
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
      Index           =   1
      Left            =   1770
      TabIndex        =   18
      Top             =   660
      Width           =   2505
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000002&
      FillColor       =   &H00FF0000&
      Height          =   1815
      Left            =   5640
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   7680
      TabIndex        =   9
      Top             =   1800
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part Fill Rate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6240
      TabIndex        =   8
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   7680
      TabIndex        =   7
      Top             =   2160
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   7680
      TabIndex        =   6
      Top             =   2520
      Width           =   60
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Satified "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      Top             =   2520
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dissatified "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6240
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   6240
      TabIndex        =   3
      Top             =   1270
      Width           =   840
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   7200
      TabIndex        =   2
      Top             =   1270
      Width           =   60
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      FillColor       =   &H00FF0000&
      Height          =   3015
      Left            =   5640
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  I C S K   H O  ::."
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
      Index           =   0
      Left            =   7080
      TabIndex        =   1
      Top             =   660
      Width           =   1965
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmRepoICSK.frx":6C0B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periode  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11760
      TabIndex        =   36
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  CHASSING FORM  ::."
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
      Index           =   2
      Left            =   11730
      TabIndex        =   37
      Top             =   660
      Width           =   2505
   End
   Begin VB.Label lblTunggu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ... (^_^)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11040
      TabIndex        =   34
      Top             =   3600
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000002&
      FillColor       =   &H00FF0000&
      Height          =   3015
      Left            =   10560
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   10560
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "frmRepoICSK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 03 Februari 2011
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================
Dim Report As Object
Private objR As New clsRepo
Dim kriteria As String

Private Sub clearForm()
          Dim i As Integer
1         Tgl1 = Format(isnow, "DD/MM/YYYY")
2         Tgl2 = Format(isnow, "DD/MM/YYYY")
3         cboThn.Clear
4         For i = 2013 To year(isnow)
5             cboThn.AddItem i
6         Next i
7         cboThn.Text = year(isnow)
8         cboBln.Clear
9         For i = 1 To 12
10            cboBln.AddItem i
11        Next i
12        cboBln.Text = Month(isnow)
          
          '---------------------------------------------------------
          '- 2015-04-10 | RUDI | PENAMBAHAN LEAD TIME CHASSINGFORM -
          '---------------------------------------------------------
13        cboThnLead.Clear
14        For i = 2013 To year(isnow)
15            cboThnLead.AddItem i
16        Next i
17        cboThnLead.Text = year(isnow)
18        cboBlnLead.Clear
19        For i = 1 To 12
20            cboBlnLead.AddItem i
21        Next i
22        cboBlnLead.Text = Month(isnow)
          
23        cbJenis.Clear
24        cbJenis.AddItem "SBE All"
25        cbJenis.AddItem "SBE Non EM"
26        cbJenis.AddItem "EM"
27        cbJenis.AddItem "SBI"
28        cbJenis.AddItem "GR"
29        cbJenis.AddItem "Semua Service"
          
      '    cbJenis.text = "SBE All"
30        cbMasuk.Clear
31        cbMasuk.AddItem "Booking"
32        cbMasuk.AddItem "Walk In"
33        cbMasuk.AddItem "Semua"
          
34        cbSublet.Clear
35        cbSublet.AddItem "Tanpa Sublet"
36        cbSublet.AddItem "Dengan Sublet"
37        cbSublet.AddItem "All (Semua)"
          
38        cmdPrint.Enabled = False
39        cmdSave.Enabled = False
40        txtDiss = 0
41        txtSats = 0
42        txtPart = 0
43        nWait.Visible = False
          
44        cmdTutup.Visible = False
45        cmdSettingPrinter.Visible = False
46        cmdExport.Visible = False
47        XPButton1.Visible = False
48        cboPrinter.Visible = False
49        cmdCetak.Visible = False
          
50        If Log.newOhbeya = 1 Then
51            Shape2(0).Visible = False
52            Ket(0).Visible = False
              
53            Shape1.Visible = False
54            Label1(3).Visible = False
55            Label12(10).Visible = False
56            cboBln.Visible = False
57            cboThn.Visible = False
58            cmdICSK.Visible = False
              
59            Shape3.Visible = False
60            cmdSave.Visible = False
61            cmdPrint.Visible = False
62            Label5(0).Visible = False
63            Label8(0).Visible = False
64            Label9(1).Visible = False
65            Label3(0).Visible = False
66            Label7(0).Visible = False
67            Label6(0).Visible = False
              
68            txtPart.Visible = False
69            txtDiss.Visible = False
70            txtSats.Visible = False
              
              
71            Shape2(2).Left = Shape2(0).Left
72            Ket(2).Left = 6810
73            Shape6.Left = 5640
74            cboBlnLead.Left = 8040
75            cboThnLead.Left = 9000
76            Label15.Left = 6940
77            Label13.Left = 6140
78            cbJenis.Left = 8040
79            cbMasuk.Left = 8040
80            cbSublet.Left = 8040
              
81            Line1.X1 = 6120
82            Line1.X2 = 10080
83            Line1.Y1 = 3120
84            Line1.Y2 = 3120
              
85            Frame1.Visible = True
              
86            tblRefresh.Left = 8040
87            TblKeluar.Left = 9240
              
88            Text(89) = "00:10:00" '"00:05:00"
89            Text(0) = "00:15:00" '"00:15:00"
90            Text(1) = "00:10:00" '"00:10:00"
91            Text(2) = "01:20:00" '"01:00:00"
92            Text(3) = "00:10:00" '"00:05:00"
93            Text(4) = "00:10:00" '"00:10:00"
94            Text(5) = "00:05:00" '"00:05:00"
95            Text(6) = "00:05:00" '"00:05:00"
96            Text(7) = "00:05:00" '"00:05:00"
              
97            Text(8) = "00:00:00" '"00:00:30"
98            Text(9) = "00:00:00" '"00:05:00"
99            Text(10) = "00:00:00" '"00:01:00"
100           Text(11) = "00:00:00" '"00:15:00"
101           Text(12) = "00:00:00" '"00:00:30"
102           Text(13) = "00:00:00" '"00:01:00"
103           Text(14) = "00:00:00" '"00:00:30"
104           Text(15) = "00:00:00" '"00:00:30"
105           Text(16) = "00:00:00" '"00:00:30"
              
106           Text(17) = "00:00:00" '"02:00:00"
107           Text(18) = "00:00:00" '"02:00:00"
108           Text(19) = "00:00:00" '"02:00:00"
109           Text(20) = "00:00:00" '"03:00:00"
110           Text(21) = "00:00:00" '"02:00:00"
111           Text(22) = "00:00:00" '"01:00:00"
112           Text(23) = "00:00:00" '"02:00:30"
113           Text(24) = "00:00:00" '"01:00:00"
114           Text(25) = "00:00:00" '"02:00:00"
              
115           Command1.Enabled = True
116           Text(89).Enabled = True
117           Text(0).Enabled = True
118           Text(1).Enabled = True
119           Text(2).Enabled = True
120           Text(3).Enabled = True
121           Text(4).Enabled = True
122           Text(5).Enabled = True
123           Text(6).Enabled = True
124           Text(7).Enabled = True
              
125           Text(8).Enabled = True
126           Text(9).Enabled = True
127           Text(10).Enabled = True
128           Text(11).Enabled = True
129           Text(12).Enabled = True
130           Text(13).Enabled = True
131           Text(14).Enabled = True
132           Text(15).Enabled = True
133           Text(16).Enabled = True
              
134           Text(17).Enabled = True
135           Text(18).Enabled = True
136           Text(19).Enabled = True
137           Text(20).Enabled = True
138           Text(21).Enabled = True
139           Text(22).Enabled = True
140           Text(23).Enabled = True
141           Text(24).Enabled = True
142           Text(25).Enabled = True
143       End If
End Sub

Private Function getJumlahJamMenitDetik(waktu As String) As Double
          
1        On Error GoTo getJumlahJamMenitDetik_Error

          Dim JUMLAH As Double
2         waktu = Trim(waktu)
          
3         JUMLAH = CDbl((CInt(Left(waktu, 2))) * 3600) + CDbl(CInt(Mid(waktu, 4, 2)) * 60) + CInt(Right(waktu, 2))
4         getJumlahJamMenitDetik = JUMLAH
5        On Error GoTo 0
6        Exit Function

getJumlahJamMenitDetik_Error:

7         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getJumlahJamMenitDetik of Form frmRepoICSK")
End Function

Private Function getDataJamMenitDetik(total As Double) As String
1     On Error GoTo errHand
          Dim jam As Double
          Dim menit As Double
          Dim detik As Double
          
2         jam = Val(Round((total / 3600), 4))
          'IF TOTAL > (JAM*3600)
3         menit = Val((total - (jam * 3600)) / 60)
4         detik = Val(total Mod 60)
          
5         getDataJamMenitDetik = Format(jam, "0#") & ":" & Format(menit, "0#") & ":" & Format(detik, "0#")
6         Exit Function
errHand:
7         MsgBox ErrMessage(Erl, err.Number, "Function getDataJamMenitDetik"), vbInformation
End Function

Private Sub btnCariLead_Click()
On Error GoTo errHand

    If cbJenis.Text = "" Then
        MsgBox "Jenis service harus dipilih !", vbInformation
        Exit Sub
    End If
    If cbMasuk.Text = "" Then
        MsgBox "Masuk service harus dipilih !", vbInformation
        Exit Sub
    End If
    If cbSublet.Text = "" Then
        MsgBox "Sublet harus dipilih !", vbInformation
        Exit Sub
    End If
    kriteria = ""
    
    If cbJenis.Text = "SBE All" Then
        kriteria = kriteria & " (j.FS_RC <> '') AND FS_LOG_DEL = '' AND"
    ElseIf cbJenis.Text = "SBE Non EM" Then
        kriteria = kriteria & " (j.FS_RC <> '' AND j.CODE = '0') AND FS_LOG_DEL = '' AND"
    ElseIf cbJenis.Text = "SBI" Then
        kriteria = kriteria & " (j.FS_RC = 'R001') AND FS_LOG_DEL = '' AND"
    ElseIf cbJenis.Text = "GR" Then
        kriteria = kriteria & " (j.FS_RC = '') AND FS_LOG_DEL = '' AND"
    Else
        kriteria = kriteria & " FS_LOG_DEL = '' AND"
    End If
    
    If cbMasuk.Text = "Booking" Then
        kriteria = kriteria & " (J.NOFAK <> '') AND"
    ElseIf Trim(cbMasuk.Text) = "Walk In" Then
        kriteria = kriteria & " (J.NOFAK = '') AND"
    End If
    
    If cbSublet.Text = "Tanpa Sublet" Then
        kriteria = kriteria & " (SBLT = 0) AND"
    ElseIf cbSublet.Text = "Dengan Sublet" Then
        kriteria = kriteria & " (SBLT > 0) AND"
    End If
    
    kriteria = Left(kriteria, Len(kriteria) - 3)
    Frame1.Visible = True
    Text(89) = "00:05:00"
    Text(0) = "00:15:00"
    Text(1) = "00:10:00"
    Text(2) = "01:00:00"
    Text(3) = "00:05:00"
    Text(4) = "00:10:00"
    Text(5) = "00:05:00"
    Text(6) = "00:05:00"
    Text(7) = "00:05:00"
    
    Text(8) = "00:00:30"
    Text(9) = "00:05:00"
    Text(10) = "00:01:00"
    Text(11) = "00:15:00"
    Text(12) = "00:00:30"
    Text(13) = "00:01:00"
    Text(14) = "00:00:30"
    Text(15) = "00:00:30"
    Text(16) = "00:00:30"
    
    Text(17) = "02:00:00"
    Text(18) = "02:00:00"
    Text(19) = "02:00:00"
    Text(20) = "03:00:00"
    Text(21) = "02:00:00"
    Text(22) = "01:00:00"
    Text(23) = "02:00:30"
    Text(24) = "01:00:00"
    Text(25) = "02:00:00"
    
    Command1.Enabled = True
    Text(89).Enabled = True
    Text(0).Enabled = True
    Text(1).Enabled = True
    Text(2).Enabled = True
    Text(3).Enabled = True
    Text(4).Enabled = True
    Text(5).Enabled = True
    Text(6).Enabled = True
    Text(7).Enabled = True
    
    Text(8).Enabled = True
    Text(9).Enabled = True
    Text(10).Enabled = True
    Text(11).Enabled = True
    Text(12).Enabled = True
    Text(13).Enabled = True
    Text(14).Enabled = True
    Text(15).Enabled = True
    Text(16).Enabled = True
    
    Text(17).Enabled = True
    Text(18).Enabled = True
    Text(19).Enabled = True
    Text(20).Enabled = True
    Text(21).Enabled = True
    Text(22).Enabled = True
    Text(23).Enabled = True
    Text(24).Enabled = True
    Text(25).Enabled = True
    
    lblTunggu.Visible = True
    DoEvents
    Exit Sub
errHand:
    MsgBox "Error number : " & err.Number & vbCrLf & "Error Deskripsi :" & _
        err.Description & vbCrLf & " BtnCariLead - frmRepoICSK", vbInformation
    lblTunggu.Visible = False
End Sub

Private Sub cboBln_Change()
    Call FalCmd
End Sub

Private Sub cboBln_Click()
    Call FalCmd
End Sub

Private Sub cboThn_Change()
    Call FalCmd
End Sub

Private Sub cboThn_Click()
    Call FalCmd
End Sub

Private Sub FalCmd()
    cmdSave.Enabled = False
    cmdPrint.Enabled = False
End Sub

Private Sub cmdCetakChassingForm_Click()
    nWait.Visible = True
    DoEvents
    Dim oRs As Object
    Screen.MousePointer = vbHourglass
        
    Set Report = New rptRepoICSKBaru
         
    sSQL = "select CASE TanggalPanggil wHEN   '01/01/1900' tHEN 0 eLSE ISNULL(DATEDIFF(s, TanggalPanggil,(A.DATE +  A.TIME)),0) end  AS PENERIMAAN ," & _
            " ISNULL(DATEDIFF(s,tanggalstart,tanggalpause),0) as produksi," & _
            " isnull(datediff(s,tanggalstartfi,tanggalstopfi),0) as FINALCHECK, " & _
            " isnull(datediff(s,tanggalvalidasi,tanggalinvoice),0) as INVOICE, " & _
            " isnull(datediff(s,tanggalnotif,tanggalregistrasi),0) as Registrasi, " & _
            " CASE TanggalPanggil wHEN   '01/01/1900' tHEN 0 eLSE CASE TanggalIn wHEN   '01/01/1900' tHEN 0 eLSE ISNULL(DATEDIFF(s,TanggalIn, TanggalPanggil),0) end end AS TUNGGU_TERIMA, " & _
            " A.NO_POLISI, A.NO_WO, A.CUSTOMER, A.FS_RC, A.FRONT, A.MECH, A.DATE+' '+A.TIME AS TGLWO, c.NoWo,c.NoPolisi,c.Nofak,C.TanggalIn AS TGL_IN,TanggalPanggil AS TGL_PANGGIL,TanggalWO AS TGLWO, TanggalStart AS TGLON,TanggalPause AS TGLPAUSE,TanggalStartFI AS TGLCEK,TanggalStopFI AS TGLOFF,  TanggalValidasi AS TGLVALID,TanggalInvoice AS TGLINV,TanggalNotif AS TGLNOTIFIKASI,TANGGALREGISTRASI AS TGLREGISTRASI,  a.TPSVC , a.pds, a.THS, a.FN_RETURN,R.tglStartDiag,R.tglFinishDiag " & _
           " FROM ChassingForm c " & _
           " inner join JOBDTA a on a.NO_WO = c.NoWo " & _
           " LEFT JOIN ReqDiagnosis R WITH (NOLOCK) ON R.No_wo = c.nowo " & _
           " WHERE A.DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND A.FS_LOG_DEL = '' AND A.PDS = 0 AND A.TPSVC <> 'W' AND A.THS = 0 AND A.FN_RETURN = 0 ORDER BY A.NO_WO"

    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.TxtAlamat.SetText gAlamat
    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    nWait.Visible = False
    DoEvents
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub

Private Sub cmdPrint_Click()
nWait.Visible = True
DoEvents
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New RptICSK_HO
    
    sSQL = "Select A.*,B.*,C.*,D.*,E.*,F.*,G.*,H.*,I.*,J.*,K.*,L.* from " & _
           "(select Tahun as Tahun1,FIR as FIR1,LeadTimeH as LeadTimeH1,LeadTimeM as LeadTimeM1,StagnasiH as StagnasiH1,StagnasiM as StagnasiM1,ProcessH as ProcessH1,ProcessM as ProcessM1,PartFillRate as PartFillRate1,OTD as OTD1,Dissatified as Dissatified1,Satified as Satified1 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 1) A Left Join " & _
           "(select Tahun as Tahun2,FIR as FIR2,LeadTimeH as LeadTimeH2,LeadTimeM as LeadTimeM2,StagnasiH as StagnasiH2,StagnasiM as StagnasiM2,ProcessH as ProcessH2,ProcessM as ProcessM2,PartFillRate as PartFillRate2,OTD as OTD2,Dissatified as Dissatified2,Satified as Satified2 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 2) B On A.Tahun1  = B.Tahun2 Left Join " & _
           "(select Tahun as Tahun3,FIR as FIR3,LeadTimeH as LeadTimeH3,LeadTimeM as LeadTimeM3,StagnasiH as StagnasiH3,StagnasiM as StagnasiM3,ProcessH as ProcessH3,ProcessM as ProcessM3,PartFillRate as PartFillRate3,OTD as OTD3,Dissatified as Dissatified3,Satified as Satified3 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 3) C On A.Tahun1  = C.Tahun3 Left Join " & _
           "(select Tahun as Tahun4,FIR as FIR4,LeadTimeH as LeadTimeH4,LeadTimeM as LeadTimeM4,StagnasiH as StagnasiH4,StagnasiM as StagnasiM4,ProcessH as ProcessH4,ProcessM as ProcessM4,PartFillRate as PartFillRate4,OTD as OTD4,Dissatified as Dissatified4,Satified as Satified4 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 4) D On A.Tahun1  = D.Tahun4 Left Join " & _
           "(select Tahun as Tahun5,FIR as FIR5,LeadTimeH as LeadTimeH5,LeadTimeM as LeadTimeM5,StagnasiH as StagnasiH5,StagnasiM as StagnasiM5,ProcessH as ProcessH5,ProcessM as ProcessM5,PartFillRate as PartFillRate5,OTD as OTD5,Dissatified as Dissatified5,Satified as Satified5 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 5) E On A.Tahun1  = E.Tahun5 Left Join " & _
           "(select Tahun as Tahun6,FIR as FIR6,LeadTimeH as LeadTimeH6,LeadTimeM as LeadTimeM6,StagnasiH as StagnasiH6,StagnasiM as StagnasiM6,ProcessH as ProcessH6,ProcessM as ProcessM6,PartFillRate as PartFillRate6,OTD as OTD6,Dissatified as Dissatified6,Satified as Satified6 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 6) F On A.Tahun1  = F.Tahun6 Left Join " & _
           "(select Tahun as Tahun7,FIR as FIR7,LeadTimeH as LeadTimeH7,LeadTimeM as LeadTimeM7,StagnasiH as StagnasiH7,StagnasiM as StagnasiM7,ProcessH as ProcessH7,ProcessM as ProcessM7,PartFillRate as PartFillRate7,OTD as OTD7,Dissatified as Dissatified7,Satified as Satified7 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 7) G On A.Tahun1  = G.Tahun7 Left Join " & _
           "(select Tahun as Tahun8,FIR as FIR8,LeadTimeH as LeadTimeH8,LeadTimeM as LeadTimeM8,StagnasiH as StagnasiH8,StagnasiM as StagnasiM8,ProcessH as ProcessH8,ProcessM as ProcessM8,PartFillRate as PartFillRate8,OTD as OTD8,Dissatified as Dissatified8,Satified as Satified8 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 8) H On A.Tahun1  = H.Tahun8 Left Join " & _
           "(select Tahun as Tahun9,FIR as FIR9,LeadTimeH as LeadTimeH9,LeadTimeM as LeadTimeM9,StagnasiH as StagnasiH9,StagnasiM as StagnasiM9,ProcessH as ProcessH9,ProcessM as ProcessM9,PartFillRate as PartFillRate9,OTD as OTD9,Dissatified as Dissatified9,Satified as Satified9 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 9) I On A.Tahun1  = I.Tahun9 Left Join " & _
           "(select Tahun as Tahun10,FIR as FIR10,LeadTimeH as LeadTimeH10,LeadTimeM as LeadTimeM10,StagnasiH as StagnasiH10,StagnasiM as StagnasiM10,ProcessH as ProcessH10,ProcessM as ProcessM10,PartFillRate as PartFillRate10,OTD as OTD10,Dissatified as Dissatified10,Satified as Satified10 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 10) J On A.Tahun1  = J.Tahun10 Left Join " & _
           "(select Tahun as Tahun11,FIR as FIR11,LeadTimeH as LeadTimeH11,LeadTimeM as LeadTimeM11,StagnasiH as StagnasiH11,StagnasiM as StagnasiM11,ProcessH as ProcessH11,ProcessM as ProcessM11,PartFillRate as PartFillRate11,OTD as OTD11,Dissatified as Dissatified11,Satified as Satified11 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 11) K On A.Tahun1  = K.Tahun11 Left Join " & _
           "(select Tahun as Tahun12,FIR as FIR12,LeadTimeH as LeadTimeH12,LeadTimeM as LeadTimeM12,StagnasiH as StagnasiH12,StagnasiM as StagnasiM12,ProcessH as ProcessH12,ProcessM as ProcessM12,PartFillRate as PartFillRate12,OTD as OTD12,Dissatified as Dissatified12,Satified as Satified12 from ICSK_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 12) L On A.Tahun1  = L.Tahun12 Where A.Tahun1 = " & cboThn.Text
            
    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.TxtAlamat.SetText gAlamat
    
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    nWait.Visible = False
    DoEvents
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub

Private Sub cmdSave_Click()
    If MsgBox("Update data ICSK HO ?", vbQuestion + vbYesNo, "Laporan ICSK HO") = vbYes Then
        Call objR.UpdateItemICSK_manual(Val(cboThn.Text), Val(cboBln.Text), (txtPart.Text), (txtDiss.Text), (txtSats.Text))
        MsgBox "Laporan siap di Cetak !", vbInformation + vbOKOnly, "Laporan ICSK HO"
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
    cmdCetak.Visible = False
End Sub

Private Sub Command1_Click()

1        On Error GoTo Command1_Click_Error
2         If Log.newOhbeya = 1 Then
3             If cbJenis.Text = "" Then
4                 MsgBox "Jenis service harus dipilih !", vbInformation
5                 Exit Sub
6             End If
7             If cbMasuk.Text = "" Then
8                 MsgBox "Masuk service harus dipilih !", vbInformation
9                 Exit Sub
10            End If
11            If cbSublet.Text = "" Then
12                MsgBox "Sublet harus dipilih !", vbInformation
13                Exit Sub
14            End If
15            kriteria = ""
              
16            If cbJenis.Text = "SBE All" Then
17                kriteria = kriteria & " (j.FS_RC <> '' and j.FS_RC <> 'R001') AND FS_LOG_DEL = '' AND"
18            ElseIf cbJenis.Text = "SBE Non EM" Then
19                kriteria = kriteria & " (j.FS_RC <> '' AND j.CODE = '0') AND FS_LOG_DEL = '' AND"
20            ElseIf cbJenis.Text = "SBI" Then
21                kriteria = kriteria & " (j.FS_RC = 'R001') AND FS_LOG_DEL = '' AND"
22            ElseIf cbJenis.Text = "GR" Then
23                kriteria = kriteria & " (j.FS_RC = '') AND FS_LOG_DEL = '' AND"
24            Else
25                kriteria = kriteria & " FS_LOG_DEL = '' AND"
26            End If
              
27            If cbMasuk.Text = "Booking" Then
28                kriteria = kriteria & " (J.NOFAK <> '') AND"
29            ElseIf Trim(cbMasuk.Text) = "Walk In" Then
30                kriteria = kriteria & " (J.NOFAK = '') AND"
31            End If
              
32            If cbSublet.Text = "Tanpa Sublet" Then
33                kriteria = kriteria & " (SBLT = 0) AND"
34            ElseIf cbSublet.Text = "Dengan Sublet" Then
35                kriteria = kriteria & " (SBLT > 0) AND"
36            End If
              
37            kriteria = Left(kriteria, Len(kriteria) - 3)
38        End If
          Dim oRs As ADODB.Recordset
39        Screen.MousePointer = vbHourglass
40        Set Report = New rptRepoICSKLead
41        Command1.Enabled = False
42        Text(89).Enabled = False
43        Text(0).Enabled = False
44        Text(1).Enabled = False
45        Text(2).Enabled = False
46        Text(3).Enabled = False
47        Text(4).Enabled = False
48        Text(5).Enabled = False
49        Text(6).Enabled = False
50        Text(7).Enabled = False
          
51        Text(8).Enabled = False
52        Text(9).Enabled = False
53        Text(10).Enabled = False
54        Text(11).Enabled = False
55        Text(12).Enabled = False
56        Text(13).Enabled = False
57        Text(14).Enabled = False
58        Text(15).Enabled = False
59        Text(16).Enabled = False
          
60        Text(17).Enabled = False
61        Text(18).Enabled = False
62        Text(19).Enabled = False
63        Text(20).Enabled = False
64        Text(21).Enabled = False
65        Text(22).Enabled = False
66        Text(23).Enabled = False
67        Text(24).Enabled = False
68        Text(25).Enabled = False

69        sSQL = "select CASE C.TANGGALIN WHEN '01/01/1900' THEN NULL ELSE CASE C.TanggalNotif WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TANGGALIN,C.TanggalNotif) END END AS total,CASE C.TANGGALIN WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TANGGALIN,C.TANGGALPANGGIL) END AS TGGPENERIMAAN,"
70        sSQL = sSQL & " CASE TanggalPanggil WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TANGGALPANGGIL,C.TanggalWO) END AS LTPENERIMAAN,"
71        sSQL = sSQL & " CASE C.TANGGALWO WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TANGGALWO,C.TANGGALSTART) END AS LTDIKERJAKAN,"
72        sSQL = sSQL & " CASE C.TanggalStart WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TanggalStart,C.TanggalPause) END AS LTPRODUKSI,"
73        sSQL = sSQL & " CASE C.TanggalPause WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TanggalPause,C.TanggalStartFI) END AS TGGFI,"
74        sSQL = sSQL & " CASE C.TanggalStartFI WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TanggalStartFI,C.TanggalStopFI) END AS LTFI,"
75        sSQL = sSQL & " CASE C.TanggalStopFI WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TanggalStopFI,C.TanggalValidasi) END AS TGGVALIDASI,"
76        sSQL = sSQL & " CASE C.TanggalValidasi WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TanggalValidasi,C.TanggalInvoice) END AS LTINVOICE,"
77        sSQL = sSQL & " CASE C.TanggalInvoice WHEN '01/01/1900' THEN NULL ELSE CASE C.TanggalNotif WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(SECOND,C.TanggalInvoice,C.TanggalNotif) END END AS TGGREGISTRASI,"
78        sSQL = sSQL & " CASE C.TanggalWO WHEN '01/01/1900' THEN NULL ELSE DATEDIFF(day,C.TanggalWO,c.tanggalinvoice) END AS hari"
        '  sSQL = sSQL & " C.*"
79        sSQL = sSQL & " from ChassingForm c with (nolock)"
80        sSQL = sSQL & " inner join JOBDTA j with (nolock) on c.NoWo = j.NO_WO"
81        sSQL = sSQL & " where " & kriteria & " AND J.FS_LOG_DEL = '' AND PDS = 0 AND THS = 0 AND MDL NOT IN (SELECT MDL FROM [" & dbReplikasiHO & "]..MODEL WHERE [DESC] LIKE '%DYNA%') AND J.JOB = 0 and month(j.date) = '" & cboBlnLead.Text & "' and YEAR(j.date) = '" & cboThnLead.Text & "'"
          
82        lblQuery.Visible = True
83        Set oRs = objO.FetchData(sSQL)
84        lblQuery.Visible = False
          
          Dim konter As Integer
          Dim notvalid1 As Double
          Dim notValid2 As Double
          Dim notValid3 As Double
          Dim notValid4 As Double
          Dim notValid5 As Double
          Dim notValid6 As Double
          Dim notValid7 As Double
          Dim notValid8 As Double
          Dim notValid9 As Double
          Dim ikuthitung As Double
          Dim ikuttotal As Double
          
          Dim ikutHitung1 As Boolean
          Dim ikutHitung2 As Boolean
          Dim ikutHitung3 As Boolean
          Dim ikutHitung4 As Boolean
          Dim ikutHitung5 As Boolean
          Dim ikutHitung6 As Boolean
          Dim ikutHitung7 As Boolean
          Dim ikutHitung8 As Boolean
          Dim ikutHitung9 As Boolean
          
          Dim min1 As Double
          Dim min2 As Double
          Dim min3 As Double
          Dim min4 As Double
          Dim min5 As Double
          Dim min6 As Double
          Dim min7 As Double
          Dim min8 As Double
          Dim min9 As Double
          
          Dim mak1 As Double
          Dim mak2 As Double
          Dim mak3 As Double
          Dim mak4 As Double
          Dim mak5 As Double
          Dim mak6 As Double
          Dim mak7 As Double
          Dim mak8 As Double
          Dim mak9 As Double
          
          Dim rata1 As Double
          Dim rata2 As Double
          Dim rata3 As Double
          Dim rata4 As Double
          Dim rata5 As Double
          Dim rata6 As Double
          Dim rata7 As Double
          Dim rata8 As Double
          Dim rata9 As Double
          
85        If Not oRs.EOF Then
86            ProgressBar1.Max = oRs.RecordCount + 1
87            ProgressBar1.Min = 0
88            konter = 1
89            oRs.MoveFirst
90            notvalid1 = 0
91            notValid2 = 0
92            notValid3 = 0
93            notValid4 = 0
94            notValid5 = 0
95            notValid6 = 0
96            notValid7 = 0
97            notValid8 = 0
98            notValid9 = 0
99            ikuthitung = 0
100           ikuttotal = 0
              
101           min1 = 0
102           min2 = 0
103           min3 = 0
104           min4 = 0
105           min5 = 0
106           min6 = 0
107           min7 = 0
108           min8 = 0
109           min9 = 0
              
110           mak1 = 0
111           mak2 = 0
112           mak3 = 0
113           mak4 = 0
114           mak5 = 0
115           mak6 = 0
116           mak7 = 0
117           mak8 = 0
118           mak9 = 0
              
119           rata1 = 0
120           rata2 = 0
121           rata3 = 0
122           rata4 = 0
123           rata5 = 0
124           rata6 = 0
125           rata7 = 0
126           rata8 = 0
127           rata9 = 0
              
128           Do While Not oRs.EOF
129               ProgressBar1.value = zx
                  
130               ikutHitung1 = True
131               ikutHitung2 = True
132               ikutHitung3 = True
133               ikutHitung4 = True
134               ikutHitung5 = True
135               ikutHitung6 = True
136               ikutHitung7 = True
137               ikutHitung8 = True
138               ikutHitung9 = True
                  
139               If IsNull(oRs("hari")) Then
140                   notvalid1 = notvalid1 + 1
141                   notValid2 = notValid2 + 1
142                   notValid3 = notValid3 + 1
143                   notValid4 = notValid4 + 1
144                   notValid5 = notValid5 + 1
145                   notValid6 = notValid6 + 1
146                   notValid7 = notValid7 + 1
147                   notValid8 = notValid8 + 1
148                   notValid9 = notValid9 + 1
                      
149                   ikutHitung1 = False
150                   ikutHitung2 = False
151                   ikutHitung3 = False
152                   ikutHitung4 = False
153                   ikutHitung5 = False
154                   ikutHitung6 = False
155                   ikutHitung7 = False
156                   ikutHitung8 = False
157                   ikutHitung9 = False
158               ElseIf oRs("hari") > 0 Then
159                   notvalid1 = notvalid1 + 1
160                   notValid2 = notValid2 + 1
161                   notValid3 = notValid3 + 1
162                   notValid4 = notValid4 + 1
163                   notValid5 = notValid5 + 1
164                   notValid6 = notValid6 + 1
165                   notValid7 = notValid7 + 1
166                   notValid8 = notValid8 + 1
167                   notValid9 = notValid9 + 1
                      
168                   ikutHitung1 = False
169                   ikutHitung2 = False
170                   ikutHitung3 = False
171                   ikutHitung4 = False
172                   ikutHitung5 = False
173                   ikutHitung6 = False
174                   ikutHitung7 = False
175                   ikutHitung8 = False
176                   ikutHitung9 = False
177               Else
                      'LAKUKAN PENGECEKAN DATA VALID
178                   If IsNull(oRs("TGGPENERIMAAN")) Then
179                       GoTo notvalid1
180                   ElseIf oRs("TGGPENERIMAAN") < 0 Then
181                       GoTo notvalid1
182                   ElseIf Log.newOhbeya = 1 And (getJumlahJamMenitDetik(Trim(Text(8))) = 0 And getJumlahJamMenitDetik(Trim(Text(17))) = 0) Then
183                       GoTo lanjut1
184                   ElseIf getJumlahJamMenitDetik(Trim(Text(8))) > oRs("TGGPENERIMAAN") Or getJumlahJamMenitDetik(Trim(Text(17))) < oRs("TGGPENERIMAAN") Then
185                       GoTo notvalid1
186                   End If
187                   GoTo lanjut1
notvalid1:
188       notvalid1 = notvalid1 + 1
189       ikutHitung1 = False
lanjut1:
                      
190                   If IsNull(oRs("LTPENERIMAAN")) Then
191                       GoTo notValid2
192                   ElseIf oRs("LTPENERIMAAN") < 0 Then
193                       GoTo notValid2
194                   ElseIf Log.newOhbeya = 1 And (getJumlahJamMenitDetik(Trim(Text(9))) = 0 And getJumlahJamMenitDetik(Trim(Text(18))) = 0) Then
195                       GoTo Lanjut2
196                   ElseIf getJumlahJamMenitDetik(Trim(Text(9))) > oRs("LTPENERIMAAN") Or getJumlahJamMenitDetik(Trim(Text(18))) < oRs("LTPENERIMAAN") Then
197                       GoTo notValid2
198                   End If
199                   GoTo Lanjut2
notValid2:
200       notValid2 = notValid2 + 1
201       ikutHitung2 = False
Lanjut2:
202                   If IsNull(oRs("LTDIKERJAKAN")) Then
203                       GoTo notValid3
204                   ElseIf oRs("LTDIKERJAKAN") < 0 Then
205                       GoTo notValid3
206                   ElseIf Log.newOhbeya = 1 And (getJumlahJamMenitDetik(Trim(Text(10))) = 0 And getJumlahJamMenitDetik(Trim(Text(19))) = 0) Then
207                       GoTo Lanjut3
208                   ElseIf getJumlahJamMenitDetik(Trim(Text(10))) > oRs("LTDIKERJAKAN") Or getJumlahJamMenitDetik(Trim(Text(19))) < oRs("LTDIKERJAKAN") Then
209                       GoTo notValid3
210                   End If
211                   GoTo Lanjut3
notValid3:
212       notValid3 = notValid3 + 1
213       ikutHitung3 = False
Lanjut3:
                      
214                   If IsNull(oRs("LTPRODUKSI")) Then
215                       GoTo notValid4
216                   ElseIf oRs("LTPRODUKSI") < 0 Then
217                       GoTo notValid4
218                   ElseIf Log.newOhbeya = 1 And (getJumlahJamMenitDetik(Trim(Text(11))) = 0 And getJumlahJamMenitDetik(Trim(Text(20))) = 0) Then
219                       GoTo Lanjut4
220                   ElseIf getJumlahJamMenitDetik(Trim(Text(11))) > oRs("LTPRODUKSI") Or getJumlahJamMenitDetik(Trim(Text(20))) < oRs("LTPRODUKSI") Then
221                       GoTo notValid4
222                   End If
223                   GoTo Lanjut4
notValid4:
224       notValid4 = notValid4 + 1
225       ikutHitung4 = False
Lanjut4:
                      
226                   If IsNull(oRs("TGGFI")) Then
227                       GoTo notValid5
228                   ElseIf oRs("TGGFI") < 0 Then
229                       GoTo notValid5
230                   ElseIf Log.newOhbeya = 1 And (getJumlahJamMenitDetik(Trim(Text(12))) = 0 And getJumlahJamMenitDetik(Trim(Text(21))) = 0) Then
231                       GoTo Lanjut5
232                   ElseIf getJumlahJamMenitDetik(Trim(Text(12))) > oRs("TGGFI") Or getJumlahJamMenitDetik(Trim(Text(21))) < oRs("TGGFI") Then
233                       GoTo notValid5
234                   End If
235                   GoTo Lanjut5
notValid5:
236       notValid5 = notValid5 + 1
237       ikutHitung5 = False
Lanjut5:
                      
238                   If IsNull(oRs("LTFI")) Then
239                       GoTo notValid6
240                   ElseIf oRs("LTFI") < 0 Then
241                       GoTo notValid6
242                   ElseIf Log.newOhbeya = 1 And (getJumlahJamMenitDetik(Trim(Text(13))) = 0 And getJumlahJamMenitDetik(Trim(Text(22))) = 0) Then
243                       GoTo Lanjut6
244                   ElseIf getJumlahJamMenitDetik(Trim(Text(13))) > oRs("LTFI") Or getJumlahJamMenitDetik(Trim(Text(22))) < oRs("LTFI") Then
245                       GoTo notValid6
246                   End If
247                   GoTo Lanjut6
notValid6:
248       notValid6 = notValid6 + 1
249       ikutHitung6 = False
Lanjut6:
                      
250                   If IsNull(oRs("TGGVALIDASI")) Then
251                       GoTo notValid7
252                   ElseIf oRs("TGGVALIDASI") < 0 Then
253                       GoTo notValid7
254                   ElseIf Log.newOhbeya = 1 And (getJumlahJamMenitDetik(Trim(Text(14))) = 0 And getJumlahJamMenitDetik(Trim(Text(23))) = 0) Then
255                       GoTo Lanjut7
256                   ElseIf getJumlahJamMenitDetik(Trim(Text(14))) > oRs("TGGVALIDASI") Or getJumlahJamMenitDetik(Trim(Text(23))) < oRs("TGGVALIDASI") Then
257                       GoTo notValid7
258                   End If
259                   GoTo Lanjut7
notValid7:
260       notValid7 = notValid7 + 1
261       ikutHitung7 = False
Lanjut7:
                      
262                   If IsNull(oRs("LTINVOICE")) Then
263                       GoTo notValid8
264                   ElseIf oRs("LTINVOICE") < 0 Then
265                       GoTo notValid8
266                   ElseIf Log.newOhbeya = 1 And (getJumlahJamMenitDetik(Trim(Text(15))) = 0 And getJumlahJamMenitDetik(Trim(Text(24))) = 0) Then
267                       GoTo Lanjut8
268                   ElseIf getJumlahJamMenitDetik(Trim(Text(15))) > oRs("LTINVOICE") Or getJumlahJamMenitDetik(Trim(Text(24))) < oRs("LTINVOICE") Then
269                       GoTo notValid8
270                   End If
271                   GoTo Lanjut8
notValid8:
272       notValid8 = notValid8 + 1
273       ikutHitung8 = False
Lanjut8:
                      
274                   If IsNull(oRs("TGGREGISTRASI")) Then
275                       GoTo notValid9
276                   ElseIf oRs("TGGREGISTRASI") < 0 Then
277                       GoTo notValid9
278                   ElseIf Log.newOhbeya = 1 And (getJumlahJamMenitDetik(Trim(Text(16))) = 0 And getJumlahJamMenitDetik(Trim(Text(25))) = 0) Then
279                       GoTo Lanjut9
280                   ElseIf getJumlahJamMenitDetik(Trim(Text(16))) > oRs("TGGREGISTRASI") Or getJumlahJamMenitDetik(Trim(Text(25))) < oRs("TGGREGISTRASI") Then
281                       GoTo notValid9
282                   End If
283                   GoTo Lanjut9
notValid9:
284       notValid9 = notValid9 + 1
285       ikutHitung9 = False
Lanjut9:
286               End If
287               If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then
288                   ikuthitung = ikuthitung + 1
289                   ikuttotal = ikuttotal + oRs("total")
290               End If
                  '-- HITUNG KRITERIA 1 ----------
291               If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then
                      'proses perhitungan minimal maksimal dan rata2
292                   If min1 = 0 Then
293                       min1 = oRs("TGGPENERIMAAN")
294                   ElseIf min1 > oRs("TGGPENERIMAAN") Then
295                       min1 = oRs("TGGPENERIMAAN")
296                   End If
                      
297                   If mak1 = 0 Then
298                       mak1 = oRs("TGGPENERIMAAN")
299                   ElseIf mak1 < oRs("TGGPENERIMAAN") Then
300                       mak1 = oRs("TGGPENERIMAAN")
301                   End If
                      
302                   rata1 = rata1 + oRs("TGGPENERIMAAN")
303               End If
                  '----- END KRITERIA 1 -------------
                  
                  '-- HITUNG KRITERIA 2 ----------
304              If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then '   If ikutHitung2 Then
                      'proses perhitungan minimal maksimal dan rata2
305                   If min2 = 0 Then
306                       min2 = oRs("LTPENERIMAAN")
307                   ElseIf min2 > oRs("LTPENERIMAAN") Then
308                       min2 = oRs("LTPENERIMAAN")
309                   End If
                      
310                   If mak2 = 0 Then
311                       mak2 = oRs("LTPENERIMAAN")
312                   ElseIf mak2 < oRs("LTPENERIMAAN") Then
313                       mak2 = oRs("LTPENERIMAAN")
314                   End If
                      
315                   rata2 = rata2 + oRs("LTPENERIMAAN")
316               End If
                  '----- END KRITERIA 2 -------------
                  
                  '-- HITUNG KRITERIA 3 ----------
317               If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then 'If ikutHitung3 Then
                      'proses perhitungan minimal maksimal dan rata2
318                   If min3 = 0 Then
319                       min3 = oRs("LTDIKERJAKAN")
320                   ElseIf min3 > oRs("LTDIKERJAKAN") Then
321                       min3 = oRs("LTDIKERJAKAN")
322                   End If
                      
323                   If mak3 = 0 Then
324                       mak3 = oRs("LTDIKERJAKAN")
325                   ElseIf mak3 < oRs("LTDIKERJAKAN") Then
326                       mak3 = oRs("LTDIKERJAKAN")
327                   End If
                      
328                   rata3 = rata3 + oRs("LTDIKERJAKAN")
329               End If
                  '----- END KRITERIA 3 -------------
                  
                  '-- HITUNG KRITERIA 4 ----------
330               If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then 'If ikutHitung4 Then
                      'proses perhitungan minimal maksimal dan rata2
331                   If min4 = 0 Then
332                       min4 = oRs("LTPRODUKSI")
333                   ElseIf min4 > oRs("LTPRODUKSI") Then
334                       min4 = oRs("LTPRODUKSI")
335                   End If
                      
336                   If mak4 = 0 Then
337                       mak4 = oRs("LTPRODUKSI")
338                   ElseIf mak4 < oRs("LTPRODUKSI") Then
339                       mak4 = oRs("LTPRODUKSI")
340                   End If
                      
341                   rata4 = rata4 + oRs("LTPRODUKSI")
342               End If
                  '----- END KRITERIA 4 -------------
                  
                  '-- HITUNG KRITERIA 5 ----------
343               If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then 'If ikutHitung5 Then
                      'proses perhitungan minimal maksimal dan rata2
344                   If min5 = 0 Then
345                       min5 = oRs("TGGFI")
346                   ElseIf min5 > oRs("TGGFI") Then
347                       min5 = oRs("TGGFI")
348                   End If
                      
349                   If mak5 = 0 Then
350                       mak5 = oRs("TGGFI")
351                   ElseIf mak5 < oRs("TGGFI") Then
352                       mak5 = oRs("TGGFI")
353                   End If
                      
354                   rata5 = rata5 + oRs("TGGFI")
355               End If
                  '----- END KRITERIA 5 -------------
                  
                  '-- HITUNG KRITERIA 6 ----------
356               If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then 'If ikutHitung6 Then
                      'proses perhitungan minimal maksimal dan rata2
357                   If min6 = 0 Then
358                       min6 = oRs("LTFI")
359                   ElseIf min6 > oRs("LTFI") Then
360                       min6 = oRs("LTFI")
361                   End If
                      
362                   If mak6 = 0 Then
363                       mak6 = oRs("LTFI")
364                   ElseIf mak6 < oRs("LTFI") Then
365                       mak6 = oRs("LTFI")
366                   End If
                      
367                   rata6 = rata6 + oRs("LTFI")
368               End If
                  '----- END KRITERIA 6 -------------
                  
                  '-- HITUNG KRITERIA 7 ----------
369               If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then 'If ikutHitung7 Then
                      'proses perhitungan minimal maksimal dan rata2
370                   If min7 = 0 Then
371                       min7 = oRs("TGGVALIDASI")
372                   ElseIf min7 > oRs("TGGVALIDASI") Then
373                       min7 = oRs("TGGVALIDASI")
374                   End If
                      
375                   If mak7 = 0 Then
376                       mak7 = oRs("TGGVALIDASI")
377                   ElseIf mak7 < oRs("TGGVALIDASI") Then
378                       mak7 = oRs("TGGVALIDASI")
379                   End If
                      
380                   rata7 = rata7 + oRs("TGGVALIDASI")
381               End If
                  '----- END KRITERIA 7 -------------
                  
                  '-- HITUNG KRITERIA 8 ----------
382               If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then 'If ikutHitung8 Then
                      'proses perhitungan minimal maksimal dan rata2
383                   If min8 = 0 Then
384                       min8 = oRs("LTINVOICE")
385                   ElseIf min8 > oRs("LTINVOICE") Then
386                       min8 = oRs("LTINVOICE")
387                   End If
                      
388                   If mak8 = 0 Then
389                       mak8 = oRs("LTINVOICE")
390                   ElseIf mak8 < oRs("LTINVOICE") Then
391                       mak8 = oRs("LTINVOICE")
392                   End If
                      
393                   rata8 = rata8 + oRs("LTINVOICE")
394               End If
                  '----- END KRITERIA 8 -------------
                  
                  '-- HITUNG KRITERIA 9 ----------
395               If ikutHitung1 And ikutHitung2 And ikutHitung3 And ikutHitung4 And ikutHitung5 And ikutHitung6 And ikutHitung7 And ikutHitung8 And ikutHitung9 Then 'If ikutHitung9 Then
                      'proses perhitungan minimal maksimal dan rata2
396                   If min9 = 0 Then
397                       min9 = oRs("TGGREGISTRASI")
398                   ElseIf min9 > oRs("TGGREGISTRASI") Then
399                       min9 = oRs("TGGREGISTRASI")
400                   End If
                      
401                   If mak9 = 0 Then
402                       mak9 = oRs("TGGREGISTRASI")
403                   ElseIf mak9 < oRs("TGGREGISTRASI") Then
404                       mak9 = oRs("TGGREGISTRASI")
405                   End If
                      
406                   rata9 = rata9 + oRs("TGGREGISTRASI")
407               End If
                  '----- END KRITERIA 9 -------------
408               konter = konter + 1
409               oRs.MoveNext
410           Loop
411           oRs.MoveFirst
              
412           If (oRs.RecordCount - notvalid1) > 0 Then
413               rata1 = rata1 / IIf(ikuthitung > 0, ikuthitung, 1) 'rata1 = rata1 / (oRs.RecordCount - notvalid1)
414           Else
415               rata1 = 0
416           End If
              
417           If (oRs.RecordCount - notValid2) > 0 Then
418               rata2 = rata2 / IIf(ikuthitung > 0, ikuthitung, 1) 'rata2 = rata2 / (oRs.RecordCount - notValid2)
419           Else
420               rata2 = 0
421           End If
              
422           If (oRs.RecordCount - notValid3) > 0 Then
423               rata3 = rata3 / IIf(ikuthitung > 0, ikuthitung, 1) 'rata3 = rata3 / (oRs.RecordCount - notValid3)
424           Else
425               rata3 = 0
426           End If
              
427           If (oRs.RecordCount - notValid4) > 0 Then
428               rata4 = rata4 / IIf(ikuthitung > 0, ikuthitung, 1) 'rata4 = rata4 / (oRs.RecordCount - notValid4)
429           Else
430               rata4 = 0
431           End If
              
432           If (oRs.RecordCount - notValid5) > 0 Then
433               rata5 = rata5 / IIf(ikuthitung > 0, ikuthitung, 1) 'rata5 = rata5 / (oRs.RecordCount - notValid5)
434           Else
435               rata5 = 0
436           End If
              
437           If (oRs.RecordCount - notValid6) > 0 Then
438               rata6 = rata6 / IIf(ikuthitung > 0, ikuthitung, 1) 'rata6 = rata6 / (oRs.RecordCount - notValid6)
439           Else
440               rata6 = 0
441           End If
              
442           If (oRs.RecordCount - notValid7) > 0 Then
443               rata7 = rata7 / IIf(ikuthitung > 0, ikuthitung, 1) 'rata7 = rata7 / (oRs.RecordCount - notValid7)
444           Else
445               rata7 = 0
446           End If
              
447           If (oRs.RecordCount - notValid8) > 0 Then
448               rata8 = rata8 / IIf(ikuthitung > 0, ikuthitung, 1) 'rata8 = rata8 / (oRs.RecordCount - notValid8)
449           Else
450               rata8 = 0
451           End If
              
452           If (oRs.RecordCount - notValid9) > 0 Then
453               rata9 = rata9 / IIf(ikuthitung > 0, ikuthitung, 1) 'rata9 = rata9 / (oRs.RecordCount - notValid9)
454           Else
455               rata9 = 0
456           End If
              
457       End If
          
      '    'Report.Database.SetDataSource oRs
      '    Report.txtPerusahaan.SetText gPerush
      '    Report.txtAlamat.SetText gAlamat
      '    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & _
      '        " s/d " & Format(Tgl2, "DD-MM-YYYY")
          
          '------------------------------------------
458       Report.txtSBE.SetText cbJenis.Text
459       Report.txtJumlahData.SetText "Jumlah Data = " & oRs.RecordCount
460       Report.txtTarget1.SetText Trim(Text(89))
461       Report.txtTarget2.SetText Trim(Text(0))
462       Report.txtTarget3.SetText Trim(Text(1))
463       Report.txtTarget4.SetText Trim(Text(2))
464       Report.txtTarget5.SetText Trim(Text(3))
465       Report.txtTarget6.SetText Trim(Text(4))
466       Report.txtTarget7.SetText Trim(Text(5))
467       Report.txtTarget8.SetText Trim(Text(6))
468       Report.txtTarget9.SetText Trim(Text(7))
          
469       Report.txtSublet.SetText Trim(cbSublet.Text)
          
470       Report.txtLTMin1.SetText Trim(Text(8))
471       Report.txtLTMin2.SetText Trim(Text(9))
472       Report.txtLTMin3.SetText Trim(Text(10))
473       Report.txtLTMin4.SetText Trim(Text(11))
474       Report.txtLTMin5.SetText Trim(Text(12))
475       Report.txtLTMin6.SetText Trim(Text(13))
476       Report.txtLTMin7.SetText Trim(Text(14))
477       Report.txtLTMin8.SetText Trim(Text(15))
478       Report.txtLTMin9.SetText Trim(Text(16))
          
479       Report.txtLTMax1.SetText Trim(Text(17))
480       Report.txtLTMax2.SetText Trim(Text(18))
481       Report.txtLTMax3.SetText Trim(Text(19))
482       Report.txtLTMax4.SetText Trim(Text(20))
483       Report.txtLTMax5.SetText Trim(Text(21))
484       Report.txtLTMax6.SetText Trim(Text(22))
485       Report.txtLTMax7.SetText Trim(Text(23))
486       Report.txtLTMax8.SetText Trim(Text(24))
487       Report.txtLTMax9.SetText Trim(Text(25))
          
488       Report.txtJumlah1.SetText oRs.RecordCount - notvalid1
489       Report.txtJumlah2.SetText oRs.RecordCount - notValid2
490       Report.txtJumlah3.SetText oRs.RecordCount - notValid3
491       Report.txtJumlah4.SetText oRs.RecordCount - notValid4
492       Report.txtJumlah5.SetText oRs.RecordCount - notValid5
493       Report.txtJumlah6.SetText oRs.RecordCount - notValid6
494       Report.txtJumlah7.SetText oRs.RecordCount - notValid7
495       Report.txtJumlah8.SetText oRs.RecordCount - notValid8
496       Report.txtJumlah9.SetText oRs.RecordCount - notValid9
497       Report.txtJumlahTotal.SetText (oRs.RecordCount - notvalid1) + (oRs.RecordCount - notValid2) + (oRs.RecordCount - notValid3) + (oRs.RecordCount - notValid4) + (oRs.RecordCount - notValid5) + (oRs.RecordCount - notValid6) + (oRs.RecordCount - notValid7) + (oRs.RecordCount - notValid8) + (oRs.RecordCount - notValid9)
          
498       If oRs.RecordCount > 0 Then
499           Report.txtRasio1.SetText Val((((oRs.RecordCount - notvalid1) / oRs.RecordCount) * 100)) & "%"
500           Report.txtRasio2.SetText Val((((oRs.RecordCount - notValid2) / oRs.RecordCount) * 100)) & "%"
501           Report.txtRasio3.SetText Val((((oRs.RecordCount - notValid3) / oRs.RecordCount) * 100)) & "%"
502           Report.txtRasio4.SetText Val((((oRs.RecordCount - notValid4) / oRs.RecordCount) * 100)) & "%"
503           Report.txtRasio5.SetText Val((((oRs.RecordCount - notValid5) / oRs.RecordCount) * 100)) & "%"
504           Report.txtRasio6.SetText Val((((oRs.RecordCount - notValid6) / oRs.RecordCount) * 100)) & "%"
505           Report.txtRasio7.SetText Val((((oRs.RecordCount - notValid7) / oRs.RecordCount) * 100)) & "%"
506           Report.txtRasio8.SetText Val((((oRs.RecordCount - notValid8) / oRs.RecordCount) * 100)) & "%"
507           Report.txtRasio9.SetText Val((((oRs.RecordCount - notValid9) / oRs.RecordCount) * 100)) & "%"
508           Report.txtRasioTotal.SetText ((Val((((oRs.RecordCount - notvalid1) / oRs.RecordCount) * 100)) + Val((((oRs.RecordCount - notValid2) / oRs.RecordCount) * 100)) + Val((((oRs.RecordCount - notValid3) / oRs.RecordCount) * 100)) + Val((((oRs.RecordCount - notValid4) / oRs.RecordCount) * 100)) + Val((((oRs.RecordCount - notValid5) / oRs.RecordCount) * 100)) + Val((((oRs.RecordCount - notValid6) / oRs.RecordCount) * 100)) + Val((((oRs.RecordCount - notValid7) / oRs.RecordCount) * 100)) + Val((((oRs.RecordCount - notValid8) / oRs.RecordCount) * 100)) + Val((((oRs.RecordCount - notValid9) / oRs.RecordCount) * 100))) / 9) & " %"
509       Else
510           Report.txtRasio1.SetText "0%"
511           Report.txtRasio2.SetText "0%"
512           Report.txtRasio3.SetText "0%"
513           Report.txtRasio4.SetText "0%"
514           Report.txtRasio5.SetText "0%"
515           Report.txtRasio6.SetText "0%"
516           Report.txtRasio7.SetText "0%"
517           Report.txtRasio8.SetText "0%"
518           Report.txtRasio9.SetText "0%"
519           Report.txtRasioTotal.SetText "0%"
520       End If
          
521       Report.txtAvg1.SetText getDataJamMenitDetik(rata1)
522       If rata1 > getJumlahJamMenitDetik(Trim(Text(89))) Then Report.txtAvg1.TextColor = vbRed
          
523       Report.txtAvg2.SetText getDataJamMenitDetik(rata2)
524       If rata2 > getJumlahJamMenitDetik(Trim(Text(0))) Then Report.txtAvg2.TextColor = vbRed

525       Report.txtAvg3.SetText getDataJamMenitDetik(rata3)
526       If rata3 > getJumlahJamMenitDetik(Trim(Text(1))) Then Report.txtAvg3.TextColor = vbRed

527       Report.txtAvg4.SetText getDataJamMenitDetik(rata4)
528       If rata4 > getJumlahJamMenitDetik(Trim(Text(2))) Then Report.txtAvg4.TextColor = vbRed

529       Report.txtAvg5.SetText getDataJamMenitDetik(rata5)
530       If rata5 > getJumlahJamMenitDetik(Trim(Text(3))) Then Report.txtAvg5.TextColor = vbRed

531       Report.txtAvg6.SetText getDataJamMenitDetik(rata6)
532       If rata6 > getJumlahJamMenitDetik(Trim(Text(4))) Then Report.txtAvg6.TextColor = vbRed

533       Report.txtAvg7.SetText getDataJamMenitDetik(rata7)
534       If rata7 > getJumlahJamMenitDetik(Trim(Text(5))) Then Report.txtAvg7.TextColor = vbRed

535       Report.txtAvg8.SetText getDataJamMenitDetik(rata8)
536       If rata8 > getJumlahJamMenitDetik(Trim(Text(6))) Then Report.txtAvg8.TextColor = vbRed

537       Report.txtAvg9.SetText getDataJamMenitDetik(rata9)
538       If rata9 > getJumlahJamMenitDetik(Trim(Text(7))) Then Report.txtAvg9.TextColor = vbRed
          
          ''Report.txtAVGTotal.SetText getDataJamMenitDetik(((rata9 + rata1 + rata2 + rata3 + rata4 + rata5 + rata6 + rata7 + rata8) / 9))
          '' 20200113DEV1
539       Report.txtAVGTotal.SetText IIf(ikuthitung > 0, getDataJamMenitDetik(ikuttotal / ikuthitung), "00:00:00") ''Report.txtAVGTotal.SetText getDataJamMenitDetik(ikuttotal / ikuthitung) ''getDataJamMenitDetik(rata9 + rata1 + rata2 + rata3 + rata4 + rata5 + rata6 + rata7 + rata8)
          
540       Report.txtLTProsesMin1.SetText getDataJamMenitDetik(min1)
541       Report.txtLTProsesMin2.SetText getDataJamMenitDetik(min2)
542       Report.txtLTProsesMin3.SetText getDataJamMenitDetik(min3)
543       Report.txtLTProsesMin4.SetText getDataJamMenitDetik(min4)
544       Report.txtLTProsesMin5.SetText getDataJamMenitDetik(min5)
545       Report.txtLTProsesMin6.SetText getDataJamMenitDetik(min6)
546       Report.txtLTProsesMin7.SetText getDataJamMenitDetik(min7)
547       Report.txtLTProsesMin8.SetText getDataJamMenitDetik(min8)
548       Report.txtLTProsesMin9.SetText getDataJamMenitDetik(min9)
          
549       Report.txtLTProsesMax1.SetText getDataJamMenitDetik(mak1)
550       Report.txtLTProsesMax2.SetText getDataJamMenitDetik(mak2)
551       Report.txtLTProsesMax3.SetText getDataJamMenitDetik(mak3)
552       Report.txtLTProsesMax4.SetText getDataJamMenitDetik(mak4)
553       Report.txtLTProsesMax5.SetText getDataJamMenitDetik(mak5)
554       Report.txtLTProsesMax6.SetText getDataJamMenitDetik(mak6)
555       Report.txtLTProsesMax7.SetText getDataJamMenitDetik(mak7)
556       Report.txtLTProsesMax8.SetText getDataJamMenitDetik(mak8)
557       Report.txtLTProsesMax9.SetText getDataJamMenitDetik(mak9)
          
558       Report.txtPeriode.SetText "PERIODE = " & GetNamaBulan(cboBlnLead.Text, 1) & " " & cboThnLead.Text
559       Report.txtBooking.SetText UCase(cbMasuk.Text)
          '------------------------------------------
560       With CRViewer1
561           .ReportSource = Report
562           .ViewReport
563           .Zoom 100
564           .Visible = True
565       End With
          
566       nWait.Visible = False
567       DoEvents
568       cmdTutup.Visible = True
569       cmdSettingPrinter.Visible = True
570       Screen.MousePointer = vbDefault
571       If Log.newOhbeya = 0 Then
572           Frame1.Visible = False
573       End If
574       Call DisableExport
575       Call setLokasiTombol
576       Exit Sub
Command1_Click_Error:

577       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Command1_Click of Form frmRepoICSK"), vbInformation
End Sub

Private Sub Form_Load()
    Call clearForm
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

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call clearForm
End Sub

Private Sub tblSimpan_Click()
nWait.Visible = True
DoEvents
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New rptRepoICSK
         
    sSQL = "SELECT A.NO_POLISI, A.NO_WO, A.CUSTOMER, A.FS_RC, A.FRONT, A.MECH, TGL_IN, C.TGL_PANGGIL, A.DATE+' '+A.TIME AS TGLWO , " & _
           "(SELECT TOP 1 TANGGAL+' '+START FROM TEMPPROGRESS WHERE NOORDER = A.NO_WO AND FLAGP <> 'M' ORDER BY TANGGAL, START) AS TGLON , " & _
           "(SELECT TOP 1 TANGGAL+' '+JAM FROM PROGRESS WHERE NOORDER = A.NO_WO AND LEFT(KETERANGAN, CHARINDEX(' |',KETERANGAN))= 'Final Check' ORDER BY TANGGAL DESC, JAM DESC) AS TGLPAUSE , " & _
           "(SELECT TOP 1 TANGGAL+' '+JAM FROM PROGRESS WHERE NOORDER = A.NO_WO AND FLAG = 'Q' ORDER BY TANGGAL DESC, JAM DESC) AS TGLCEK , " & _
           "(SELECT TOP 1 TANGGAL+' '+PAUSE FROM TEMPPROGRESS WHERE NOORDER = A.NO_WO AND FLAGP = 'F' ORDER BY TANGGAL DESC, START DESC) AS TGLOFF , " & _
           "(SELECT TOP 1 TANGGAL FROM VALIDORDER WHERE NO_WO = A.NO_WO AND USERDEL = '') AS TGLVALID , " & _
           "E.INVDATE+' '+INVTIME AS TGLINV, F.TglPanggil AS tglNotifikasi, TglKasir AS tglRegistrasi, " & _
           "A.TPSVC, A.PDS, A.THS, A.FN_RETURN, ISNULL(DATEDIFF(s, TGL_IN, C.TGL_PANGGIL),0) AS TUNGGU_TERIMA, " & _
           "ISNULL(DATEDIFF(s, C.TGL_PANGGIL, (A.DATE+' '+A.TIME)),0) AS PENERIMAAN , " & _
           "ISNULL(DATEDIFF(s, ((SELECT TOP 1 TANGGAL+' '+START FROM TEMPPROGRESS WHERE NOORDER = A.NO_WO AND FLAGP <> 'M' ORDER BY TANGGAL, START)), ((SELECT TOP 1 TANGGAL+' '+PAUSE FROM TEMPPROGRESS WHERE NOORDER = A.NO_WO AND FLAGP = 'P' ORDER BY TANGGAL DESC, START DESC))),0) AS PRODUKSI , " & _
           "ISNULL(DATEDIFF(s, ((SELECT TOP 1 TANGGAL+' '+JAM FROM PROGRESS WHERE NOORDER = A.NO_WO AND LEFT(KETERANGAN, CHARINDEX(' |',KETERANGAN))= 'Final Check' ORDER BY TANGGAL DESC, JAM DESC)), ((SELECT TOP 1 TANGGAL+' '+PAUSE FROM TEMPPROGRESS WHERE NOORDER = A.NO_WO AND FLAGP = 'F' ORDER BY TANGGAL DESC, START DESC))),0) AS FINALCHECK , " & _
           "ISNULL(DATEDIFF(s, ((SELECT TOP 1 TANGGAL FROM VALIDORDER WHERE NO_WO = A.NO_WO AND USERDEL = '')), (E.INVDATE+' '+INVTIME)),0) AS INVOICE, " & _
           "CASE TglKasir wHEN   '01/01/1900' tHEN 0 eLSE ISNULL(DATEDIFF(s, CONVERT(datetime,CONVERT(varchar(16),F.TGLPANGGIL, 20)),TglKasir),0) eND AS Registrasi, " & _
           "a.NoFak " & _
           "FROM JOBDTA A LEFT JOIN LISTANTRIAN B ON A.NO_WO = B.WO " & _
           "LEFT JOIN " & Log.LogQueue & ".dbo.QUEUE C ON B.IDPRINT = C.ID " & _
           "LEFT JOIN " & Log.LogQueue & ".dbo.PANGGILAN D ON B.IDCALL = D.ID " & _
           "LEFT JOIN T_INVOICE E ON A.NO_WO = E.NO_WO AND A.NOINV = LEFT(E.NOINV,8) AND E.NOINV NOT LIKE '%-B' " & _
           "LEFT JOIN " & Log.LogQueue & ".dbo.FINISHSERVICE F ON A.NO_POLISI = F.NOPOL AND A.INVDATE = F.TGLFINISH AND F.PANGGIL <> 0 " & _
           "WHERE A.DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND A.FS_LOG_DEL = '' AND A.PDS = 0 AND A.TPSVC <> 'W' AND A.THS = 0 AND A.FN_RETURN = 0 ORDER BY A.NO_WO"

    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.TxtAlamat.SetText gAlamat
    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    nWait.Visible = False
    DoEvents
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub


Private Sub cmdICSK_Click()
    PicBars.Visible = True
    nBars.value = 0
    nBars.Max = 21
    nLabs.Caption = ""
    DoEvents
    
    nWait.Visible = True
    DoEvents
    Dim nThn As Double
    Dim nBln As Double
    nThn = Val(cboThn.Text)
    nBln = Val(cboBln.Text)
    
    Dim nIC As Double
    For nIC = 1 To 12
        Call objR.FetchICSK_Bulanan(nThn, nIC)
        nBars.value = nBars.value + 1
        nLabs.Caption = "Get Data Master ICSK (" & nIC & "/12)"
        DoEvents
    Next nIC
    
    nBars.value = nBars.value + 1
    nLabs.Caption = "Get Data ICSK ..."
    DoEvents
    Dim rsI As ADODB.Recordset
    Set rsI = objR.FetchICSK(nThn, nBln)
    If rsI.RecordCount > 0 Then
        txtPart = (rsI.Fields(0).value)
        txtDiss = (rsI.Fields(1).value)
        txtSats = (rsI.Fields(2).value)
    End If

        nBars.value = nBars.value + 1
        nLabs.Caption = "Get Data Fix It Right ..."
        DoEvents
        Dim nFir As Double
        nFir = HitungFIR(nThn, nBln)

        nBars.value = nBars.value + 1
        nLabs.Caption = "Get Data Short Lead Time ..."
        DoEvents
        Dim nSLT_Jam As Double
        Dim nSLT_Min As Double
        Dim rsSLT As ADODB.Recordset
        Set rsSLT = objR.FetchSLT_ICSK(nThn, nBln)
        If rsSLT.RecordCount = 0 Then
            nSLT_Jam = 0
            nSLT_Min = 0
        Else
            nSLT_Jam = rsSLT.Fields(1).value
            nSLT_Min = rsSLT.Fields(2).value
        End If

        nBars.value = nBars.value + 1
        nLabs.Caption = "Get Data Short Lead Time (Process And Stagnasi) ..."
        DoEvents
        Dim nSLT_JamT As Double
        Dim nSLT_MinT As Double
        Dim nSLT_JamP As Double
        Dim nSLT_MinP As Double
        Dim rsLTs As ADODB.Recordset
        Set rsLTs = objR.FetchSLT_SPLIT_ICSK(nThn, nBln)
        If rsLTs.RecordCount = 0 Then
            nSLT_JamT = 0
            nSLT_MinT = 0
            nSLT_JamP = 0
            nSLT_MinP = 0
        Else
            nSLT_JamT = rsLTs.Fields(1).value
            nSLT_MinT = rsLTs.Fields(2).value
            nSLT_JamP = rsLTs.Fields(4).value
            nSLT_MinP = rsLTs.Fields(5).value
        End If

        nBars.value = nBars.value + 1
        nLabs.Caption = "Get Data On Time Delivery ..."
        DoEvents
        'Dim nTD As Double
        'Dim nUEt As Double
        'Dim nTht As Double
        Dim nOTDt As Double
        'Dim rsOt As ADodb.Recordset
        'Set rsOt = objR.FetchOTD_Notif_ICSK(nThn, nBln)
        'If rsOt.RecordCount = 0 Then
        '    nTD = 0
        'Else
        '    nTD = rsOt.Fields(0).value
        'End If

        nBars.value = nBars.value + 1
        'nLabs.Caption = "Get Data Unit Entry ..."
        DoEvents
        'nUEt = 0
        'Set rsOt = objO.FetchData("SELECT COUNT(*) AS JUMLAH FROM T_INVOICE WHERE MONTH(DATE) = '" & nBln & "' AND YEAR(DATE) = '" & nThn & "' AND DELIV <> 'BTL' AND NOINV NOT LIKE '%-B' AND FS_LOG_DEL = ''")
        'If Not rsOt.EOF Then
        '    nUEt = CDbl(rsOt("JUMLAH"))
        'End If
        'nTht = 0
        'Set rsOt = objR.FetchOTD_THS_ICSK(nThn, nBln)
        'If Not rsOt.EOF Then
        '    nTht = CDbl(rsOt("THS"))
        'End If
        'nOTDt = Round((nTD / (nUEt - nTht)) * 100, 1)

        
        nOTDt = objO.FetchData_Double("select (((Select Count(A.Slisih) from " & _
                        "(SELECT CASE WHEN DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), " & _
                        "ISNULL(TGLPANGGIL,'')) <= 0 AND DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), ISNULL(TGLPANGGIL,'')) >= -30 THEN 1 ELSE 0 END AS SLISIH " & _
                        "FROM JOBDTA A " & _
                        "INNER JOIN JOBDTL B ON A.NO_WO = B.NO_WO " & _
                        "LEFT JOIN " & Log.LogQueue & "..FINISHSERVICE C ON A.NO_POLISI = C.NOPOL AND A.INVDATE = C.TGLFINISH AND PANGGIL <> 0 " & _
                        "WHERE NOINV <> '' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND PDS = 0 AND THS = 0 AND TPSVC <> 'W' AND " & _
                        "MONTH(DATE) = " & nBln & " AND YEAR(DATE) = " & nThn & ") A Where A.Slisih = 1)*100) " & _
                        "/(Select Count(A.Slisih) from " & _
                        "(SELECT CASE WHEN DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), " & _
                        "ISNULL(TGLPANGGIL,'')) <= 0 AND DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), ISNULL(TGLPANGGIL,'')) >= -30 THEN 1 ELSE 0 END AS SLISIH " & _
                        "FROM JOBDTA A " & _
                        "INNER JOIN JOBDTL B ON A.NO_WO = B.NO_WO " & _
                        "LEFT JOIN " & Log.LogQueue & "..FINISHSERVICE C ON A.NO_POLISI = C.NOPOL AND A.INVDATE = C.TGLFINISH AND PANGGIL <> 0 " & _
                        "WHERE NOINV <> '' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND PDS = 0 AND THS = 0 AND TPSVC <> 'W' AND " & _
                        "MONTH(DATE) = " & nBln & " AND YEAR(DATE) = " & nThn & ") A)) as OTD")
        


        nBars.value = nBars.value + 1
        nLabs.Caption = "Create Report ..."
        DoEvents
        If rsI.RecordCount = 0 Then
            Call objR.InsertItemICSK(nThn, nBln, nFir, nSLT_Jam, nSLT_Min, nSLT_JamT, nSLT_MinT, nSLT_JamP, nSLT_MinP, (txtPart.Text), nOTDt, (txtDiss.Text), (txtSats.Text))
        Else
            Call objR.UpdateItemICSK(nThn, nBln, nFir, nSLT_Jam, nSLT_Min, nSLT_JamT, nSLT_MinT, nSLT_JamP, nSLT_MinP, (txtPart.Text), nOTDt, (txtDiss.Text), (txtSats.Text))
        End If
  
    'If Log.LogUser = "" Then
        cmdSave.Enabled = True
    'End If
    nBars.value = nBars.value + 1
    nLabs.Caption = "Create Report ..."
    cmdPrint.Enabled = True
    nWait.Visible = False
    DoEvents
    MsgBox "Laporan siap di Cetak !", vbInformation + vbOKOnly, "Laporan ICSK HO"
    PicBars.Visible = False
End Sub

Private Function HitungFIR(xThn As Double, xBln As Double) As Double
    Dim nDiFollow As Double
    nDiFollow = objR.FetchDiFollow_FIR(xThn, xBln)
    HitungFIR = Round(1 - (Round(objR.FetchNoQ_FIR(xThn, xBln, 1) / nDiFollow, 2) + Round(objR.FetchNoQ_FIR(xThn, xBln, 2) / nDiFollow, 2) + Round(objR.FetchNoQ_FIR(xThn, xBln, 3) / nDiFollow, 2)), 2) * 100
End Function


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

