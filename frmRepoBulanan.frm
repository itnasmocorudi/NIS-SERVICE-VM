VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoBulanan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Tag             =   "  "
   Begin ucXPButton.XPButton XPButton2 
      Height          =   300
      Left            =   5370
      TabIndex        =   92
      Top             =   8610
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
      MICON           =   "frmRepoBulanan.frx":0000
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
      Left            =   6480
      TabIndex        =   91
      Top             =   8625
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
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   8595
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
      Left            =   4155
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   8595
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
      Left            =   2670
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   8595
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   8595
      Width           =   1050
   End
   Begin VB.Frame framePenjualanPartCounter 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   480
      TabIndex        =   77
      Top             =   600
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CheckBox ckFakturA 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A"
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
         Left            =   2160
         TabIndex        =   82
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox ckFakturB 
         BackColor       =   &H00FFFFFF&
         Caption         =   "B"
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
         Left            =   3000
         TabIndex        =   81
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A"
         Height          =   375
         Left            =   -4920
         TabIndex        =   80
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox ckFakturC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C"
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
         Left            =   3840
         TabIndex        =   79
         Top             =   720
         Width           =   615
      End
      Begin ucXPButton.XPButton cmdPilihRepo 
         Height          =   375
         Left            =   5160
         TabIndex        =   78
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Proses"
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
         BCOL            =   65280
         BCOLO           =   65280
         FCOL            =   16777215
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRepoBulanan.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Ket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: PENJUALAN PART COUNTER  ::."
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
         Index           =   10
         Left            =   1560
         TabIndex        =   85
         Top             =   240
         Width           =   3555
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   7215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NOFAKTUR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Ket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: PENJUALAN PART COUNTER  ::."
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
         Index           =   9
         Left            =   1665
         TabIndex        =   83
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.PictureBox picBarByModel 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7920
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   64
      Top             =   8280
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   5295
         TabIndex        =   65
         Top             =   480
         Width           =   5295
         Begin VB.Label nKetByModel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   66
            Top             =   0
            Width           =   5295
         End
      End
      Begin MSComctlLib.ProgressBar nBarByModel 
         Height          =   495
         Left            =   0
         TabIndex        =   67
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nProByModel 
         Alignment       =   1  'Right Justify
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
         Left            =   5160
         TabIndex        =   69
         Top             =   0
         Width           =   1785
      End
      Begin VB.Label nLabByModel 
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
         TabIndex        =   68
         Top             =   0
         Width           =   4935
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   43
      Top             =   8280
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
   Begin VB.ComboBox cboThnP 
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
      Left            =   10080
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   5475
      Width           =   1095
   End
   Begin VB.ComboBox cboBlnP 
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
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   5475
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoBulanan.frx":0038
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   52
      Top             =   0
      Width           =   450
   End
   Begin VB.PictureBox PicPS 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7920
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   44
      Top             =   3960
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.ProgressBar BarPS 
         Height          =   495
         Left            =   0
         TabIndex        =   45
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nPS 
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
         TabIndex        =   46
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox PicNEA 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7920
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   40
      Top             =   1800
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.ProgressBar BarNEA 
         Height          =   495
         Left            =   0
         TabIndex        =   41
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nNEA 
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
         TabIndex        =   42
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox PicBars 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.ProgressBar nBars 
         Height          =   495
         Left            =   0
         TabIndex        =   22
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
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
         TabIndex        =   23
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.ComboBox cboBulan1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmRepoBulanan.frx":2140
      Left            =   1920
      List            =   "frmRepoBulanan.frx":2142
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1155
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   405
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   1155
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin VB.PictureBox PicBar 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   5295
         TabIndex        =   5
         Top             =   480
         Width           =   5295
         Begin VB.Label nKet 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   6
            Top             =   0
            Width           =   5295
         End
      End
      Begin MSComctlLib.ProgressBar nBar 
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nLab 
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
         TabIndex        =   9
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label nPro 
         Alignment       =   1  'Right Justify
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
         Left            =   5160
         TabIndex        =   8
         Top             =   0
         Width           =   1785
      End
   End
   Begin VB.ComboBox cboBulan2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   405
      Index           =   1
      Left            =   3720
      TabIndex        =   3
      Top             =   1155
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin ucXPButton.XPButton TblSimpan 
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   1080
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
      MICON           =   "frmRepoBulanan.frx":2144
      PICN            =   "frmRepoBulanan.frx":2160
      PICH            =   "frmRepoBulanan.frx":25B7
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
      Left            =   3960
      TabIndex        =   15
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
      MICON           =   "frmRepoBulanan.frx":2AA1
      PICN            =   "frmRepoBulanan.frx":2ABD
      PICH            =   "frmRepoBulanan.frx":2F14
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3315
      Width           =   855
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
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3315
      Width           =   1095
   End
   Begin ucXPButton.XPButton cmdPrintS 
      Height          =   495
      Left            =   5520
      TabIndex        =   28
      Top             =   5400
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
      MICON           =   "frmRepoBulanan.frx":33FE
      PICN            =   "frmRepoBulanan.frx":341A
      PICH            =   "frmRepoBulanan.frx":3871
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox PicBarEM 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.ProgressBar nBarEM 
         Height          =   495
         Left            =   0
         TabIndex        =   26
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nLabEM 
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
         TabIndex        =   27
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.ComboBox cboBlnS 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   5475
      Width           =   855
   End
   Begin VB.ComboBox cboThnS 
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
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   5475
      Width           =   1095
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   33
      Top             =   5520
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   34
      Top             =   5520
      Width           =   615
   End
   Begin ucXPButton.XPButton cmdPrintNEA 
      Height          =   495
      Left            =   13560
      TabIndex        =   35
      Top             =   1080
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
      MICON           =   "frmRepoBulanan.frx":3D5B
      PICN            =   "frmRepoBulanan.frx":3D77
      PICH            =   "frmRepoBulanan.frx":41CE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBDate6Ctl.TDBDate dt0 
      Height          =   315
      Left            =   9480
      TabIndex        =   37
      Top             =   1155
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   556
      Calendar        =   "frmRepoBulanan.frx":46B8
      Caption         =   "frmRepoBulanan.frx":47E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoBulanan.frx":4840
      Keys            =   "frmRepoBulanan.frx":485E
      Spin            =   "frmRepoBulanan.frx":48BC
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      ForeColor       =   0
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   2
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
      ValueVT         =   6815745
      Value           =   39183
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate dt1 
      Height          =   315
      Left            =   11520
      TabIndex        =   38
      Top             =   1155
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   556
      Calendar        =   "frmRepoBulanan.frx":48E4
      Caption         =   "frmRepoBulanan.frx":4A10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoBulanan.frx":4A6C
      Keys            =   "frmRepoBulanan.frx":4A8A
      Spin            =   "frmRepoBulanan.frx":4AE8
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      ForeColor       =   0
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   2
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
      ValueVT         =   6815745
      Value           =   39183
      CenturyMode     =   0
   End
   Begin ucXPButton.XPButton cmdPrintPS 
      Height          =   495
      Left            =   13560
      TabIndex        =   47
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
      MICON           =   "frmRepoBulanan.frx":4B10
      PICN            =   "frmRepoBulanan.frx":4B2C
      PICH            =   "frmRepoBulanan.frx":4F83
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBDate6Ctl.TDBDate dtP0 
      Height          =   315
      Left            =   9480
      TabIndex        =   48
      Top             =   3315
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   556
      Calendar        =   "frmRepoBulanan.frx":546D
      Caption         =   "frmRepoBulanan.frx":5599
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoBulanan.frx":55F5
      Keys            =   "frmRepoBulanan.frx":5613
      Spin            =   "frmRepoBulanan.frx":5671
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      ForeColor       =   0
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   2
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
      ValueVT         =   6815745
      Value           =   39183
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate dtP1 
      Height          =   315
      Left            =   11520
      TabIndex        =   49
      Top             =   3315
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   556
      Calendar        =   "frmRepoBulanan.frx":5699
      Caption         =   "frmRepoBulanan.frx":57C5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoBulanan.frx":5821
      Keys            =   "frmRepoBulanan.frx":583F
      Spin            =   "frmRepoBulanan.frx":589D
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      ForeColor       =   0
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   2
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
      ValueVT         =   6815745
      Value           =   39183
      CenturyMode     =   0
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
      BackColor       =   12577002
   End
   Begin ucXPButton.XPButton cmdPrintP 
      Height          =   495
      Left            =   11280
      TabIndex        =   53
      Top             =   5400
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
      MICON           =   "frmRepoBulanan.frx":58C5
      PICN            =   "frmRepoBulanan.frx":58E1
      PICH            =   "frmRepoBulanan.frx":5D38
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdView 
      Height          =   495
      Left            =   3600
      TabIndex        =   60
      Top             =   7560
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&v"
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
      MICON           =   "frmRepoBulanan.frx":6222
      PICN            =   "frmRepoBulanan.frx":623E
      PICH            =   "frmRepoBulanan.frx":67AC
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
      Left            =   13680
      TabIndex        =   62
      Top             =   9240
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
      MICON           =   "frmRepoBulanan.frx":6DA0
      PICN            =   "frmRepoBulanan.frx":6DBC
      PICH            =   "frmRepoBulanan.frx":7290
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   405
      Index           =   2
      Left            =   11040
      TabIndex        =   63
      Top             =   7635
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   405
      Index           =   3
      Left            =   11040
      TabIndex        =   71
      Top             =   7635
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   495
      Left            =   13080
      TabIndex        =   72
      Top             =   7560
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
      MICON           =   "frmRepoBulanan.frx":77DB
      PICN            =   "frmRepoBulanan.frx":77F7
      PICH            =   "frmRepoBulanan.frx":7C4E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboBulan3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   7635
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":8138
            Key             =   "cmdfirst"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":88B2
            Key             =   "cmdprevious"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":902C
            Key             =   "cmdnext"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":97A6
            Key             =   "cmdlast"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":9F20
            Key             =   "cmdadd"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":A61A
            Key             =   "cmdquit2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":AA6C
            Key             =   "cmdquit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":B1E6
            Key             =   "cmddelete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":B8E0
            Key             =   "cmdcopy"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":C05A
            Key             =   "cmdedit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":C754
            Key             =   "cmdquit1"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":D42E
            Key             =   "cmdcancel1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":DB28
            Key             =   "cmdcancel"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":E222
            Key             =   "cmdsave"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepoBulanan.frx":E99C
            Key             =   "cmdprint"
         EndProperty
      EndProperty
   End
   Begin ucXPButton.XPButton cmdPrintLaporanByModel 
      Height          =   495
      Left            =   11880
      TabIndex        =   86
      Top             =   7560
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
      MICON           =   "frmRepoBulanan.frx":FC4E
      PICN            =   "frmRepoBulanan.frx":FC6A
      PICH            =   "frmRepoBulanan.frx":100C1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: LAPORAN SERVICE GENERAL REPAIR - UNIT BY MODEL  ::."
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
      Index           =   8
      Left            =   7695
      TabIndex        =   76
      Top             =   7155
      Width           =   7395
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   7215
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   7800
      Top             =   7440
      Width           =   7215
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  K P I   LAPORAN  BULANAN  ::."
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
      Index           =   7
      Left            =   9720
      TabIndex        =   75
      Top             =   7140
      Width           =   3525
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
      Index           =   4
      Left            =   8160
      TabIndex        =   74
      Top             =   7680
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
      Index           =   3
      Left            =   9120
      TabIndex        =   73
      Top             =   7680
      Width           =   60
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      X1              =   15120
      X2              =   360
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  VIEW  PROSES  WARANTY  ::."
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
      Index           =   6
      Left            =   2400
      TabIndex        =   61
      Top             =   7140
      Width           =   3435
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      Top             =   7440
      Width           =   7215
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   7800
      Top             =   5280
      Width           =   7215
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  LAPORAN  BULANAN  PART  ::."
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
      Index           =   5
      Left            =   9720
      TabIndex        =   59
      Top             =   4980
      Width           =   3525
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
      Index           =   2
      Left            =   8160
      TabIndex        =   58
      Top             =   5520
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
      Index           =   2
      Left            =   9120
      TabIndex        =   57
      Top             =   5520
      Width           =   60
   End
   Begin VB.Label nWaitPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ... (^_^)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12720
      TabIndex        =   56
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8415
      Left            =   0
      Picture         =   "frmRepoBulanan.frx":105AB
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   375
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: LAPORAN  PROFIT  SHARING  ::."
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
      Index           =   4
      Left            =   9660
      TabIndex        =   51
      Top             =   2820
      Width           =   3645
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   7800
      Top             =   3120
      Width           =   7215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Range Invoice :  "
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
      Left            =   7920
      TabIndex        =   50
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   11280
      X2              =   11400
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   11280
      X2              =   11400
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Range Invoice :  "
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
      Left            =   7920
      TabIndex        =   39
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   7800
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: LAPORAN  N E A  ::."
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
      Index           =   3
      Left            =   10320
      TabIndex        =   36
      Top             =   660
      Width           =   2325
   End
   Begin VB.Label nWait 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ... (^_^)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
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
      Index           =   0
      Left            =   1800
      TabIndex        =   18
      Top             =   3360
      Width           =   60
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
      Index           =   0
      Left            =   840
      TabIndex        =   19
      Top             =   3360
      Width           =   840
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  LAPORAN  BULANAN  DYNA  ::."
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
      Left            =   2370
      TabIndex        =   14
      Top             =   2820
      Width           =   3585
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
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   60
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
      Left            =   840
      TabIndex        =   11
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  K P I   LAPORAN  BULANAN  ::."
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
      Left            =   2400
      TabIndex        =   10
      Top             =   660
      Width           =   3525
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      Top             =   960
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   7215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      Top             =   3120
      Width           =   7215
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  LAPORAN  BULANAN  SB - NP / EM  ::."
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
      Left            =   1995
      TabIndex        =   24
      Top             =   4980
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   7215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      Top             =   5280
      Width           =   7215
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
      Index           =   1
      Left            =   840
      TabIndex        =   32
      Top             =   5520
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
      Index           =   1
      Left            =   1800
      TabIndex        =   31
      Top             =   5520
      Width           =   60
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000080FF&
      Height          =   375
      Left            =   3960
      Top             =   5460
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   7215
   End
End
Attribute VB_Name = "frmRepoBulanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rsTemp As ADODB.Recordset
Dim vNamaTabel As String
Dim objBilling As New BClsBilling
Private obj As New ClsKPI
Private objKPI As New ClsKPICEK
Private objR As New clsRepo
Private rsNEA As New ADODB.Recordset
Private rsPS As New ADODB.Recordset
Dim Report As Object

Private noInv  As String
Private nUnit  As Double
Private nUnitR As Double
Private nINV   As Double
Private nInvR  As Double

'penambahan variable untuk total unit per model sbe
Private nTotSBE1, nTotSBE2, nTotSBE3, nTotSBE4, nTotSBE5, nTotSBE6, nTotSBE7, nTotSBE8, nTotSBE9, nTotSBE10, nTotSBE11, nTotSBE12, nTotSBE13, nTotSBE14, nTotSBE15, nTotSBE16, nTotSBE17, nTotSBE18, nTotSBE19, nTotSBE20, nTotSBE21 As Double
Private nTotGR1, nTotGR2, nTotGR3, nTotGR4, nTotGR5, nTotGR6, nTotGR7, nTotGR8, nTotGR9, nTotGR10, nTotGR11, nTotGR12, nTotGR13, nTotGR14, nTotGR15, nTotGR16, nTotGR17, nTotGR18, nTotGR19, nTotGR20, nTotGR21 As Double
Private totCPUS1, totCPUS2, totCPUS3, totCPUS4, totCPUS5, totCPUS6, totCPUS7, totCPUS8, totCPUS9, totCPUS10, totCPUS11, totCPUS12, totCPUS13, totCPUS14, totCPUS15, totCPUS16, totCPUS17, totCPUS18, totCPUS19, totCPUS20, totCPUS21 As Double
Private nonCPUS1, nonCPUS2, nonCPUS3, nonCPUS4, nonCPUS5, nonCPUS6, nonCPUS7, nonCPUS8, nonCPUS9, nonCPUS10, nonCPUS11, nonCPUS12, nonCPUS13, nonCPUS14, nonCPUS15, nonCPUS16, nonCPUS17, nonCPUS18, nonCPUS19, nonCPUS20, nonCPUS21 As Double
Private totUE1, totUE2, totUE3, totUE4, totUE5, totUE6, totUE7, totUE8, totUE9, totUE10, totUE11, totUE12, totUE13, totUE14, totUE15, totUE16, totUE17, totUE18, totUE19, totUE20, totUE21 As Double
Private labor1, labor2, labor3, labor4, labor5, labor6, labor7, labor8, labor9, labor10, labor11, labor12, labor13, labor14, labor15, labor16, labor17, labor18, labor19, labor20, labor21 As Double
Private keteranganFaktur As String

Dim totalLead2Sales     As Double
Dim totalDOSAtoSales    As Double
Dim totalReqBookBySales As Double
Dim totalBookBySales    As Double
Dim totalReminderSales  As Double
Dim totalReminderTerhubungSales As Double
Dim totalInvoiceSales   As Double
Dim totalWOBPtoGRP      As Double
Dim tglAwal             As String
Dim tglAkhir            As String


Private Sub hapusForm()
    For i = 0 To 1
        Text(i) = 0
    Next
    
    cboBulan1.Clear
    cboBulan2.Clear
    For i = 1 To 12
        cboBulan1.AddItem Format(i, "00") & " | " & GetNamaBulan(i)
        cboBulan2.AddItem Format(i, "00") & " | " & GetNamaBulan(i)
        cboBulan3.AddItem Format(i, "00") & " | " & GetNamaBulan(i)
    Next
    cboBulan1 = Format(Month(isnow), "00") & " | " & GetNamaBulan(Month(isnow))
    cboBulan2 = Format(Month(isnow), "00") & " | " & GetNamaBulan(Month(isnow))
    cboBulan3 = Format(Month(isnow), "00") & " | " & GetNamaBulan(Month(isnow))
    Text(0) = year(isnow)
    Text(1) = year(isnow)
    Text(2) = year(isnow)
    
    cboThn.Clear
    cboThnS.Clear
    cboThnP.Clear
    cboThnP.AddItem "2012"
    For i = 2013 To year(isnow)
        cboThn.AddItem i
        cboThnS.AddItem i
        cboThnP.AddItem i
        If i <> 2013 Then
           ' cboThnSR.AddItem i
        End If
    Next i
    cboThn.Text = year(isnow)
    cboThnS.Text = year(isnow)
    cboThnP.Text = year(isnow)
'    cboThnSR.Text = Year(isnow)
    cboBln.Clear
    cboBlnS.Clear
    For i = 1 To 12
        cboBln.AddItem i
        cboBlnS.AddItem i
        cboBlnP.AddItem i
'        cboBlnSR.AddItem i
    Next i
    cboBln.Text = Month(isnow)
    cboBlnS.Text = Month(isnow)
    cboBlnP.Text = Month(isnow)
'    cboBlnSR.Text = Month(isnow)
    

    dt0.Text = Replace(TglDDMMYYYY("01/" + Format(Month(isnow), "00") + "/" + Format(year(isnow), "0000")), "-", "/")
    dt1.Text = Replace(GetTglAkhir(Month(isnow), year(isnow), 0), "-", "/")

    dtP0.Text = Replace(TglDDMMYYYY("01/" + Format(Month(isnow), "00") + "/" + Format(year(isnow), "0000")), "-", "/")
    dtP1.Text = Replace(GetTglAkhir(Month(isnow), year(isnow), 0), "-", "/")

    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton2.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
    
End Sub

Private Function cekInsertByModel(nID As Integer, nKet As String) As Boolean
Dim rsHlp As ADODB.Recordset
Dim hasil As Boolean
    hasil = False
    'cekInsert = False
    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE urut = '" & nID & "' AND KETERANGAN = '" & nKet & "'")
    If rsHlp.EOF Then
        'cekInsert = True
        hasil = True
    End If
    cekInsertByModel = hasil
End Function

Private Function cekNilaiKolom(nID As Integer, nKet As String, model As String) As Double
    Dim rsHlp As ADODB.Recordset
    Dim hasil As Double
    hasil = 0
    Set rsHlp = objO.FetchData("SELECT " & model & " FROM [TEMPDB]..[" & vNamaTabel & "] WHERE urut = '" & nID & "' AND KETERANGAN = '" & nKet & "'")
    If rsHlp.EOF Then
        hasil = CDbl(rsHlp.Fields(0).value)
    End If
    cekNilaiKolom = hasil
End Function

Private Function cekInsert(nID As Integer, nKet As String) As Boolean
    Dim rsHlp As ADODB.Recordset

    cekInsert = False
    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE urut = '" & nID & "' AND KETERANGAN = '" & nKet & "'")
    If rsHlp.EOF Then
        cekInsert = True
    End If
    
End Function

Private Function kpi6_b(Bln As Integer, Thn As Integer) As Integer
    Dim rs2 As ADODB.Recordset
    Dim P2 As Integer
    
    On Error Resume Next
    kpi6_b = 0
    'Set rs2 = obj.KPI_6(False, Bln, Thn)
    Set rs2 = objKPI.KPI_6(False, Bln, Thn)
    If rs2.RecordCount <> 0 Then P2 = rs2.Fields(0).value
    'Set rs2 = obj.KPI_6b(True, Bln, Thn)
    Set rs2 = objKPI.KPI_6b(True, Bln, Thn)
    If rs2.RecordCount <> 0 Then P2 = P2 + CInt(rs2.Fields(0).value)
    
    kpi6_b = CDbl(P2)
End Function

Private Function kpi6_a(Bln As Integer, Thn As Integer) As Integer
    Dim rs2 As ADODB.Recordset
    Dim P1 As Integer
    
    On Error Resume Next
    kpi6_a = 0
    
    'Set rs1 = obj.KPI_6(True, Bln, Thn)
    Set rs1 = objKPI.KPI_6(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then P1 = rs1.Fields(0).value
    'Set rs1 = obj.KPI_6a(True, Bln, Thn)
    Set rs1 = objKPI.KPI_6a(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then P1 = P1 + CInt(rs1.Fields(0).value)
    
    kpi6_a = CDbl(P1)
End Function

Private Function kpi9_a(Bln As Integer, Thn As Integer) As Integer
    Dim rs2 As ADODB.Recordset
    Dim Pda As Integer
    
    On Error Resume Next
    kpi9_a = 0
    'Set rs1 = obj.KPI_9(True, Bln, Thn)
    Set rs1 = objKPI.KPI_9(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then Pda = rs1.Fields(0).value
    ''Set rs1 = obj.KPI_9(False, Bln, Thn)
    Set rs1 = objKPI.KPI_9(False, Bln, Thn)
    If rs1.RecordCount <> 0 Then Pda = Pda + rs1.Fields(0).value
    
    kpi9_a = CDbl(Pda)
End Function

Private Function insertTabel(nBulan As String, nBaris As Integer, nKet As String, nJml As Double, nkode As Integer, Optional diBold As Boolean = False) As String
    If nkode = 1 Then
        insertTabel = "INSERT INTO [TEMPDB]..[" & vNamaTabel & "] (URUT, KETERANGAN, BOLD, " & nBulan & ")  VALUES (" & nBaris & ",'" & nKet & "','" & IIf(diBold, 1, 0) & "'," & Replace(CDbl(nJml), ",", ".") & ")"
    ElseIf nkode = 2 Then
        insertTabel = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET " & nBulan & " = " & Replace(CDbl(nJml), ",", ".") & " WHERE URUT = " & nBaris & " AND KETERANGAN = '" & nKet & "'"
    ElseIf nkode = 3 Then
        insertTabel = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET " & nBulan & " = " & nBulan & " + " & Replace(CDbl(nJml), ",", ".") & " WHERE URUT = " & nBaris & " AND KETERANGAN = '" & nKet & "'"
    End If
    
End Function


Private Function insertTabelByModel(model As String, nBulan As String, nBaris As Integer, nKet As String, nJml As Double, nkode As Integer, Optional diBold As Boolean = False) As String
    If nkode = 1 Then
        insertTabelByModel = "INSERT INTO [TEMPDB]..[" & vNamaTabel & "] (URUT, KETERANGAN, BOLD, " & nBulan & "," & model & ")  VALUES (" & nBaris & ",'" & nKet & "','" & IIf(diBold, 1, 0) & "'," & Replace(CDbl(nJml), ",", ".") & "," & Replace(CDbl(nJml), ",", ".") & ")"
    ElseIf nkode = 2 Then
        insertTabelByModel = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET " & model & " = " & Replace(CDbl(nJml), ",", ".") & " , " & nBulan & " = " & Replace(CDbl(nJml), ",", ".") & " WHERE URUT = " & nBaris & " AND KETERANGAN = '" & nKet & "'"
    ElseIf nkode = 3 Then
        insertTabelByModel = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET " & nBulan & " = " & nBulan & " + " & Replace(CDbl(nJml), ",", ".") & " WHERE URUT = " & nBaris & " AND KETERANGAN = '" & nKet & "'"
    End If
    
End Function


Private Sub cboBulan1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cboBulan2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cmdPilihRepo_Click()
'    If ckFakturA.value = vbChecked Or ckFakturB.value = vbChecked Or ckFakturC.value = vbChecked Then
        cmdPilihRepo.Enabled = False
        ckFakturA.Enabled = False
        ckFakturB.Enabled = False
        ckFakturC.Enabled = False
        Screen.MousePointer = vbHourglass
        
        Call cetakLaporanBulanan
'
'    Else
'        MsgBox "Salah satu nofaktur harus dipilih.", vbInformation
'        Exit Sub
'    End If
End Sub

Private Sub cmdPrint_Click()
    nWait.Visible = True

        '2 April 2013
        nBars.Min = 0
        nBars.Max = 20
        nBars.value = 0
        nLabs.Caption = "0 Of 20  Get Data Report"
        PicBars.Visible = True
        DoEvents


    DoEvents
    
    Dim nThn As Double
    Dim nBln As Double
    nThn = Val(cboThn.Text)
    nBln = Val(cboBln.Text)
    
    Dim nIC As Double
    For nIC = 1 To 12
        Call objR.FetchDYNA_Bulanan(nThn, nIC)
        nLabs.Caption = nIC & " Of 20  Cex Existing Data"
        nBars.value = nIC
        DoEvents
    Next nIC
    
    Dim jSPK As Double
        jSPK = objR.FetchSPK_DYNA(nThn, nBln)
        nLabs.Caption = "13 Of 20  Penjualan (Retail SPK)"
        nBars.value = 13
        DoEvents
    Dim jSB1 As Double
        jSB1 = objR.FetchSB_DYNA(nThn, nBln, "R001")
        nLabs.Caption = "14 Of 20  Service 1.000 Km"
        nBars.value = 14
        DoEvents
    Dim jSB5 As Double
        jSB5 = objR.FetchSB_DYNA(nThn, nBln, "R005")
        nLabs.Caption = "15 Of 20  Service 5.000 Km"
        nBars.value = 15
        DoEvents
    Dim jSBE As Double
        jSBE = objR.FetchSBE_DYNA(nThn, nBln)
        nLabs.Caption = "16 Of 20  Service Berkala External"
        nBars.value = 16
        DoEvents
    Dim jGRP As Double
        jGRP = objR.FetchSB_DYNA(nThn, nBln, "")
        nLabs.Caption = "17 Of 20  General Repair"
        nBars.value = 17
        DoEvents
    Dim jTHS As Double
        jTHS = objR.FetchTHS_DYNA(nThn, nBln)
        nLabs.Caption = "18 Of 20  Service - Dealer n Visit"
        nBars.value = 18
        DoEvents
    Dim jDLR As Double
        jDLR = (jSB1 + jSB5 + jSBE + jGRP) - jTHS
        Call objR.UpdateItemDYNA(nThn, nBln, jSPK, jSB1, jSB5, jSBE, jGRP, jDLR, jTHS)
        nLabs.Caption = "19 Of 20  Service - Dealer n Visit"
        nBars.value = 19
        DoEvents
    Call Print_DYNA
End Sub

Private Sub Print_DYNA()
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New RptDYNA_HO

    sSQL = "Select A.*,B.*,C.*,D.*,E.*,F.*,G.*,H.*,I.*,J.*,K.*,L.* from " & _
           "(select Tahun as Tahun1,SPK as SPK1,SB1 as SB11,SB5 as SB51,SBE as SBE1,GRP as GRP1,SVCD as SVCD1,SVCV as SVCV1 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 1) A Left Join " & _
           "(select Tahun as Tahun2,SPK as SPK2,SB1 as SB12,SB5 as SB52,SBE as SBE2,GRP as GRP2,SVCD as SVCD2,SVCV as SVCV2 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 2) B On A.Tahun1  = B.Tahun2 Left Join " & _
           "(select Tahun as Tahun3,SPK as SPK3,SB1 as SB13,SB5 as SB53,SBE as SBE3,GRP as GRP3,SVCD as SVCD3,SVCV as SVCV3 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 3) C On A.Tahun1  = C.Tahun3 Left Join " & _
           "(select Tahun as Tahun4,SPK as SPK4,SB1 as SB14,SB5 as SB54,SBE as SBE4,GRP as GRP4,SVCD as SVCD4,SVCV as SVCV4 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 4) D On A.Tahun1  = D.Tahun4 Left Join " & _
           "(select Tahun as Tahun5,SPK as SPK5,SB1 as SB15,SB5 as SB55,SBE as SBE5,GRP as GRP5,SVCD as SVCD5,SVCV as SVCV5 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 5) E On A.Tahun1  = E.Tahun5 Left Join " & _
           "(select Tahun as Tahun6,SPK as SPK6,SB1 as SB16,SB5 as SB56,SBE as SBE6,GRP as GRP6,SVCD as SVCD6,SVCV as SVCV6 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 6) F On A.Tahun1  = F.Tahun6 Left Join " & _
           "(select Tahun as Tahun7,SPK as SPK7,SB1 as SB17,SB5 as SB57,SBE as SBE7,GRP as GRP7,SVCD as SVCD7,SVCV as SVCV7 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 7) G On A.Tahun1  = G.Tahun7 Left Join " & _
           "(select Tahun as Tahun8,SPK as SPK8,SB1 as SB18,SB5 as SB58,SBE as SBE8,GRP as GRP8,SVCD as SVCD8,SVCV as SVCV8 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 8) H On A.Tahun1  = H.Tahun8 Left Join " & _
           "(select Tahun as Tahun9,SPK as SPK9,SB1 as SB19,SB5 as SB59,SBE as SBE9,GRP as GRP9,SVCD as SVCD9,SVCV as SVCV9 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 9) I On A.Tahun1  = I.Tahun9 Left Join " & _
           "(select Tahun as Tahun10,SPK as SPK10,SB1 as SB110,SB5 as SB510,SBE as SBE10,GRP as GRP10,SVCD as SVCD10,SVCV as SVCV10 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 10) J On A.Tahun1  = J.Tahun10 Left Join " & _
           "(select Tahun as Tahun11,SPK as SPK11,SB1 as SB111,SB5 as SB511,SBE as SBE11,GRP as GRP11,SVCD as SVCD11,SVCV as SVCV11 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 11) K On A.Tahun1  = K.Tahun11 Left Join " & _
           "(select Tahun as Tahun12,SPK as SPK12,SB1 as SB112,SB5 as SB512,SBE as SBE12,GRP as GRP12,SVCD as SVCD12,SVCV as SVCV12 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 12) L On A.Tahun1  = L.Tahun12 Where A.Tahun1 = " & cboThn.Text

            
    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.TxtAlamat.SetText gAlamatvk
    
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    PicBars.Visible = False
    nWait.Visible = False
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub

Private Sub Print_Part()
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New RptDYNA_HO

    sSQL = "Select A.*,B.*,C.*,D.*,E.*,F.*,G.*,H.*,I.*,J.*,K.*,L.* from " & _
           "(select Tahun as Tahun1,SPK as SPK1,SB1 as SB11,SB5 as SB51,SBE as SBE1,GRP as GRP1,SVCD as SVCD1,SVCV as SVCV1 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 1) A Left Join " & _
           "(select Tahun as Tahun2,SPK as SPK2,SB1 as SB12,SB5 as SB52,SBE as SBE2,GRP as GRP2,SVCD as SVCD2,SVCV as SVCV2 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 2) B On A.Tahun1  = B.Tahun2 Left Join " & _
           "(select Tahun as Tahun3,SPK as SPK3,SB1 as SB13,SB5 as SB53,SBE as SBE3,GRP as GRP3,SVCD as SVCD3,SVCV as SVCV3 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 3) C On A.Tahun1  = C.Tahun3 Left Join " & _
           "(select Tahun as Tahun4,SPK as SPK4,SB1 as SB14,SB5 as SB54,SBE as SBE4,GRP as GRP4,SVCD as SVCD4,SVCV as SVCV4 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 4) D On A.Tahun1  = D.Tahun4 Left Join " & _
           "(select Tahun as Tahun5,SPK as SPK5,SB1 as SB15,SB5 as SB55,SBE as SBE5,GRP as GRP5,SVCD as SVCD5,SVCV as SVCV5 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 5) E On A.Tahun1  = E.Tahun5 Left Join " & _
           "(select Tahun as Tahun6,SPK as SPK6,SB1 as SB16,SB5 as SB56,SBE as SBE6,GRP as GRP6,SVCD as SVCD6,SVCV as SVCV6 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 6) F On A.Tahun1  = F.Tahun6 Left Join " & _
           "(select Tahun as Tahun7,SPK as SPK7,SB1 as SB17,SB5 as SB57,SBE as SBE7,GRP as GRP7,SVCD as SVCD7,SVCV as SVCV7 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 7) G On A.Tahun1  = G.Tahun7 Left Join " & _
           "(select Tahun as Tahun8,SPK as SPK8,SB1 as SB18,SB5 as SB58,SBE as SBE8,GRP as GRP8,SVCD as SVCD8,SVCV as SVCV8 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 8) H On A.Tahun1  = H.Tahun8 Left Join " & _
           "(select Tahun as Tahun9,SPK as SPK9,SB1 as SB19,SB5 as SB59,SBE as SBE9,GRP as GRP9,SVCD as SVCD9,SVCV as SVCV9 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 9) I On A.Tahun1  = I.Tahun9 Left Join " & _
           "(select Tahun as Tahun10,SPK as SPK10,SB1 as SB110,SB5 as SB510,SBE as SBE10,GRP as GRP10,SVCD as SVCD10,SVCV as SVCV10 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 10) J On A.Tahun1  = J.Tahun10 Left Join " & _
           "(select Tahun as Tahun11,SPK as SPK11,SB1 as SB111,SB5 as SB511,SBE as SBE11,GRP as GRP11,SVCD as SVCD11,SVCV as SVCV11 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 11) K On A.Tahun1  = K.Tahun11 Left Join " & _
           "(select Tahun as Tahun12,SPK as SPK12,SB1 as SB112,SB5 as SB512,SBE as SBE12,GRP as GRP12,SVCD as SVCD12,SVCV as SVCV12 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 12) L On A.Tahun1  = L.Tahun12 Where A.Tahun1 = " & cboThn.Text

            
    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.TxtAlamat.SetText gAlamatvk
    
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    PicBars.Visible = False
    nWait.Visible = False
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdPrintLaporanByModel_Click()
    Dim boleh As Boolean
    cmdPrintLaporanByModel.Enabled = False
    
    boleh = True
    
    If Text(2) = "" Then
        pesan = pesan + "Tahun Report. "
        boleh = False
    End If
    
    If cboBulan3 = "" Then
        pesan = pesan + "Bulan Report. "
        boleh = False
    End If
    
    If boleh = True Then
      frmRepoBulananByModel.Bulan = Left(cboBulan3, InStr(cboBulan3, " |") - 1)
      frmRepoBulananByModel.Tahun = Trim(Text(2).Text)
      
      
      frmRepoBulananByModel.show vbModal
    End If
    cmdPrintLaporanByModel.Enabled = True
   
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    On Error GoTo errNulDB
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
errNulDB:
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton2.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

Private Sub cmdView_Click()
    Dim f As New FrmTwc_ViewCabang
    f.show
End Sub

Private Sub Form_Activate()
    vNamaTabel = "RepoBulanan" & DriveSerialNumber("c:") & Second(Now)
   
End Sub

Private Sub CetakButton()
    With ImageList1
       ' CmdAdd.PictureNormal = .ListImages("cmdadd").Picture
        cmdAdd.ToolTipText = ""
        cmdAdd.Tag = "Tambah Record"
        cmdAdd.Caption = "&Add"
    End With
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    Call hapusForm
    'SetIcon Me
    'CetakButton
End Sub


Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub BarProsesByModel(labA As Integer, labB As Integer, proA As Integer, proB As Integer, bar As Integer, Ket As String)
    nBarByModel.value = bar
    nLabByModel.Caption = "Collecting Data Report : " & labA & " Of " & labB & ""
    nProByModel.Caption = "Process :  " & proA & " Of " & proB & ""
    nKetByModel.Caption = Ket
    DoEvents
End Sub

Private Sub BarProses(labA As Integer, labB As Integer, proA As Integer, proB As Integer, bar As Integer, Ket As String)
    nBar.value = bar
    nLab.Caption = "Collecting Data Report : " & labA & " Of " & labB & ""
    nPro.Caption = "Process :  " & proA & " Of " & proB & ""
    nKet.Caption = Ket
    DoEvents
End Sub

Private Sub isiBatchBulanan(i As Integer)
      Dim strSql As String

1     nBar.Min = 0
2             nBar.Max = 32
3             nBar.value = 0
4             nLab.Caption = "Get Data Report : 0 Of 32"
5             nPro.Caption = "Process :  1 Of 2"
6             PicBar.Visible = True
7             DoEvents
              
              'For i = Left(cboBulan1, InStr(cboBulan1, " |") - 1) To Left(cboBulan2, InStr(cboBulan2, " |") - 1)
8                 i = Left(cboBulan1, InStr(cboBulan1, " |") - 1)
                  'Set rsTemp = objO.FetchData("EXEC REPOBULANAN3 " & Left(cboBulan1, InStr(cboBulan1, " |") - 1) & ", " & Text(0) & "")
                  
9                 Call objO.DeleteData("Delete FROM BatchBulanan Where Bulan = '" & i & "' and Tahun = '" & Text(0).Text & "'")
                  
                 '6 Agustus 2012
10                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3A " & i & "," & Text(0) & "")
11                Call BarProses(1, 32, 1, 2, 1, "Get Data Unit Entry ...")
12                DoEvents
13                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3B " & i & "," & Text(0) & "")
14                Call BarProses(2, 32, 1, 2, 2, "Get Data Unit Entry ...")
15                DoEvents
16                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3C " & i & "," & Text(0) & "")
17                Call BarProses(3, 32, 1, 2, 3, "Get Data Unit Entry ...")
18                DoEvents
19                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3D " & i & "," & Text(0) & "")
20                Call BarProses(4, 32, 1, 2, 4, "Get Data Revenue ...")
21                DoEvents
22                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3E " & i & "," & Text(0) & "")
23                Call BarProses(5, 32, 1, 2, 5, "Get Data Revenue ...")
24                DoEvents
                  
                  '2015-01-29 PISAH REPO F
25                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3F1 '" & i & "','" & Text(0) & "'")
26                Call BarProses(6, 32, 1, 2, 6, "Get Data Service Time ...")
27                DoEvents
                  
28                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3F2 '" & i & "','" & Text(0) & "'")
29                Call BarProses(6, 32, 1, 2, 6, "Get Data Service Time ...")
30                DoEvents
                  
31                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3F3 '" & i & "','" & Text(0) & "'")
32                Call BarProses(6, 32, 1, 2, 6, "Get Data Service Time ...")
33                DoEvents
                  
                  '======================================
                  
      '            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3F '" & i & "','" & Text(0) & "'")
      '            Call BarProses(6, 32, 1, 2, 6, "Get Data Service Time ...")
      '            DoEvents
34                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3G " & i & "," & Text(0) & "")
35                Call BarProses(7, 32, 1, 2, 7, "Get Data Follow Up ...")
36                DoEvents
37                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3H " & i & "," & Text(0) & "")
38                Call BarProses(8, 32, 1, 2, 8, "Get Data Lead Time ...")
39                DoEvents
      '            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3I " & i & "," & Text(0) & "")
      '            Call BarProses(9, 32, 1, 2, 9, "Get Data Lead Time Notifcation ...")
      '            DoEvents
                  
40                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3I1 " & i & "," & Text(0) & "")
41                Call BarProses(9, 32, 1, 2, 9, "Get Data Lead Time Notifcation ...")
42                DoEvents
                  
43                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3I2 " & i & "," & Text(0) & "")
44                Call BarProses(9, 32, 1, 2, 9, "Get Data Lead Time Notifcation ...")
45                DoEvents
                  
46                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3I3 " & i & "," & Text(0) & "")
47                Call BarProses(9, 32, 1, 2, 9, "Get Data Lead Time Notifcation ...")
48                DoEvents
                  
49                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3I4 " & i & "," & Text(0) & "")
50                Call BarProses(9, 32, 1, 2, 9, "Get Data Lead Time Notifcation ...")
51                DoEvents
                  
52                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3I5 " & i & "," & Text(0) & "")
53                Call BarProses(9, 32, 1, 2, 9, "Get Data Lead Time Notifcation ...")
54                DoEvents
                  
                  '2015-01-27 - PROSES PECAH REPOBULANAN3J - RUDI
55                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3J1 " & i & "," & Text(0) & "")
56                Call BarProses(10, 32, 1, 2, 10, "Get Data Lead Time Delivery ...")
57                DoEvents
58                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3J2 " & i & "," & Text(0) & "")
59                Call BarProses(11, 32, 1, 2, 11, "Get Data Unit Gate 1 Day ...")
60                DoEvents
61                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3J3 " & i & "," & Text(0) & "")
62                Call BarProses(12, 32, 1, 2, 12, "Get Data Unit Gate 1 Day THS...")
63                DoEvents
64                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3J4 " & i & "," & Text(0) & "")
65                Call BarProses(13, 32, 1, 2, 13, "Get Data Unit Gate SBE Total 1 Day ...")
66                DoEvents
67                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3J5 " & i & "," & Text(0) & "")
68                Call BarProses(14, 32, 1, 2, 14, "Get Data Unit Gate SBE Total 1 Day THS ...")
69                DoEvents
                  '------------------------------------------------
                              
      '            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3J " & i & "," & Text(0) & "")
      '            Call BarProses(10, 31,1,2, 10, "Get Data Lead Time Total ...")
      '            DoEvents
                  
                  
70                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3K " & i & "," & Text(0) & "")
71                Call BarProses(15, 32, 1, 2, 15, "Get Data Lead Time Total ...")
72                DoEvents
73                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3L " & i & "," & Text(0) & "")
74                Call BarProses(16, 32, 1, 2, 16, "Get Data Lead Time Total ...")
75                DoEvents
76                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3M " & i & "," & Text(0) & "")
77                Call BarProses(17, 32, 1, 2, 17, "Get Data Lead Time Total ...")
78                DoEvents
                  
79                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3N " & i & "," & Text(0) & "")
80                Call BarProses(18, 32, 1, 2, 18, "Get Data Total Partman...")
81                DoEvents
                  
                  '2015-01-27 - PROSES PECAH REPOBULANAN3O - RUDI
82                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3O1 " & i & "," & Text(0) & "")
83                Call BarProses(19, 32, 1, 2, 19, "Get Data Lead Time Total ...")
84                DoEvents
85                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3O2 " & i & "," & Text(0) & "")
86                Call BarProses(20, 32, 1, 2, 20, "Get Data Lead Time Total ...")
87                DoEvents
88                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3O3 " & i & "," & Text(0) & "")
89                Call BarProses(21, 32, 1, 2, 21, "Get Data Lead Time Total ...")
90                DoEvents
91                Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3O4 " & i & "," & Text(0) & "")
92                Call BarProses(22, 32, 1, 2, 22, "Get Data Lead Time Total ...")
93                DoEvents
                  
                  '----------------------------------------------
      '            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3O " & i & "," & Text(0) & "")
      '            Call BarProses(18, 16, 1, 2, 18, "Get Data Lead Time Total ...")
      '            DoEvents
                  
                  '03/02/2015 penambahan checkbox nofaktur
                  Dim kriteria As String
94                kriteria = ""
95                keteranganFaktur = " NoFaktur "
96                If ckFakturA.value = vbChecked Then
97                    kriteria = kriteria & " A.NoFaktur LIKE 'A%' OR"
98                    keteranganFaktur = keteranganFaktur & " A,"
99                End If
100               If ckFakturB.value = vbChecked Then
101                   kriteria = kriteria & " A.NoFaktur LIKE 'B%' OR"
102                   keteranganFaktur = keteranganFaktur & " B,"
103               End If
104               If ckFakturC.value = vbChecked Then
105                   kriteria = kriteria & " A.NoFaktur LIKE 'C%' OR"
106                   keteranganFaktur = keteranganFaktur & " C,"
107               End If
108               If keteranganFaktur = " NoFaktur " Then
109                   keteranganFaktur = " NoFaktur A,B,C"
110               Else
111                   keteranganFaktur = Left(keteranganFaktur, Len(keteranganFaktur) - 1)
112               End If
                  ''If kriteria = " WHERE " Then
                      ''kriteria = " WHERE NoFaktur NOT LIKE 'C%' and NoFaktur NOT LIKE 'B%' AND NoFaktur NOT LIKE 'A%'"
                  ''Else
                  
                     Dim nilaippn As Integer
                     Dim nilaidpp As Integer
                     Dim nilaitotal As Integer
                     Dim rsPPn As ADODB.Recordset
113                  strSql = "select * from ProfilACC..settingppn WITH (NOLOCK) where ('" & Trim(Text(0)) & "-" & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "-01' between tglawal and tglakhir) "
114                  nilaippn = 10
115                  nilaidpp = 100
116                  nilaitotal = 110
117                  Set rsPPn = objO.FetchData(strSql)
118                  If Not rsPPn.EOF Then
119                       nilaippn = rsPPn!nilaippn
120                       nilaitotal = nilaidpp + nilaippn
121                  End If
122               If kriteria <> "" Then
123                   kriteria = Left(kriteria, Len(kriteria) - 2)
124                   kriteria = " WHERE ( " & kriteria & " )" & "  AND RIGHT(B.SCC_Code, 1 ) IN ('1')"
      '                strSql = "INSERT INTO BatchBulanan "
      '               strSql = strSql & " SELECT Y.KETERANGAN , SUM(Y.REV) AS REVENUE, BLN, THN FROM "
      '               strSql = strSql & " (SELECT MONTH(A.TGLENTRY) AS BLN, YEAR(A.TGLENTRY) AS THN, SUM(CASE WHEN QTY >= 0 THEN ((HargaJual*ABS(QTY))-ItemDisc) ELSE -1*((HargaJual*ABS(QTY))-ItemDisc) END) AS REV,"
      '               strSql = strSql & " case when (Jenis IN ('PART','ACCESSORIES') and TG = 0) then 'Revenue Oli dan Bahan Non Toyota Total' else case when (Jenis IN ('PART','ACCESSORIES') and TG = 1) then 'Revenue Part dan Oli TGMO Total' else"
      '               strSql = strSql & " case when (Jenis = 'OLI' and TG = 0) then 'Revenue Oli dan Bahan Non Toyota Total' else CASE WHEN (Jenis = 'OLI' and TG = 1) then 'Revenue Part dan Oli TGMO Total' END END end end AS KETERANGAN"
      '               strSql = strSql & " FROM [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(text(0)) & Trim(Left(cboBulan1.text, InStr(cboBulan1.text, " |") - 1)) & "]..SISAL A"
      '               strSql = strSql & " INNER JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(text(0)) & Trim(Left(cboBulan1.text, InStr(cboBulan1.text, " |") - 1)) & "]..INMST B ON A.KODEITEM = B.KODEITEM"
      '               strSql = strSql & " LEFT JOIN ProfilACC..Scc C ON B.SCC_Code = C.SCC_Code"
      '               strSql = strSql & kriteria '" WHERE (NoFaktur LIKE 'A%' OR NoFaktur LIKE 'B%' OR NoFaktur LIKE 'C%' )"
      '               strSql = strSql & " group BY Jenis, TG, MONTH(A.TGLENTRY), YEAR(A.TGLENTRY)) AS Y"
      '               strSql = strSql & " where bln = '" & Trim(Left(cboBulan1.text, InStr(cboBulan1.text, " |") - 1)) & "' and thn = '" & Trim(text(0)) & "'"
      '               strSql = strSql & " GROUP BY Y.KETERANGAN,bln,thn"
                     
                     
125                  strSql = "INSERT INTO BatchBulanan "
126                  strSql = strSql & " SELECT Y.KETERANGAN , SUM(Y.REV) AS REVENUE, BLN, THN FROM "
                      ''strSql = strSql & " (SELECT MONTH(A.TGLENTRY) AS BLN, YEAR(A.TGLENTRY) AS THN, sum(( ( A.Qty * A.HargaJual ) - ( A.Qty * A.HargaJual * A.ItemDisPsn ) ) * ( 1- CASE D.TypePpn WHEN 'I' THEN 0.0909091 ELSE 0 END )) AS REV,"
127                  strSql = strSql & " (SELECT MONTH(A.TGLENTRY) AS BLN, YEAR(A.TGLENTRY) AS THN, sum(( ( A.Qty * A.HargaJual ) - ( A.Qty * A.HargaJual * A.ItemDisPsn ) ) * ( 1- CASE D.TypePpn WHEN 'I' THEN " & k2t(nilaippn / nilaitotal) & " ELSE 0 END )) AS REV,"
128                  strSql = strSql & " case when (Jenis IN ('PART','ACCESSORIES') and TG = 0) then 'Revenue Oli dan Bahan Non Toyota Total' else case when (Jenis IN ('PART','ACCESSORIES') and TG = 1) then 'Revenue Part dan Oli TGMO Total' else"
129                  strSql = strSql & " case when (Jenis = 'OLI' and TG = 0) then 'Revenue Oli dan Bahan Non Toyota Total' else CASE WHEN (Jenis = 'OLI' and TG = 1) then 'Revenue Part dan Oli TGMO Total' END END end end AS KETERANGAN"
130                  strSql = strSql & " FROM [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SISAL A"
131                  strSql = strSql & " INNER JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..INMST B ON A.KODEITEM = B.KODEITEM"
132                  strSql = strSql & " INNER JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..Sihdr D ON A.NoFaktur = D.NoFaktur"
133                  strSql = strSql & " LEFT JOIN ProfilACC..Scc C ON B.SCC_Code = C.SCC_Code"
134                  strSql = strSql & kriteria '" WHERE (NoFaktur LIKE 'A%' OR NoFaktur LIKE 'B%' OR NoFaktur LIKE 'C%' )"
135                  strSql = strSql & " group BY Jenis, TG, MONTH(A.TGLENTRY), YEAR(A.TGLENTRY)) AS Y"
136                  strSql = strSql & " where bln = '" & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "' and thn = '" & Trim(Text(0)) & "'"
137                  strSql = strSql & " GROUP BY Y.KETERANGAN,bln,thn"
                     
138                  Call objO.InsertData(strSql)
139               End If
                  
140               Call BarProses(23, 32, 1, 2, 23, "Get Data Report ...")
141               DoEvents

142               Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3N " & i & "," & Text(0) & "")
143               Call BarProses(24, 32, 1, 2, 24, "Get Data Total Partman...")
144               DoEvents

145               Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3P " & i & "," & Text(0) & "")
146               Call BarProses(25, 32, 1, 2, 25, "Get Data Total Partman...")
147               DoEvents

148               Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3Q " & i & "," & Text(0) & "")
149               Call BarProses(26, 32, 1, 2, 26, "Get Data Total Partman...")
150               DoEvents

                  '2015-01-29 PENGAMBILAN DATA SERVICE POINT
      '            strSql = "INSERT INTO BatchBulanan "
      '            strSql = strSql & " SELECT * FROM ("
      '            strSql = strSql & " SELECT 'Unit SP SBE' as keterangan,COUNT(NOINV) AS JUMLAH, '" & i & "' AS BULAN, '" & text(0) & "' AS TAHUN  FROM T_INVOICE where job='1' AND FS_RC<> '' AND FS_RC <> 'R001' AND MONTH(INVDATE) = '" & i & "' AND YEAR(INVDATE) = '" & text(0) & "' "
      '            strSql = strSql & " Union"
      '            strSql = strSql & " SELECT 'Unit SP Total' as keterangan,COUNT(NOINV) AS JUMLAH, '" & i & "' AS BULAN, '" & text(0) & "' AS TAHUN FROM T_INVOICE where job='1' and MONTH(INVDATE) = '" & i & "' AND YEAR(INVDATE) = '" & text(0) & "'"
      '            strSql = strSql & " Union"
      '            strSql = strSql & " SELECT top 1 'Jumlah Service Point' as keterangan, sp as jumlah,'" & i & "' AS BULAN, '" & text(0) & "' AS TAHUN from QTY_STALL order by Tanggal desc"
      '            strSql = strSql & " ) A"
      '            strSql = strSql & " GROUP BY JUMLah,keterangan,BULAN,TAHUN"
      '            strSql = strSql & " ORDER BY keterangan"
      '            Call objO.InsertData(strSql)
151               Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3R " & i & "," & Text(0) & "")
152               Call BarProses(27, 32, 1, 2, 27, "Get Data Unit SP ...")
153               DoEvents

                  '2015-01-29 PENGAMBILAN DATA EM
154               strSql = "INSERT INTO BATCHBULANAN "
155               strSql = strSql & " select 'Unit EM Total' as keterangan, COUNT(NOINV) AS JUMLAH,'" & i & "' AS BULAN, '" & Text(0) & "' AS TAHUN FROM T_INVOICE where code='1' and MONTH(INVDATE) = '" & i & "' AND YEAR(INVDATE) = '" & Text(0) & "' "
156               Call objO.InsertData(strSql)
157               Call BarProses(28, 32, 1, 2, 28, "Get Data TOTAL EM ...")
158               DoEvents

                  '2015-01-29 PENGAMBILAN DATA FILLRATE / SERVICE RATE
                  ''
159               If Log.isFill = 1 Then
160                   strSql = " SELECT SUM(Q.QTYREQ), SUM(QTYSUPPLY) AS LiSupply, SUM(BOQTY) AS LiBO, SUM(QTYCANCEL) AS LiCancel, isnull(SUM(CAST(Q.FR AS DECIMAL(10,2))) / COUNT(Q.NOSO),0) AS FR,"
161                   strSql = strSql & " isnull(((SUM(CAST(QTYSUPPLY AS DECIMAL(10,2)))/ CAST(SUM(QTYSUPPLY)+SUM(BOQTY)+SUM(QTYCANCEL) AS DECIMAL(10,2)))*100 ),0)AS SR"
162                   strSql = strSql & " FROM ("
163                   strSql = strSql & " SELECT Z.NOSO, SUM(Z.Qty) AS QTYREQ, SUM(QTYSUPPLY) AS QTYSUPPLY, SUM(BOQTY) AS BOQTY, SUM(QTYCANCEL) AS QTYCANCEL"
164                   strSql = strSql & " ,case when (SELECT distinct(A.noso)"
165                   strSql = strSql & " From [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..soord a WITH (NOLOCK) "
166                   strSql = strSql & " LEFT JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..Sohdr B WITH (NOLOCK) ON A.NoSo = B.NOSO"
167                   strSql = strSql & " LEFT JOIN BENGKEL" & Trim(Right(gDb, 2)) & ".DBO.t_invoice C WITH (NOLOCK) ON A.NoSo = C.NO_WO AND FS_LOG_DEL = '' AND NOINV NOT LIKE '%-B'"
168                   strSql = strSql & " where A.NOSO NOT IN (SELECT NOSO FROM SOORD WITH (NOLOCK)  WHERE DATEDIFF(hh,tglrange,tglsupply) > 0) AND"
169                   strSql = strSql & " a.noso not in (select NoSo from [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..BoFile WITH (NOLOCK) ) and"
170                   strSql = strSql & " isnull(a.bo,'') <> 'CANCEL' and a.Qty = (isnull(a.QtyKirimBi,0)+ISNULL(a.qtykirimbl,0)+ISNULL(a.Qtykeluar,0))"
171                   strSql = strSql & " and a.NoSpk<>a.noso AND B.KodeSumber = 'SV' and a.NoSo = z.NoSo) <> '' then 100 else 0 end as FR"
172                   strSql = strSql & " ,cast(convert(decimal(10,4), SUM(QTYSUPPLY)) / convert(decimal(10,4), (SUM(QTYSUPPLY)+SUM(BOQTY)+SUM(QTYCANCEL))) as decimal(10,4))*100 AS SR"
173                   strSql = strSql & " FROM ("
174                   strSql = strSql & " select A.NOSO, A.NOMOR, SuM(A.Qty) AS QTY, CASE WHEN SUM(ISNULL(A.QtyKirimBI,0)+ISNULL(A.QtyKirimBL,0)+ISNULL(A.Qtykeluar,0)) > 0 THEN 1 ELSE 0 END AS QTYSUPPLY,"
175                   strSql = strSql & " CASE WHEN isnull(bo,'') <> 'CANCEL' AND SUM(ISNULL(B.Qty,0)) > 0 THEN 1 ELSE 0 END AS BOQTY,"
176                   strSql = strSql & " CASE WHEN BO = 'CANCEL' AND SUM(ISNULL(A.Qty,0) - (ISNULL(A.QtyKirimBI,0)+ISNULL(A.QtyKirimBL,0)+ISNULL(A.Qtykeluar,0))) > 0 THEN 1 ELSE 0 END AS QTYCANCEL"
177                   strSql = strSql & " FROM [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "].dbo.Soord A WITH (NOLOCK) LEFT JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "].dbo.BoFile B WITH (NOLOCK) ON A.NoSo = B.NoSo AND A.KodeItem = B.KodeItem"
178                   strSql = strSql & " LEFT JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "].dbo.Sohdr C WITH (NOLOCK) ON A.NoSo = C.NoSo"
179                   strSql = strSql & " LEFT JOIN BENGKEL" & Trim(Right(gDb, 2)) & ".DBO.T_INVOICE D WITH (NOLOCK) ON A.NoSo = D.NO_WO AND FS_LOG_DEL = '' AND NOINV NOT LIKE '%-B'"
180                   strSql = strSql & " WHERE MONTH(D.DATE) = " & i & " and YEAR(D.DATE) = " & Text(0) & " AND A.NoSo <> A.NOSPK AND C.KodeSumber = 'SV'"
181                   strSql = strSql & " GROUP BY A.NOSO, A.NOMOR, BO"
182                   strSql = strSql & " ) Z"
183                   strSql = strSql & " GROUP BY Z.NoSo"
184                   strSql = strSql & " Having (Sum(QTYSUPPLY) + Sum(BOQTY) + Sum(QTYCANCEL)) > 0"
185                   strSql = strSql & " ) Q"
                      Dim rsIsi As New ADODB.Recordset
                      'strSql = "select top 10 * from inmst"
186                   Set rsTemp = objO.FetchDataACC(strSql)
187                   If Not rsTemp.EOF Then
188                       rsTemp.MoveFirst
189                       Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Fill Rate'," & Round(rsTemp("FR"), 0) & ",'" & i & "','" & Text(0) & "')")
190                       Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Service Rate'," & Round(rsTemp("SR"), 0) & ",'" & i & "','" & Text(0) & "')")

                          'STOCK DAY MASIH DIBERI NILAI KOSONG
191                       Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Stock Day',0,'" & i & "','" & Text(0) & "')")
192                   Else
193                       Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Fill Rate',0,'" & i & "','" & Text(0) & "')")
194                       Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Service Rate',0,'" & i & "','" & Text(0) & "')")

                          'STOCK DAY MASIH DIBERI NILAI KOSONG
195                       Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Stock Day',0,'" & i & "','" & Text(0) & "')")
196                   End If
197               Else
198                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Fill Rate',0,'" & i & "','" & Text(0) & "')")
199                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Service Rate',0,'" & i & "','" & Text(0) & "')")

                      'STOCK DAY MASIH DIBERI NILAI KOSONG
200                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Stock Day',0,'" & i & "','" & Text(0) & "')")
201               End If
202               Call BarProses(29, 32, 1, 2, 29, "Get Data TOTAL EM ...")
203               DoEvents
                  
                  '' ---------------------------------------------------------------------------------
                  '' 20191108 REQ PAK ANTOK ATAS LAPORAN FASA : OLIVOLUME DIGANTI DENGAN TipePartList
                  '' ---------------------------------------------------------------------------------

                  '2015-01-29
                  'TMO
                  ''strSql = "SELECT 'TMO Mineral' as Keterangan, ISNULL(SUM(B.QTY*D.Liter),0) AS total "
204               strSql = "SELECT 'TMO Mineral' as Keterangan, ISNULL(SUM(B.QTY*D.QTY),0) AS total "
205               strSql = strSql & " FROM [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SIHDR A WITH(NOLOCK) INNER JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR"
206               strSql = strSql & " LEFT JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..Inmst C ON C.KodeItem = B.KodeItem"
                  ''strSql = strSql & " LEFT JOIN ProfilACC..OliVolume D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
207               strSql = strSql & " LEFT JOIN TipePartList D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
208               strSql = strSql & " WHERE (A.NoFaktur LIKE 'A%' OR A.NoFaktur LIKE 'C%' OR A.NoFaktur LIKE 'D%' OR A.NoFaktur LIKE 'F%')"
209               strSql = strSql & " AND C.SCC_Code LIKE 'TMO%'"
      '            srtsql = strSql & " GROUP BY A.TGLFAKTUR, SCC_Code"

210               Set rsTemp = objO.FetchData(strSql)
211               If Not rsTemp.EOF Then
212                   rsTemp.MoveFirst
213                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('" & rsTemp("keterangan") & "'," & k2t(rsTemp("total")) & ",'" & i & "','" & Text(0) & "')")
214               Else
215                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('TMO Mineral',0,'" & i & "','" & Text(0) & "')")
216               End If
                  '-----------------====

                  'TMO SYNTETIC TMS
      ''            strSql = "SELECT 'TMO Syntetic' as Keterangan, ISNULL(SUM(B.QTY*D.Liter),0) AS total "
217               strSql = "SELECT 'TMO Syntetic' as Keterangan, ISNULL(SUM(B.QTY*D.QTY),0) AS total "
218               strSql = strSql & " FROM [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SIHDR A WITH(NOLOCK) INNER JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR"
219               strSql = strSql & " LEFT JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..Inmst C ON C.KodeItem = B.KodeItem"
                  ''strSql = strSql & " LEFT JOIN ProfilACC..OliVolume D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
220               strSql = strSql & " LEFT JOIN TipePartList D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
221               strSql = strSql & " WHERE (A.NoFaktur LIKE 'A%' OR A.NoFaktur LIKE 'C%' OR A.NoFaktur LIKE 'D%' OR A.NoFaktur LIKE 'F%')"
222               strSql = strSql & " AND C.SCC_Code LIKE 'TMS%'"
      '            srtsql = strSql & " GROUP BY A.TGLFAKTUR, SCC_Code"

223               Set rsTemp = objO.FetchData(strSql)
224               If Not rsTemp.EOF Then
225                   rsTemp.MoveFirst
226                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('" & rsTemp("keterangan") & "'," & rsTemp("total") & ",'" & i & "','" & Text(0) & "')")
227               Else
228                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('TMO Syntetic',0,'" & i & "','" & Text(0) & "')")
229               End If
                  '-----------------====

                  'TMO FULL SYNTETIC TMF
      ''            strSql = "SELECT 'TMO Full Syntetic' as Keterangan, ISNULL(SUM(B.QTY*D.Liter),0) AS total "
230               strSql = "SELECT 'TMO Full Syntetic' as Keterangan, ISNULL(SUM(B.QTY*D.QTY),0) AS total "
231               strSql = strSql & " FROM [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SIHDR A WITH(NOLOCK) INNER JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR"
232               strSql = strSql & " LEFT JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..Inmst C ON C.KodeItem = B.KodeItem"
                  ''strSql = strSql & " LEFT JOIN ProfilACC..OliVolume D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
233               strSql = strSql & " LEFT JOIN TipePartList D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
234               strSql = strSql & " WHERE (A.NoFaktur LIKE 'A%' OR A.NoFaktur LIKE 'C%' OR A.NoFaktur LIKE 'D%' OR A.NoFaktur LIKE 'F%')"
235               strSql = strSql & " AND C.SCC_Code LIKE 'TMF%'"
      '            srtsql = strSql & " GROUP BY A.TGLFAKTUR, SCC_Code"

236               Set rsTemp = objO.FetchData(strSql)
237               If Not rsTemp.EOF Then
238                   rsTemp.MoveFirst
239                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('" & rsTemp("keterangan") & "'," & rsTemp("total") & ",'" & i & "','" & Text(0) & "')")
240               Else
241                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('TMO Full Syntetic',0,'" & i & "','" & Text(0) & "')")
242               End If
                  '-----------------====

                  'TGGO TGG
                  ''strSql = "SELECT 'TGGO' as Keterangan, ISNULL(SUM(B.QTY*D.Liter),0) AS total "
243               strSql = "SELECT 'TGGO' as Keterangan, ISNULL(SUM(B.QTY*D.QTY),0) AS total "
244               strSql = strSql & " FROM [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SIHDR A WITH(NOLOCK) INNER JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR"
245               strSql = strSql & " LEFT JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..Inmst C ON C.KodeItem = B.KodeItem"
                  ''strSql = strSql & " LEFT JOIN ProfilACC..OliVolume D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
246               strSql = strSql & " LEFT JOIN TipePartList D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
247               strSql = strSql & " WHERE (A.NoFaktur LIKE 'A%' OR A.NoFaktur LIKE 'C%' OR A.NoFaktur LIKE 'D%' OR A.NoFaktur LIKE 'F%')"
248               strSql = strSql & " AND C.SCC_Code LIKE 'TGG%'"
      '            srtsql = strSql & " GROUP BY A.TGLFAKTUR, SCC_Code"

249               Set rsTemp = objO.FetchData(strSql)
250               If Not rsTemp.EOF Then
251                   rsTemp.MoveFirst
252                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('" & rsTemp("keterangan") & "'," & rsTemp("total") & ",'" & i & "','" & Text(0) & "')")
253               Else
254                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('TGGO',0,'" & i & "','" & Text(0) & "')")
255               End If
                  '-----------------====

                   'MINERAL NON TOYOTA NOE
                  ''strSql = "SELECT 'Mineral NOE' as Keterangan, ISNULL(SUM(B.QTY*D.Liter),0) AS total "
256               strSql = "SELECT 'Mineral NOE' as Keterangan, ISNULL(SUM(B.QTY*D.QTY),0) AS total "
257               strSql = strSql & " FROM [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SIHDR A WITH(NOLOCK) INNER JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR"
258               strSql = strSql & " LEFT JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..Inmst C ON C.KodeItem = B.KodeItem"
                  ''strSql = strSql & " LEFT JOIN ProfilACC..OliVolume D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
259               strSql = strSql & " LEFT JOIN TipePartList D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
260               strSql = strSql & " WHERE (A.NoFaktur LIKE 'A%' OR A.NoFaktur LIKE 'C%' OR A.NoFaktur LIKE 'D%' OR A.NoFaktur LIKE 'F%')"
261               strSql = strSql & " AND C.SCC_Code LIKE 'NOE%'"
      '            srtsql = strSql & " GROUP BY A.TGLFAKTUR, SCC_Code"

262               Set rsTemp = objO.FetchData(strSql)
263               If Not rsTemp.EOF Then
264                   rsTemp.MoveFirst
265                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('" & rsTemp("keterangan") & "'," & rsTemp("total") & ",'" & i & "','" & Text(0) & "')")
266               Else
267                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Mineral NOE',0,'" & i & "','" & Text(0) & "')")
268               End If
                  '-----------------====

                  'OLI TRANSMISI NOO
                  ''strSql = "SELECT 'Oli Transmisi' as Keterangan, ISNULL(SUM(B.QTY*D.Liter),0) AS total "
269               strSql = "SELECT 'Oli Transmisi' as Keterangan, ISNULL(SUM(B.QTY*D.QTY),0) AS total "
270               strSql = strSql & " FROM [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SIHDR A WITH(NOLOCK) INNER JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..SISAL B WITH(NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR"
271               strSql = strSql & " LEFT JOIN [ACC" & Trim(Right(gDb, 2)) & "-" & Trim(Text(0)) & Trim(Left(cboBulan1.Text, InStr(cboBulan1.Text, " |") - 1)) & "]..Inmst C ON C.KodeItem = B.KodeItem"
                  ''strSql = strSql & " LEFT JOIN ProfilACC..OliVolume D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
272               strSql = strSql & " LEFT JOIN TipePartList D ON SUBSTRING(C.KodeItem,3,LEN(C.KODEITEM)-2) = D.KodeItem"
273               strSql = strSql & " WHERE (A.NoFaktur LIKE 'A%' OR A.NoFaktur LIKE 'C%' OR A.NoFaktur LIKE 'D%' OR A.NoFaktur LIKE 'F%')"
274               strSql = strSql & " AND C.SCC_Code LIKE 'NOO%'"
      '            srtsql = strSql & " GROUP BY A.TGLFAKTUR, SCC_Code"

275               Set rsTemp = objO.FetchData(strSql)
276               If Not rsTemp.EOF Then
277                   rsTemp.MoveFirst
278                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('" & rsTemp("keterangan") & "'," & rsTemp("total") & ",'" & i & "','" & Text(0) & "')")
279               Else
280                   Call objO.InsertData("INSERT INTO BATCHBULANAN VALUES ('Oli Transmisi',0,'" & i & "','" & Text(0) & "')")
281               End If
                  '-----------------====
282               Call BarProses(30, 32, 1, 2, 30, "Get Data Penjualan Oli  ...")
283               DoEvents

End Sub

Private Sub cetakLaporanBulanan()
1     cboBulan2.Text = cboBulan1.Text
2     Text(1).Text = Text(0).Text
3     Screen.MousePointer = vbHourglass
      Dim pesan As String
      Dim temp As String
      Dim boleh As Boolean
      Dim nI As Integer
      Dim nTotal As Double
      Dim nTotal2 As Double
      Dim nSubTotal As Double
      Dim namaBulan As String
      Dim nUrut As Integer
      Dim nKeterangan As String
      ''========================''

      Dim nilJasaPPn As Double
      Dim nilJasaNonPPn As Double
      Dim nilJasaTotal As Double

      ''========================''
      Dim nHrsFollow As Integer
      Dim nDiFollow As Integer
      Dim nNoQ1 As Integer
      Dim nNoQ2 As Integer
      Dim nNoQ3 As Integer
      Dim nRTJ As Integer

      ''========================''
      Dim nHrsUndang  As Integer
      Dim nDiUndang   As Integer
      Dim nDtgUndang  As Integer

      ''========================''
      Dim GTG         As Double
      Dim UE          As Double
      Dim UEnT        As Double
      Dim PN          As Double
      Dim GTGs        As Double
      Dim UEs         As Double
      Dim UEsTHS      As Double
      Dim PNs         As Double
      Dim TDN         As Double
      Dim SBESBI      As Double
      Dim PTh         As Double
      Dim TTh         As Double
      Dim jTHS        As Double
      Dim TotalSBE    As Double
      Dim totalUnitEntryTHS As Double
      Dim totalUnitEntry As Double
      Dim TotalUnitSBE As Double
      Dim Jumlah1K As Double
      Dim i As Integer
      Dim strSql As String


      Dim subTotalTMO As Double
      Dim oliMesinNonToyota As Double
      Dim totalLaporan As Integer

4         boleh = True
5         If Text(0) = "" Then
6             pesan = pesan + "Tahun Awal. "
7             boleh = False
8         End If
9         If Text(1) = "" Then
10            pesan = pesan + "Tahun Akhir. "
11            boleh = False
12        End If
13        If CDbl(Text(0)) <> CDbl(Text(1)) Then
14            pesan = pesan + "Tahun harus sama. "
15            boleh = False
16        End If
17        If cboBulan1 = "" Then
18            pesan = pesan + "Bulan Awal. "
19            boleh = False
20        End If
21        If cboBulan2 = "" Then
22            pesan = pesan + "Bulan Akhir. "
23            boleh = False
24        End If
          
25        If boleh Then
26            Shape1.Height = 1575
27            sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
                   & " drop table [tempdb]..[" & vNamaTabel & "]"
28            Call objO.UpdateData(sSQL)
              
29            temp = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ([URUT][NUMERIC](9) NOT NULL DEFAULT 0, [KETERANGAN][varchar](100) NOT NULL DEFAULT '', [BOLD][char](1) NOT NULL DEFAULT '0' "
30            For i = Left(cboBulan1, InStr(cboBulan1, " |") - 1) To Left(cboBulan2, InStr(cboBulan2, " |") - 1)
31                temp = temp + ", [" & GetNamaBulan(i) & "][money] NOT NULL DEFAULT 0"
32            Next
33            temp = temp + ") ON [PRIMARY]"
34            Call objO.UpdateData(temp)
              
              '6 Agustus 2012
35                Call isiBatchBulanan(i)
36                subTotalTMO = 0
37                oliMesinNonToyota = 0
                  
38                Set rsTemp = objO.FetchData("Select Keterangan, SUM(JUMLAH) AS JUMLAH, Bulan, TAHUN from BatchBulanan  Where Bulan = '" & i & "' and Tahun = '" & Text(0).Text & "' group by BULAN, Tahun, KETERANGAN")
39                Call BarProses(31, 32, 1, 2, 31, "Get Data Report ...")
40                DoEvents
                  
41                If Not rsTemp.EOF Then
42                    nTotal = 0
43                    nTotal2 = 0
44                    nSubTotal = 0
45                    nI = 1
46                    namaBulan = GetNamaBulan(Left(cboBulan1, InStr(cboBulan1, " |") - 1))

47                    nBar.Min = 0
48                    nBar.Max = 125
49                    nBar.value = 0
50                    nLab.Caption = "Calculation Data : 0 Of 120"
51                    nPro.Caption = "Process :  2 Of 3"
52                    DoEvents
                      
53                    totalLaporan = 119
54                    If Log.hansaRenkeiBP = 1 Then
55                        totalLaporan = 127
56                        nBar.Max = 140
57                    End If
58                    Do While nI <= totalLaporan
59                        Select Case nI
                              Case 1
60                                nUrut = nI
61                                nKeterangan = "10K saja "
62                                rsTemp.Find "KETERANGAN = '10'", , adSearchForward, 1
63                                If Not rsTemp.EOF Then
64                                    If cekInsert(nUrut, nKeterangan) Then
65                                        sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
66                                    Else
67                                        sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
68                                    End If
69                                    nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
70                                Else
71                                    If cekInsert(nUrut, nKeterangan) Then
72                                        sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
73                                    Else
74                                        sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
75                                    End If
76                                    nTotal = nTotal + 0
77                                End If
78                                Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 10 K ...")
79                            Case 2
80                                nUrut = nI
81                                nKeterangan = "20K saja "
82                                rsTemp.Find "KETERANGAN = '20'", , adSearchForward, 1
83                                If Not rsTemp.EOF Then
84                                    If cekInsert(nUrut, nKeterangan) Then
85                                        sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
86                                    Else
87                                        sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
88                                    End If
89                                    nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
90                                Else
91                                    If cekInsert(nUrut, nKeterangan) Then
92                                        sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
93                                    Else
94                                        sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
95                                    End If
96                                    nTotal = nTotal + 0
97                                End If
98                                Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 20 K ...")
99                            Case 3
100                               nUrut = nI
101                               nKeterangan = "30K saja "
102                               rsTemp.Find "KETERANGAN = '30'", , adSearchForward, 1
103                               If Not rsTemp.EOF Then
104                                   If cekInsert(nUrut, nKeterangan) Then
105                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
106                                   Else
107                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
108                                   End If
109                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
110                               Else
111                                   If cekInsert(nUrut, nKeterangan) Then
112                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
113                                   Else
114                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
115                                   End If
116                                   nTotal = nTotal + 0
117                               End If
118                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 30 K ...")
119                           Case 4
120                               nUrut = nI
121                               nKeterangan = "40K saja "
122                               rsTemp.Find "KETERANGAN = '40'", , adSearchForward, 1
123                               If Not rsTemp.EOF Then
124                                   If cekInsert(nUrut, nKeterangan) Then
125                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
126                                   Else
127                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
128                                   End If
129                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
130                               Else
131                                   If cekInsert(nUrut, nKeterangan) Then
132                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
133                                   Else
134                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
135                                   End If
136                                   nTotal = nTotal + 0
137                               End If
138                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 40 K ...")
139                           Case 5
140                               nUrut = nI
141                               nKeterangan = "50K saja "
142                               rsTemp.Find "KETERANGAN = '50'", , adSearchForward, 1
143                               If Not rsTemp.EOF Then
144                                   If cekInsert(nUrut, nKeterangan) Then
145                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
146                                   Else
147                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
148                                   End If
149                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
150                               Else
151                                   If cekInsert(nUrut, nKeterangan) Then
152                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
153                                   Else
154                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
155                                   End If
156                                   nTotal = nTotal + 0
157                               End If
158                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 50 K ...")
159                           Case 6
160                               nUrut = nI
161                               nKeterangan = "60K saja "
162                               rsTemp.Find "KETERANGAN = '60'", , adSearchForward, 1
163                               If Not rsTemp.EOF Then
164                                   If cekInsert(nUrut, nKeterangan) Then
165                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
166                                   Else
167                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
168                                   End If
169                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
170                               Else
171                                   If cekInsert(nUrut, nKeterangan) Then
172                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
173                                   Else
174                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
175                                   End If
176                                   nTotal = nTotal + 0
177                               End If
178                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 60 K ...")
179                           Case 7
180                               nUrut = nI
181                               nKeterangan = "70K saja "
182                               rsTemp.Find "KETERANGAN = '70'", , adSearchForward, 1
183                               If Not rsTemp.EOF Then
184                                   If cekInsert(nUrut, nKeterangan) Then
185                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
186                                   Else
187                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
188                                   End If
189                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
190                               Else
191                                   If cekInsert(nUrut, nKeterangan) Then
192                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
193                                   Else
194                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
195                                   End If
196                                   nTotal = nTotal + 0
197                               End If
198                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 70 K ...")
199                           Case 8
200                               nUrut = nI
201                               nKeterangan = "80K saja "
202                               rsTemp.Find "KETERANGAN = '80'", , adSearchForward, 1
203                               If Not rsTemp.EOF Then
204                                   If cekInsert(nUrut, nKeterangan) Then
205                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
206                                   Else
207                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
208                                   End If
209                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
210                               Else
211                                   If cekInsert(nUrut, nKeterangan) Then
212                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
213                                   Else
214                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
215                                   End If
216                                   nTotal = nTotal + 0
217                               End If
218                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 80 K ...")
219                           Case 9
220                               nUrut = nI
221                               nKeterangan = "90K saja "
222                               rsTemp.Find "KETERANGAN = '90'", , adSearchForward, 1
223                               If Not rsTemp.EOF Then
224                                   If cekInsert(nUrut, nKeterangan) Then
225                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
226                                   Else
227                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
228                                   End If
229                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
230                               Else
231                                   If cekInsert(nUrut, nKeterangan) Then
232                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
233                                   Else
234                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
235                                   End If
236                                   nTotal = nTotal + 0
237                               End If
238                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 90 K ...")
239                           Case 10
240                               nUrut = nI
241                               nKeterangan = "100K saja "
242                               rsTemp.Find "KETERANGAN = '100'", , adSearchForward, 1
243                               If Not rsTemp.EOF Then
244                                   If cekInsert(nUrut, nKeterangan) Then
245                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
246                                   Else
247                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
248                                   End If
249                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
250                               Else
251                                   If cekInsert(nUrut, nKeterangan) Then
252                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
253                                   Else
254                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
255                                   End If
256                                   nTotal = nTotal + 0
257                               End If
258                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 100 K ...")
259                           Case 11
260                               nUrut = nI
261                               nKeterangan = "Kelipatan 5K dibawah 100K "
262                               rsTemp.Find "KETERANGAN = 'Kelipatan 5K'", , adSearchForward, 1
263                               If Not rsTemp.EOF Then
264                                   If cekInsert(nUrut, nKeterangan) Then
265                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
266                                   Else
267                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
268                                   End If
269                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
270                               Else
271                                   If cekInsert(nUrut, nKeterangan) Then
272                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
273                                   Else
274                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
275                                   End If
276                                   nTotal = nTotal + 0
277                               End If
278                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Kelipatan 5 K < 100 K ...")
279                           Case 12
280                               nUrut = nI
281                               nKeterangan = "SBE lainnya diatas 100K "
282                               rsTemp.Find "KETERANGAN = 'Lainnya Diatas 100K'", , adSearchForward, 1
283                               If Not rsTemp.EOF Then
284                                   If cekInsert(nUrut, nKeterangan) Then
285                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
286                                   Else
287                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
288                                   End If
289                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
290                               Else
291                                   If cekInsert(nUrut, nKeterangan) Then
292                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
293                                   Else
294                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
295                                   End If
296                                   nTotal = nTotal + 0
297                               End If
298                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data > 100 K ...")
299                           Case 13
                                  ' TOTAL SBE
300                               nUrut = nI
301                               nKeterangan = "  TOTAL SERVIS BERKALA EKSTERNAL "
302                               If cekInsert(nUrut, nKeterangan) Then
303                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
304                               Else
305                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
306                               End If
307                               nSubTotal = nSubTotal + nTotal
308                               TotalSBE = nSubTotal
309                               TotalUnitSBE = TotalSBE
310                               nTotal = 0
311                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Total SBE ...")
312                           Case 14
313                               nUrut = nI
314                               nKeterangan = "Unit Ganti Oli Mesin "
315                               rsTemp.Find "KETERANGAN = 'Unit Ganti Oli Mesin'", , adSearchForward, 1
316                               If Not rsTemp.EOF Then
317                                   If cekInsert(nUrut, nKeterangan) Then
318                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
319                                   Else
320                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
321                                   End If
322                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
323                               Else
324                                   If cekInsert(nUrut, nKeterangan) Then
325                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
326                                   Else
327                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
328                                   End If
329                                   nTotal = nTotal + 0
330                               End If
331                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Ganti Oli Mesin ...")
332                           Case 15
333                               nUrut = nI
334                               nKeterangan = "General Repair Lainnya "
335                               rsTemp.Find "KETERANGAN = 'Lainnya General Repair1'", , adSearchForward, 1
336                               If Not rsTemp.EOF Then
337                                   If cekInsert(nUrut, nKeterangan) Then
338                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
339                                   Else
340                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
341                                   End If
342                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
343                               Else
344                                   If cekInsert(nUrut, nKeterangan) Then
345                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
346                                   Else
347                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
348                                   End If
349                                   nTotal = nTotal + 0
350                               End If
351                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Other GR ...")
352                           Case 16
353                               nUrut = nI - 1
354                               nKeterangan = "General Repair Lainnya "
355                               rsTemp.Find "KETERANGAN = 'Lainnya General Repair2'", , adSearchForward, 1
356                               If Not rsTemp.EOF Then
357                                   If cekInsert(nUrut, nKeterangan) Then
358                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
359                                   Else
360                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 3)
361                                   End If
362                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
363                               Else
364                                   If cekInsert(nUrut, nKeterangan) Then
365                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
366                                   Else
367                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 3)
368                                   End If
369                                   nTotal = nTotal + 0
370                               End If
371                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Other GR ...")
372                           Case 17
                                  ' SUB TOTAL GRP
373                               nUrut = nI - 1
374                               nKeterangan = "  SUBTOTAL GENERAL REPAIR "
375                               If cekInsert(nUrut, nKeterangan) Then
376                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
377                               Else
378                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
379                               End If
380                               nSubTotal = nSubTotal + nTotal
381                               nTotal = 0
382                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Sub Total GR ...")
383                           Case 18
                                  ' TOTAL CPUS
384                               nUrut = nI - 1
385                               nKeterangan = "  TOTAL CPUS "
386                               If cekInsert(nUrut, nKeterangan) Then
387                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 1, True)
388                               Else
389                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 2, True)
390                               End If
391                               nTotal = 0
392                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Total CPUS ...")
393                           Case 19
394                               nUrut = nI - 1
395                               nKeterangan = "1K Penjualan Outlet Sama "
396                               rsTemp.Find "KETERANGAN = 'SBI 1 dealer'", , adSearchForward, 1
397                               If Not rsTemp.EOF Then
398                                   If cekInsert(nUrut, nKeterangan) Then
399                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
400                                   Else
401                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
402                                   End If
403                               Else
404                                   If cekInsert(nUrut, nKeterangan) Then
405                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
406                                   Else
407                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
408                                   End If
409                               End If
410                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 1K Local ...")
411                           Case 20
412                               nUrut = nI - 1
413                               nKeterangan = "1K Total "
414                               rsTemp.Find "KETERANGAN = 'SBI beda dealer'", , adSearchForward, 1
415                               If Not rsTemp.EOF Then
416                                   If cekInsert(nUrut, nKeterangan) Then
417                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
418                                   Else
419                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
420                                   End If
421                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
422                               Else
423                                   If cekInsert(nUrut, nKeterangan) Then
424                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
425                                   Else
426                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
427                                   End If
428                               End If
429                               Jumlah1K = nTotal
430                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 1K Total ...")
431                           Case 21
432                               nUrut = nI - 1
433                               nKeterangan = "Warranty "
434                               rsTemp.Find "KETERANGAN = 'nilai WARRANTY'", , adSearchForward, 1
435                               If Not rsTemp.EOF Then
436                                   If cekInsert(nUrut, nKeterangan) Then
437                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
438                                   Else
439                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
440                                   End If
441                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
442                               Else
443                                   If cekInsert(nUrut, nKeterangan) Then
444                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
445                                   Else
446                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
447                                   End If
448                                   nTotal = nTotal + 0
449                               End If
450                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Warranty ...")
451                           Case 22
452                               nUrut = nI - 1
453                               nKeterangan = "Return Job "
454                               rsTemp.Find "KETERANGAN = 'Return Job'", , adSearchForward, 1
455                               If Not rsTemp.EOF Then
456                                   If cekInsert(nUrut, nKeterangan) Then
457                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
458                                   Else
459                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
460                                   End If
461                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
462                               Else
463                                   If cekInsert(nUrut, nKeterangan) Then
464                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
465                                   Else
466                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
467                                   End If
468                                   nTotal = nTotal + 0
469                               End If
470                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data RTJ ...")
471                           Case 23
472                               nUrut = nI - 1
473                               nKeterangan = "Pre Delivery Servis "
474                               rsTemp.Find "KETERANGAN = 'PDS'", , adSearchForward, 1
475                               If Not rsTemp.EOF Then
476                                   If cekInsert(nUrut, nKeterangan) Then
477                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
478                                   Else
479                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
480                                   End If
481                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
482                               Else
483                                   If cekInsert(nUrut, nKeterangan) Then
484                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
485                                   Else
486                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
487                                   End If
488                                   nTotal = nTotal + 0
489                               End If
490                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data PDS ...")
491                           Case 24
                                  ' SUBTOTAL NONCPUS
492                               nUrut = nI - 1
493                               nKeterangan = "  SUBTOTAL NON CPUS "
494                               If cekInsert(nUrut, nKeterangan) Then
495                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
496                               Else
497                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
498                               End If
499                               nSubTotal = nSubTotal + nTotal
500                               nTotal = 0
501                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Sub Total Non CPUS ...")
                              'LAPORAN KPI BARU
                              
502                           Case 25
                                  'UNIT THS TOTAL
503                               nUrut = nI - 1
504                               nKeterangan = "Unit TMS Total "
505                               rsTemp.Find "KETERANGAN = 'THS'", , adSearchForward, 1
506                               If Not rsTemp.EOF Then
507                                   If cekInsert(nUrut, nKeterangan) Then
508                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
509                                   Else
510                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
511                                   End If
512                               Else
513                                   If cekInsert(nUrut, nKeterangan) Then
514                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
515                                   Else
516                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
517                                   End If
518                               End If
519                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data THS ...")
520                           Case 26
                                  'UNIT THS SBE
521                               nUrut = nI - 1
522                               nKeterangan = "Unit TMS SBE "
523                               rsTemp.Find "KETERANGAN = 'THS SBE'", , adSearchForward, 1
524                               If Not rsTemp.EOF Then
525                                   If cekInsert(nUrut, nKeterangan) Then
526                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
527                                   Else
528                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
529                                   End If
530                               Else
531                                   If cekInsert(nUrut, nKeterangan) Then
532                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
533                                   Else
534                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
535                                   End If
536                               End If
537                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data THS ...")
538                           Case 27
                                  'UNIT SP SBE
539                               nUrut = nI - 1
540                               nKeterangan = "Unit SP SBE "
541                               rsTemp.Find "KETERANGAN = 'Unit SP SBE'", , adSearchForward, 1
542                               If Not rsTemp.EOF Then
543                                   If cekInsert(nUrut, nKeterangan) Then
544                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
545                                   Else
546                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
547                                   End If
548                               Else
549                                   If cekInsert(nUrut, nKeterangan) Then
550                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
551                                   Else
552                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
553                                   End If
554                               End If
555                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Service Point ...")
556                           Case 28
                                  'UNIT SP SBE
557                               nUrut = nI - 1
558                               nKeterangan = "Unit SP Total "
559                               rsTemp.Find "KETERANGAN = 'Unit SP Total'", , adSearchForward, 1
560                               If Not rsTemp.EOF Then
561                                   If cekInsert(nUrut, nKeterangan) Then
562                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
563                                   Else
564                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
565                                   End If
566                               Else
567                                   If cekInsert(nUrut, nKeterangan) Then
568                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
569                                   Else
570                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
571                                   End If
572                               End If
573                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Service Point ...")
574                           Case 29
                                  'UNIT SP SBE
575                               nUrut = nI - 1
576                               nKeterangan = "Jumlah Service Point "
577                               rsTemp.Find "KETERANGAN = 'Jumlah Service Point'", , adSearchForward, 1
578                               If Not rsTemp.EOF Then
579                                   If cekInsert(nUrut, nKeterangan) Then
580                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
581                                   Else
582                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
583                                   End If
584                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
585                               Else
586                                   If cekInsert(nUrut, nKeterangan) Then
587                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
588                                   Else
589                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
590                                   End If
591                                   nTotal = nTotal + 0
592                               End If
593                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Service Point ...")
594                           Case 30
                                  'UNIT EM TOtal
595                               nUrut = nI - 1
596                               nKeterangan = "Unit EM Total "
597                               rsTemp.Find "KETERANGAN = 'Unit EM Total'", , adSearchForward, 1
598                               If Not rsTemp.EOF Then
599                                   If cekInsert(nUrut, nKeterangan) Then
600                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
601                                   Else
602                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
603                                   End If
604                               Else
605                                   If cekInsert(nUrut, nKeterangan) Then
606                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
607                                   Else
608                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
609                                   End If
610                               End If
611                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Service Point ...")
      '                        Case 31
      '                            nUrut = nI - 1
      '                            nKeterangan = "Jumlah Unit THS "
      '                            rsTemp.Find "KETERANGAN = 'THS'", , adSearchForward, 1
      '                            If Not rsTemp.EOF Then
      '                                If cekInsert(nUrut, nKeterangan) Then
      '                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
      '                                Else
      '                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
      '                                End If
      '                                totalUnitEntryTHS = CDbl(rsTemp("JUMLAH"))
      '                            Else
      '                                If cekInsert(nUrut, nKeterangan) Then
      '                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
      '                                Else
      '                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
      '                                End If
      '                                totalUnitEntryTHS = 0
      '                            End If
      '                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data THS ...")
612                           Case 31
                                  ' JUMLAH TOTAL UNIT
613                               nUrut = nI - 1
614                               nKeterangan = "  JUMLAH TOTAL UNIT "
615                               If cekInsert(nUrut, nKeterangan) Then
616                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 1, True)
617                               Else
618                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 2, True)
619                               End If
620                               totalUnitEntry = CDbl(nSubTotal)
621                               nSubTotal = 0
622                               nTotal = 0
623                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Total Unit ...")
624                           Case 32
625                               nUrut = nI - 1
626                               nKeterangan = "Pelanggan Booking yang datang pada hari Booking yg sama "
627                               rsTemp.Find "KETERANGAN = 'Booking Hari Sama'", , adSearchForward, 1
628                               If Not rsTemp.EOF Then
629                                   If cekInsert(nUrut, nKeterangan) Then
630                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
631                                   Else
632                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
633                                   End If
634                               Else
635                                   If cekInsert(nUrut, nKeterangan) Then
636                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
637                                   Else
638                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
639                                   End If
640                               End If
641                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Booking On Day ...")
642                           Case 33
643                               nUrut = nI - 1
644                               nKeterangan = "Labor "
645                               rsTemp.Find "KETERANGAN = 'Revenue Jasa'", , adSearchForward, 1
646                               If Not rsTemp.EOF Then
647                                   If cekInsert(nUrut, nKeterangan) Then
648                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
649                                   Else
650                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
651                                   End If
652                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
653                               Else
654                                   If cekInsert(nUrut, nKeterangan) Then
655                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
656                                   Else
657                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
658                                   End If
659                                   nTotal = nTotal + 0
660                               End If
661                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Revenue Jasa ...")
662                           Case 34
663                               nUrut = nI - 1
664                               nKeterangan = "Labor Non Ppn"
665                               rsTemp.Find "KETERANGAN = 'Revenue Jasa Non ppn'", , adSearchForward, 1
666                               If Not rsTemp.EOF Then
667                                   If cekInsert(nUrut, nKeterangan) Then
668                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
669                                   Else
670                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
671                                   End If
672                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
673                               Else
674                                   If cekInsert(nUrut, nKeterangan) Then
675                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
676                                   Else
677                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
678                                   End If
679                                   nTotal = nTotal + 0
680                               End If
681                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Revenue Non PPn ...")
682                           Case 35
683                               nUrut = nI - 1
684                               nKeterangan = "Labor Total (CPUS + Non CPUS) "
685                               rsTemp.Find "KETERANGAN = 'Revenue Jasa Total'", , adSearchForward, 1
686                               If Not rsTemp.EOF Then
687                                   If cekInsert(nUrut, nKeterangan) Then
688                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
689                                   Else
690                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
691                                   End If
692                                   nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
693                               Else
694                                   If cekInsert(nUrut, nKeterangan) Then
695                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
696                                   Else
697                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
698                                   End If
699                                   nTotal2 = nTotal2 + 0
700                               End If
701                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Revenue Total ...")
702                           Case 36
703                               nUrut = nI - 1
704                               nKeterangan = "Part Toyota & Oli TGMO " & keteranganFaktur
705                               rsTemp.Find "KETERANGAN = 'Revenue Part dan Oli TGMO'", , adSearchForward, 1
706                               If Not rsTemp.EOF Then
707                                   If cekInsert(nUrut, nKeterangan) Then
708                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
709                                   Else
710                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
711                                   End If
712                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
713                               Else
714                                   If cekInsert(nUrut, nKeterangan) Then
715                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
716                                   Else
717                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
718                                   End If
719                                   nTotal = nTotal + 0
720                               End If
721                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part & Oli Genuine ...")
722                           Case 37
723                               nUrut = nI - 1
724                               nKeterangan = "Part Toyota & Oli TGMO (CPUS + Non CPUS)" & keteranganFaktur
725                               rsTemp.Find "KETERANGAN = 'Revenue Part dan Oli TGMO Total'", , adSearchForward, 1
726                               If Not rsTemp.EOF Then
727                                   If cekInsert(nUrut, nKeterangan) Then
728                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
729                                   Else
730                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
731                                   End If
732                                   nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
733                               Else
734                                   If cekInsert(nUrut, nKeterangan) Then
735                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
736                                   Else
737                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
738                                   End If
739                                   nTotal2 = nTotal2 + 0
740                               End If
741                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part & Oli Genuine ...")
742                           Case 38
743                               nUrut = nI - 1
744                               nKeterangan = "Oli / Bahan Non Toyota " & keteranganFaktur
745                               rsTemp.Find "KETERANGAN = 'Revenue Oli dan Bahan Non Toyota'", , adSearchForward, 1
746                               If Not rsTemp.EOF Then
747                                   If cekInsert(nUrut, nKeterangan) Then
748                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
749                                   Else
750                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
751                                   End If
752                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
753                               Else
754                                   If cekInsert(nUrut, nKeterangan) Then
755                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
756                                   Else
757                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
758                                   End If
759                                   nTotal = nTotal + 0
760                               End If
761                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part & Oli Non Genuine ...")
762                           Case 39
763                               nUrut = nI - 1
764                               nKeterangan = "Oli / Bahan Non Toyota (CPUS + Non CPUS)" & keteranganFaktur
765                               rsTemp.Find "KETERANGAN = 'Revenue Oli dan Bahan Non Toyota Total'", , adSearchForward, 1
766                               If Not rsTemp.EOF Then
767                                   If cekInsert(nUrut, nKeterangan) Then
768                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
769                                   Else
770                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
771                                   End If
772                                   nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
773                               Else
774                                   If cekInsert(nUrut, nKeterangan) Then
775                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
776                                   Else
777                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
778                                   End If
779                                   nTotal2 = nTotal2 + 0
780                               End If
781                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part & Oli Non Genuine ...")
782                           Case 40
783                               nUrut = nI - 1
784                               nKeterangan = "Pekerjaan Luar "
785                               rsTemp.Find "KETERANGAN = 'Revenue Pekerjaan Luar'", , adSearchForward, 1
786                               If Not rsTemp.EOF Then
787                                   If cekInsert(nUrut, nKeterangan) Then
788                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
789                                   Else
790                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
791                                   End If
792                                   nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
793                               Else
794                                   If cekInsert(nUrut, nKeterangan) Then
795                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
796                                   Else
797                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
798                                   End If
799                                   nTotal = nTotal + 0
800                               End If
801                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Sublet ...")
802                           Case 41
803                               nUrut = nI - 1
804                               nKeterangan = "Pekerjaan Luar (CPUS + Non CPUS)"
805                               rsTemp.Find "KETERANGAN = 'Revenue Pekerjaan Luar Total'", , adSearchForward, 1
806                               If Not rsTemp.EOF Then
807                                   If cekInsert(nUrut, nKeterangan) Then
808                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
809                                   Else
810                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
811                                   End If
812                                   nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
813                               Else
814                                   If cekInsert(nUrut, nKeterangan) Then
815                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
816                                   Else
817                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
818                                   End If
819                                   nTotal2 = nTotal2 + 0
820                               End If
821                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Sublet ...")
822                           Case 42
                                  ' SUB TOTAL REVENUE CPUS
823                               nUrut = nI - 1
824                               nKeterangan = "  SUBTOTAL REVENUE CPUS "
825                               If cekInsert(nUrut, nKeterangan) Then
826                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
827                               Else
828                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
829                               End If
830                               nSubTotal = 0
831                               nTotal = 0
832                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Revenue CPUS ...")
833                           Case 43
                                  ' SUB TOTAL REVENUE
834                               nUrut = nI - 1
835                               nKeterangan = "  SUBTOTAL REVENUE CPUS + Non CPUS "
836                               If cekInsert(nUrut, nKeterangan) Then
837                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal2), 1, True)
838                               Else
839                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal2), 2, True)
840                               End If
841                               nSubTotal = 0
842                               nTotal2 = 0
843                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Revenue CPUS & Non CPUS ...")
                              

844                           Case 44
845                               nUrut = nI - 1
846                               nKeterangan = "Total Servis Berkala Eks Revenue "
847                               rsTemp.Find "KETERANGAN = 'Revenue SBE'", , adSearchForward, 1
848                               If Not rsTemp.EOF Then
849                                   If cekInsert(nUrut, nKeterangan) Then
850                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
851                                   Else
852                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
853                                   End If
854                               Else
855                                   If cekInsert(nUrut, nKeterangan) Then
856                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
857                                   Else
858                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
859                                   End If
860                               End If
861                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Revenue SBE ...")
                                  
862                           Case 45
                                  ' JML STALl EM
863                               nUrut = nI - 1
864                               nKeterangan = "Stall Express Maintenance (EM)"
865                               rsTemp.Find "KETERANGAN = 'EM'", , adSearchForward, 1
866                               If Not rsTemp.EOF Then
867                                   If cekInsert(nUrut, nKeterangan) Then
868                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
869                                   Else
870                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
871                                   End If
872                                   nSubTotal = CDbl(rsTemp("JUMLAH"))
873                               Else
874                                   If cekInsert(nUrut, nKeterangan) Then
875                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
876                                   Else
877                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
878                                   End If
879                                   nSubTotal = 0
880                               End If
881                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Stall ...")
882                           Case 46
                                  ' JML STALl Non EM ( Produktif )
883                               nUrut = nI - 1
884                               nKeterangan = "Stall Produktif Lainnya (Non EM)"
885                               rsTemp.Find "KETERANGAN = 'NonEM'", , adSearchForward, 1
886                               If Not rsTemp.EOF Then
887                                   If cekInsert(nUrut, nKeterangan) Then
888                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
889                                   Else
890                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
891                                   End If
892                                   nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
893                               Else
894                                   If cekInsert(nUrut, nKeterangan) Then
895                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
896                                   Else
897                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
898                                   End If
899                                   nSubTotal = nSubTotal + 0
900                               End If
901                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Stall ...")
902                           Case 47
                                  ' Sub Total Stall Produktif
903                               nUrut = nI - 1
904                               nKeterangan = "  SUBTOTAL STALL PRODUKTIF "
905                               If cekInsert(nUrut, nKeterangan) Then
906                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
907                               Else
908                                   sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
909                               End If
910                               nSubTotal = 0
911                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Stall ...")
912                           Case 48
                                  ' JML STALL Non Produktif
913                               nUrut = nI - 1
914                               nKeterangan = "Stall Non Produktif"
915                               rsTemp.Find "KETERANGAN = 'NonPro'", , adSearchForward, 1
916                               If Not rsTemp.EOF Then
917                                   If cekInsert(nUrut, nKeterangan) Then
918                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
919                                   Else
920                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
921                                   End If
922                               Else
923                                   If cekInsert(nUrut, nKeterangan) Then
924                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
925                                   Else
926                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
927                                   End If
928                               End If
929                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Stall ...")
930                           Case 49
                                  ' JML MOBIL THS
931                               nUrut = nI - 1
932                               nKeterangan = "Jumlah Kendaraan THS"
933                               rsTemp.Find "KETERANGAN = 'mobilTHS'", , adSearchForward, 1
934                               If Not rsTemp.EOF Then
935                                   If cekInsert(nUrut, nKeterangan) Then
936                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
937                                   Else
938                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
939                                   End If
940                               Else
941                                   If cekInsert(nUrut, nKeterangan) Then
942                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
943                                   Else
944                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
945                                   End If
946                               End If
947                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi THS ...")
                                  
948                           Case 50
                                  ' Man Power
949                               nUrut = nI - 1
950                               nKeterangan = "Teknisi PSG (Magang)"
951                               rsTemp.Find "KETERANGAN = 'PSG'", , adSearchForward, 1
952                               If Not rsTemp.EOF Then
953                                   If cekInsert(nUrut, nKeterangan) Then
954                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
955                                   Else
956                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
957                                   End If
958                                   nSubTotal = CDbl(rsTemp("JUMLAH"))
959                               Else
960                                   If cekInsert(nUrut, nKeterangan) Then
961                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
962                                   Else
963                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
964                                   End If
965                                   nSubTotal = 0
966                               End If
967                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
968                           Case 51
                                  ' Man Power
969                               nUrut = nI - 1
970                               nKeterangan = "Apprentice Technician"
971                               rsTemp.Find "KETERANGAN = 'APPMech'", , adSearchForward, 1
972                               If Not rsTemp.EOF Then
973                                   If cekInsert(nUrut, nKeterangan) Then
974                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
975                                   Else
976                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
977                                   End If
978                                   nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
979                               Else
980                                   If cekInsert(nUrut, nKeterangan) Then
981                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
982                                   Else
983                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
984                                   End If
985                                   nSubTotal = nSubTotal + 0
986                               End If
987                               Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
988                           Case 52
                                  ' Man Power
989                               nUrut = nI - 1
990                               nKeterangan = "Toyota Technician"
991                               rsTemp.Find "KETERANGAN = 'TTMech'", , adSearchForward, 1
992                               If Not rsTemp.EOF Then
993                                   If cekInsert(nUrut, nKeterangan) Then
994                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
995                                   Else
996                                       sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
997                                   End If
998                                   nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
999                               Else
1000                                  If cekInsert(nUrut, nKeterangan) Then
1001                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1002                                  Else
1003                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1004                                  End If
1005                                  nSubTotal = nSubTotal + 0
1006                              End If
1007                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
1008                          Case 53
                                  ' Man Power
1009                              nUrut = nI - 1
1010                              nKeterangan = "Pro Technician"
1011                              rsTemp.Find "KETERANGAN = 'PTMech'", , adSearchForward, 1
1012                              If Not rsTemp.EOF Then
1013                                  If cekInsert(nUrut, nKeterangan) Then
1014                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1015                                  Else
1016                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1017                                  End If
1018                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1019                              Else
1020                                  If cekInsert(nUrut, nKeterangan) Then
1021                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1022                                  Else
1023                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1024                                  End If
1025                                  nSubTotal = nSubTotal + 0
1026                              End If
1027                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
1028                          Case 54
                                  ' Man Power
1029                              nUrut = nI - 1
1030                              nKeterangan = "Diagnosa Technician"
1031                              rsTemp.Find "KETERANGAN = 'DTMech'", , adSearchForward, 1
1032                              If Not rsTemp.EOF Then
1033                                  If cekInsert(nUrut, nKeterangan) Then
1034                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1035                                  Else
1036                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1037                                  End If
1038                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1039                              Else
1040                                  If cekInsert(nUrut, nKeterangan) Then
1041                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1042                                  Else
1043                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1044                                  End If
1045                                  nSubTotal = nSubTotal + 0
1046                              End If
1047                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
1048                          Case 55
                                  ' Man Power
1049                              nUrut = nI - 1
1050                              nKeterangan = "Diagnosa Master Technician"
1051                              rsTemp.Find "KETERANGAN = 'DMTMech'", , adSearchForward, 1
1052                              If Not rsTemp.EOF Then
1053                                  If cekInsert(nUrut, nKeterangan) Then
1054                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1055                                  Else
1056                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1057                                  End If
1058                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1059                              Else
1060                                  If cekInsert(nUrut, nKeterangan) Then
1061                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1062                                  Else
1063                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1064                                  End If
1065                                  nSubTotal = nSubTotal + 0
1066                              End If
1067                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
1068                          Case 56
                                  ' Total Tech
1069                              nUrut = nI - 1
1070                              nKeterangan = "  TOTAL TEKNISI "
1071                              If cekInsert(nUrut, nKeterangan) Then
1072                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
1073                              Else
1074                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
1075                              End If
1076                              nTotal2 = nSubTotal
1077                              nSubTotal = 0
1078                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
                                  
                                  
1079                          Case 57
                                  ' Man Power
1080                              nUrut = nI - 1
1081                              nKeterangan = "App SA"
1082                              rsTemp.Find "KETERANGAN = 'APPSA'", , adSearchForward, 1
1083                              If Not rsTemp.EOF Then
1084                                  If cekInsert(nUrut, nKeterangan) Then
1085                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1086                                  Else
1087                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1088                                  End If
1089                                  nSubTotal = CDbl(rsTemp("JUMLAH"))
1090                              Else
1091                                  If cekInsert(nUrut, nKeterangan) Then
1092                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1093                                  Else
1094                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1095                                  End If
1096                                  nSubTotal = 0
1097                              End If
1098                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi SA ...")
1099                          Case 58
                                  ' Man Power
1100                              nUrut = nI - 1
1101                              nKeterangan = "TSA"
1102                              rsTemp.Find "KETERANGAN = 'TSA'", , adSearchForward, 1
1103                              If Not rsTemp.EOF Then
1104                                  If cekInsert(nUrut, nKeterangan) Then
1105                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1106                                  Else
1107                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1108                                  End If
1109                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1110                              Else
1111                                  If cekInsert(nUrut, nKeterangan) Then
1112                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1113                                  Else
1114                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1115                                  End If
1116                                  nSubTotal = nSubTotal + 0
1117                              End If
1118                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi SA ...")
1119                          Case 59
                                  ' Man Power
1120                              nUrut = nI - 1
1121                              nKeterangan = "PSA"
1122                              rsTemp.Find "KETERANGAN = 'PSA'", , adSearchForward, 1
1123                              If Not rsTemp.EOF Then
1124                                  If cekInsert(nUrut, nKeterangan) Then
1125                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1126                                  Else
1127                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1128                                  End If
1129                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1130                              Else
1131                                  If cekInsert(nUrut, nKeterangan) Then
1132                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1133                                  Else
1134                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1135                                  End If
1136                                  nSubTotal = nSubTotal + 0
1137                              End If
1138                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi SA ...")
1139                          Case 60
                                  ' Man Power
1140                              nUrut = nI - 1
1141                              nKeterangan = "MSA"
1142                              rsTemp.Find "KETERANGAN = 'MSA'", , adSearchForward, 1
1143                              If Not rsTemp.EOF Then
1144                                  If cekInsert(nUrut, nKeterangan) Then
1145                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1146                                  Else
1147                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1148                                  End If
1149                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1150                              Else
1151                                  If cekInsert(nUrut, nKeterangan) Then
1152                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1153                                  Else
1154                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1155                                  End If
1156                                  nSubTotal = nSubTotal + 0
1157                              End If
1158                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi SA ...")
1159                          Case 61
                                  ' Total Tech
1160                              nUrut = nI - 1
1161                              nKeterangan = "  TOTAL SA"
1162                              If cekInsert(nUrut, nKeterangan) Then
1163                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
1164                              Else
1165                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
1166                              End If
1167                              nTotal2 = nTotal2 + nSubTotal
1168                              nSubTotal = 0
1169                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi SA ...")
                                  
                                  
                                  
1170                          Case 62
                                  ' Man Power
1171                              nUrut = nI - 1
1172                              nKeterangan = "Toyota Foreman Training"
1173                              rsTemp.Find "KETERANGAN = 'TFT'", , adSearchForward, 1
1174                              If Not rsTemp.EOF Then
1175                                  If cekInsert(nUrut, nKeterangan) Then
1176                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1177                                  Else
1178                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1179                                  End If
1180                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1181                              Else
1182                                  If cekInsert(nUrut, nKeterangan) Then
1183                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1184                                  Else
1185                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1186                                  End If
1187                                  nSubTotal = nSubTotal + 0
1188                              End If
1189                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Foreman ...")
1190                          Case 63
                                  ' Man Power
1191                              nUrut = nI - 1
1192                              nKeterangan = "Non Toyota Foreman Training"
1193                              rsTemp.Find "KETERANGAN = 'NonTFT'", , adSearchForward, 1
1194                              If Not rsTemp.EOF Then
1195                                  If cekInsert(nUrut, nKeterangan) Then
1196                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1197                                  Else
1198                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1199                                  End If
1200                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1201                              Else
1202                                  If cekInsert(nUrut, nKeterangan) Then
1203                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1204                                  Else
1205                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1206                                  End If
1207                                  nSubTotal = nSubTotal + 0
1208                              End If
1209                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Foreman ...")
1210                          Case 64
                                  ' Total FO
1211                              nUrut = nI - 1
1212                              nKeterangan = "  TOTAL FOREMAN "
1213                              If cekInsert(nUrut, nKeterangan) Then
1214                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
1215                              Else
1216                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
1217                              End If
1218                              nTotal2 = nTotal2 + nSubTotal
1219                              nSubTotal = 0
1220                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Foreman ...")
                                  
                                  
1221                          Case 65
                                  ' Man Power
1222                              nUrut = nI - 1
1223                              nKeterangan = "Team 21 PT dan Lainnya"
1224                              rsTemp.Find "KETERANGAN = 'T21 PT'", , adSearchForward, 1
1225                              If Not rsTemp.EOF Then
1226                                  If cekInsert(nUrut, nKeterangan) Then
1227                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1228                                  Else
1229                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1230                                  End If
1231                              Else
1232                                  If cekInsert(nUrut, nKeterangan) Then
1233                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1234                                  Else
1235                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1236                                  End If
1237                              End If
1238                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Foreman Team 21 ...")
1239                          Case 66
                                  ' Man Power
1240                              nUrut = nI - 1
1241                              nKeterangan = "Team 21 DT"
1242                              rsTemp.Find "KETERANGAN = 'T21 DT'", , adSearchForward, 1
1243                              If Not rsTemp.EOF Then
1244                                  If cekInsert(nUrut, nKeterangan) Then
1245                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1246                                  Else
1247                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1248                                  End If
1249                              Else
1250                                  If cekInsert(nUrut, nKeterangan) Then
1251                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1252                                  Else
1253                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1254                                  End If
1255                              End If
1256                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Foreman Team 21 ...")
1257                          Case 67
                                  ' Man Power
1258                              nUrut = nI - 1
1259                              nKeterangan = "Team 21 DMT"
1260                              rsTemp.Find "KETERANGAN = 'T21 DMT'", , adSearchForward, 1
1261                              If Not rsTemp.EOF Then
1262                                  If cekInsert(nUrut, nKeterangan) Then
1263                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1264                                  Else
1265                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1266                                  End If
1267                              Else
1268                                  If cekInsert(nUrut, nKeterangan) Then
1269                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1270                                  Else
1271                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1272                                  End If
1273                              End If
1274                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Foreman Team 21 ...")
1275                              nSubTotal = 0
                          
                              
1276                          Case 68
                                  ' Man Power
1277                              nUrut = nI - 1
1278                              nKeterangan = "APP Partman"
1279                              rsTemp.Find "KETERANGAN = 'APPpart'", , adSearchForward, 1
1280                              If Not rsTemp.EOF Then
1281                                  If cekInsert(nUrut, nKeterangan) Then
1282                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1283                                  Else
1284                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1285                                  End If
1286                                  nSubTotal = CDbl(rsTemp("JUMLAH"))
1287                              Else
1288                                  If cekInsert(nUrut, nKeterangan) Then
1289                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1290                                  Else
1291                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1292                                  End If
1293                                  nSubTotal = 0
1294                              End If
1295                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Partman ...")
1296                          Case 69
                                  ' Man Power
1297                              nUrut = nI - 1
1298                              nKeterangan = "Partman Level 1"
1299                              rsTemp.Find "KETERANGAN = 'Part1'", , adSearchForward, 1
1300                              If Not rsTemp.EOF Then
1301                                  If cekInsert(nUrut, nKeterangan) Then
1302                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1303                                  Else
1304                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1305                                  End If
1306                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1307                              Else
1308                                  If cekInsert(nUrut, nKeterangan) Then
1309                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1310                                  Else
1311                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1312                                  End If
1313                                  nSubTotal = nSubTotal + 0
1314                              End If
1315                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Partman ...")
1316                          Case 70
                                  ' Man Power
1317                              nUrut = nI - 1
1318                              nKeterangan = "Partman Level 2"
1319                              rsTemp.Find "KETERANGAN = 'Part2'", , adSearchForward, 1
1320                              If Not rsTemp.EOF Then
1321                                  If cekInsert(nUrut, nKeterangan) Then
1322                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1323                                  Else
1324                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1325                                  End If
1326                                  nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
1327                              Else
1328                                  If cekInsert(nUrut, nKeterangan) Then
1329                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1330                                  Else
1331                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1332                                  End If
1333                                  nSubTotal = nSubTotal + 0
1334                              End If
1335                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Partman ...")
1336                          Case 71
                                  ' Total Partman
1337                              nUrut = nI - 1
1338                              nKeterangan = "  TOTAL PARTMAN "
1339                              If cekInsert(nUrut, nKeterangan) Then
1340                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
1341                              Else
1342                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
1343                              End If
1344                              nTotal2 = nTotal2 + nSubTotal
1345                              nSubTotal = 0
1346                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Partman ...")
                                  
                                  
1347                          Case 72
                                  ' Man Power
1348                              nUrut = nI - 1
1349                              nKeterangan = "Personel Bengkel Lainnya"
1350                              rsTemp.Find "KETERANGAN = 'PBLain'", , adSearchForward, 1
1351                              If Not rsTemp.EOF Then
1352                                  If cekInsert(nUrut, nKeterangan) Then
1353                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1354                                  Else
1355                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1356                                  End If
1357                                  nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
1358                              Else
1359                                  If cekInsert(nUrut, nKeterangan) Then
1360                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1361                                  Else
1362                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1363                                  End If
1364                              End If
1365                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Personel ...")
1366                          Case 73
                                  ' Total Personel ALL
1367                              nUrut = nI - 1
1368                              nKeterangan = "  TOTAL PERSONEL"
1369                              If cekInsert(nUrut, nKeterangan) Then
1370                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nTotal2, 1, True)
1371                              Else
1372                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nTotal2, 2, True)
1373                              End If
1374                              nTotal2 = 0
1375                              nSubTotal = 0
1376                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Komposisi Personel ...")
                              'LAPORAN KPI BARU
1377                          Case 74
1378                              nUrut = nI - 1
1379                              nKeterangan = "Fill Rate "
1380                              rsTemp.Find "Keterangan = 'Fill Rate'", , adSearchForward, 1
1381                              If Not rsTemp.EOF Then
1382                                  If cekInsert(nUrut, nKeterangan) Then
1383                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 1)
1384                                  Else
1385                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 2)
1386                                  End If
1387                              Else
1388                                  If cekInsert(nUrut, nKeterangan) Then
1389                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1390                                  Else
1391                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1392                                  End If
1393                              End If
1394                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1395                          Case 75
1396                              nUrut = nI - 1
1397                              nKeterangan = "Service Rate "
1398                              rsTemp.Find "Keterangan = 'Service Rate'", , adSearchForward, 1
1399                              If Not rsTemp.EOF Then
1400                                  If cekInsert(nUrut, nKeterangan) Then
1401                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 1)
1402                                  Else
1403                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 2)
1404                                  End If
1405                              Else
1406                                  If cekInsert(nUrut, nKeterangan) Then
1407                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1408                                  Else
1409                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1410                                  End If
1411                              End If
1412                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1413                          Case 76
1414                              nUrut = nI - 1
1415                              nKeterangan = "Stock Day (Satuannya Hari) "
1416                              rsTemp.Find "Keterangan = 'Stock Day'", , adSearchForward, 1
1417                              If Not rsTemp.EOF Then
1418                                  If cekInsert(nUrut, nKeterangan) Then
1419                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 1)
1420                                  Else
1421                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 2)
1422                                  End If
1423                              Else
1424                                  If cekInsert(nUrut, nKeterangan) Then
1425                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1426                                  Else
1427                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1428                                  End If
1429                              End If
1430                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1431                          Case 77
1432                              nUrut = nI - 1
1433                              nKeterangan = "TMO Mineral [Liter] "
1434                              rsTemp.Find "Keterangan = 'TMO Mineral'", , adSearchForward, 1
1435                              If Not rsTemp.EOF Then
1436                                  If cekInsert(nUrut, nKeterangan) Then
1437                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 1)
1438                                  Else
1439                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 2)
1440                                  End If
1441                                  subTotalTMO = subTotalTMO + Round(CDbl(rsTemp("JUMLAH")))
1442                              Else
1443                                  If cekInsert(nUrut, nKeterangan) Then
1444                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1445                                  Else
1446                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1447                                  End If
1448                                  subTotalTMO = subTotalTMO + 0
1449                              End If
1450                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1451                          Case 78
1452                              nUrut = nI - 1
1453                              nKeterangan = "TMO Syntetic [Liter] "
1454                              rsTemp.Find "Keterangan = 'TMO Syntetic'", , adSearchForward, 1
1455                              If Not rsTemp.EOF Then
1456                                  If cekInsert(nUrut, nKeterangan) Then
1457                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 1)
1458                                  Else
1459                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 2)
1460                                  End If
1461                                  subTotalTMO = subTotalTMO + Round(CDbl(rsTemp("JUMLAH")))
1462                              Else
1463                                  If cekInsert(nUrut, nKeterangan) Then
1464                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1465                                  Else
1466                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1467                                  End If
1468                                  subTotalTMO = subTotalTMO + 0
1469                              End If
1470                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1471                          Case 79
1472                              nUrut = nI - 1
1473                              nKeterangan = "TMO Full Syntetic [Liter] "
1474                              rsTemp.Find "Keterangan = 'TMO Full Syntetic'", , adSearchForward, 1
1475                              If Not rsTemp.EOF Then
1476                                  If cekInsert(nUrut, nKeterangan) Then
1477                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 1)
1478                                  Else
1479                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 2)
1480                                  End If
1481                                  subTotalTMO = subTotalTMO + Round(CDbl(rsTemp("JUMLAH")))
1482                              Else
1483                                  If cekInsert(nUrut, nKeterangan) Then
1484                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1485                                  Else
1486                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1487                                  End If
1488                                  subTotalTMO = subTotalTMO + 0
1489                              End If
1490                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1491                          Case 80
1492                              nUrut = nI - 1
1493                              nKeterangan = "Sub Total TMO (liter) "
1494                              If cekInsert(nUrut, nKeterangan) Then
1495                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, subTotalTMO, 1, True)
1496                              Else
1497                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, subTotalTMO, 2, True)
1498                              End If
                                  
1499                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1500                          Case 81
1501                              nUrut = nI - 1
1502                              nKeterangan = "TGGO [Liter] "
1503                              rsTemp.Find "Keterangan = 'TGGO'", , adSearchForward, 1
1504                              If Not rsTemp.EOF Then
1505                                  If cekInsert(nUrut, nKeterangan) Then
1506                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 1)
1507                                  Else
1508                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 2)
1509                                  End If
1510                              Else
1511                                  If cekInsert(nUrut, nKeterangan) Then
1512                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1513                                  Else
1514                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1515                                  End If
1516                              End If
1517                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1518                          Case 82
1519                              nUrut = nI - 1
1520                              nKeterangan = "Oli Mesin Non TOYOTA Mineral [Liter] "
1521                              rsTemp.Find "Keterangan = 'Mineral NOE'", , adSearchForward, 1
1522                              If Not rsTemp.EOF Then
1523                                  If cekInsert(nUrut, nKeterangan) Then
1524                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 1)
1525                                  Else
1526                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 2)
1527                                  End If
1528                                  oliMesinNonToyota = oliMesinNonToyota + Round(CDbl(rsTemp("JUMLAH")))
1529                              Else
1530                                  If cekInsert(nUrut, nKeterangan) Then
1531                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1532                                  Else
1533                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1534                                  End If
1535                                  oliMesinNonToyota = oliMesinNonToyota + 0
1536                              End If
1537                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1538                          Case 83
1539                              nUrut = nI - 1
1540                              nKeterangan = "Oli Mesin Non TOYOTA Syntetic [Liter] "
                                  'rsTemp.Find "Keterangan = 'Mineral NOE'", , adSearchForward, 1
1541                              If cekInsert(nUrut, nKeterangan) Then
1542                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1543                              Else
1544                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1545                              End If
1546                              oliMesinNonToyota = oliMesinNonToyota + 0
1547                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1548                          Case 84
1549                              nUrut = nI - 1
1550                              nKeterangan = "Oli Mesin Non TOYOTA [Liter] "
                                  'rsTemp.Find "Keterangan = 'Mineral NOE'", , adSearchForward, 1
1551                              If cekInsert(nUrut, nKeterangan) Then
1552                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, oliMesinNonToyota, 1, True)
1553                              Else
1554                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, oliMesinNonToyota, 2, True)
1555                              End If
1556                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1557                          Case 85
1558                              nUrut = nI - 1
1559                              nKeterangan = "Oli Transmisi & Diff [liter] "
1560                              rsTemp.Find "Keterangan = 'Oli Transmisi'", , adSearchForward, 1
1561                              If Not rsTemp.EOF Then
1562                                  If cekInsert(nUrut, nKeterangan) Then
1563                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 1, True)
1564                                  Else
1565                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH"))), 2, True)
1566                                  End If
1567                              Else
1568                                  If cekInsert(nUrut, nKeterangan) Then
1569                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1, True)
1570                                  Else
1571                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2, True)
1572                                  End If
1573                              End If
1574                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part ...")
1575                          Case 86 '69
1576                              nUrut = nI - 1
1577                              nKeterangan = "Jam Tersedia (Biasa & Lembur) "
1578                              rsTemp.Find "Keterangan = 'Jam Tersedia'", , adSearchForward, 1
1579                              If Not rsTemp.EOF Then
1580                                  If cekInsert(nUrut, nKeterangan) Then
1581                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH")) / 60, 2), 1)
1582                                  Else
1583                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH")) / 60, 2), 2)
1584                                  End If
1585                              Else
1586                                  If cekInsert(nUrut, nKeterangan) Then
1587                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1588                                  Else
1589                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1590                                  End If
1591                              End If
1592                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Service Time ...")
1593                          Case 87 '70
1594                              nUrut = nI - 1
1595                              nKeterangan = "Jam Terpakai "
1596                              rsTemp.Find "Keterangan = 'Jam Terpakai'", , adSearchForward, 1
1597                              If Not rsTemp.EOF Then
1598                                  If cekInsert(nUrut, nKeterangan) Then
1599                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1600                                  Else
1601                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1602                                  End If
1603                              Else
1604                                  If cekInsert(nUrut, nKeterangan) Then
1605                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1606                                  Else
1607                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1608                                  End If
1609                              End If
1610                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Service Time ...")
1611                          Case 88 '71
1612                              nUrut = nI - 1
1613                              nKeterangan = "Jam Terjual "
1614                              rsTemp.Find "Keterangan = 'Jam Terjual'", , adSearchForward, 1
1615                              If Not rsTemp.EOF Then
1616                                  If cekInsert(nUrut, nKeterangan) Then
1617                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1618                                  Else
1619                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1620                                  End If
1621                              Else
1622                                  If cekInsert(nUrut, nKeterangan) Then
1623                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1624                                  Else
1625                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1626                                  End If
1627                              End If
1628                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Service Time ...")
1629                          Case 89 '72
1630                              nUrut = nI
1631                              nHrsFollow = 0
1632                              nKeterangan = "Jumlah Unit yang harus di Follow Up "
1633                              rsTemp.Find "Keterangan = 'Unit hrs Follow'", , adSearchForward, 1
1634                              If Not rsTemp.EOF Then
1635                                  If cekInsert(nUrut, nKeterangan) Then
1636                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1637                                  Else
1638                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1639                                  End If
1640                                  nHrsFollow = CDbl(rsTemp("JUMLAH"))
1641                              Else
1642                                  If cekInsert(nUrut, nKeterangan) Then
1643                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1644                                  Else
1645                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1646                                  End If
1647                              End If
1648                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Follow Up ...")
1649                          Case 90 '73
1650                              nUrut = nI
1651                              nDiFollow = 0
1652                              nKeterangan = "Jumlah Unit yang berhasil di Follow Up "
1653                              rsTemp.Find "Keterangan = 'Unit di Follow'", , adSearchForward, 1
1654                              If Not rsTemp.EOF Then
1655                                  If cekInsert(nUrut, nKeterangan) Then
1656                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1657                                  Else
1658                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1659                                  End If
1660                                  nDiFollow = CDbl(rsTemp("JUMLAH"))
1661                              Else
1662                                  If cekInsert(nUrut, nKeterangan) Then
1663                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1664                                  Else
1665                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1666                                  End If
1667                              End If
1668                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Follow Up ...")
1669                          Case 91 '74
                                  ' PSFU Success Call Rate
1670                              nUrut = nI - 3
1671                              nKeterangan = "  PSFU Success Call Rate (%) "
1672                              nTotal = 0
1673                              If nHrsFollow > 0 Then
1674                                  nTotal = Round(nDiFollow / nHrsFollow, 2) * 100
1675                              End If
1676                              If cekInsert(nUrut, nKeterangan) Then
1677                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
1678                              Else
1679                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
1680                              End If
1681                              nTotal = 0
1682                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Follow Up ...")
1683                          Case 92 '75
1684                              nUrut = nI - 1
1685                              nKeterangan = "Jawaban Tidak untuk Q1 "
1686                              nNoQ1 = 0
1687                              rsTemp.Find "Keterangan = 'Unit NO Q1'", , adSearchForward, 1
1688                              If Not rsTemp.EOF Then
1689                                  If cekInsert(nUrut, nKeterangan) Then
1690                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1691                                  Else
1692                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1693                                  End If
1694                                  nNoQ1 = CDbl(rsTemp("JUMLAH"))
1695                              Else
1696                                  If cekInsert(nUrut, nKeterangan) Then
1697                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1698                                  Else
1699                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1700                                  End If
1701                              End If
1702                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Follow Up ...")
1703                          Case 93 '76
1704                              nUrut = nI - 1
1705                              nKeterangan = "Jawaban Tidak untuk Q2 "
1706                              nNoQ2 = 0
1707                              rsTemp.Find "Keterangan = 'Unit NO Q2'", , adSearchForward, 1
1708                              If Not rsTemp.EOF Then
1709                                  If cekInsert(nUrut, nKeterangan) Then
1710                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1711                                  Else
1712                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1713                                  End If
1714                                  nNoQ2 = CDbl(rsTemp("JUMLAH"))
1715                              Else
1716                                  If cekInsert(nUrut, nKeterangan) Then
1717                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1718                                  Else
1719                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1720                                  End If
1721                              End If
1722                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Follow Up ...")
1723                          Case 94 '77
1724                              nUrut = nI - 1
1725                              nKeterangan = "Jawaban Tidak untuk Q3 "
1726                              nNoQ3 = 0
1727                              rsTemp.Find "Keterangan = 'Unit NO Q3'", , adSearchForward, 1
1728                              If Not rsTemp.EOF Then
1729                                  If cekInsert(nUrut, nKeterangan) Then
1730                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1731                                  Else
1732                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1733                                  End If
1734                                  nNoQ3 = CDbl(rsTemp("JUMLAH"))
1735                              Else
1736                                  If cekInsert(nUrut, nKeterangan) Then
1737                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1738                                  Else
1739                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1740                                  End If
1741                              End If
1742                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Follow Up ...")
1743                          Case 95 '78
1744                              nUrut = nI - 1
1745                              nKeterangan = "Jumlah Unit RTJ "
1746                              nRTJ = 0
1747                              rsTemp.Find "Keterangan = 'Unit Follow Rtj'", , adSearchForward, 1
1748                              If Not rsTemp.EOF Then
1749                                  If cekInsert(nUrut, nKeterangan) Then
1750                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1751                                  Else
1752                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1753                                  End If
1754                                  nRTJ = CDbl(rsTemp("JUMLAH"))
1755                              Else
1756                                  If cekInsert(nUrut, nKeterangan) Then
1757                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1758                                  Else
1759                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1760                                  End If
1761                              End If
1762                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data RTJ Total ...")
1763                          Case 96 '79
                                  ' FIR Ratio (dari hasil PSFU)
1764                              nUrut = nI - 1
1765                              nKeterangan = "  FIR Ratio (dari hasil PSFU) (%) "
1766                              nTotal = 0
1767                              If nDiFollow > 0 Then
1768                                  nTotal = Round(1 - (Round(nNoQ1 / nDiFollow, 2) + Round(nNoQ2 / nDiFollow, 2) + Round(nNoQ3 / nDiFollow, 2)), 2) * 100
1769                              End If
1770                              If cekInsert(nUrut, nKeterangan) Then
1771                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
1772                              Else
1773                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
1774                              End If
1775                              nTotal = 0
1776                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data FIR ...")
1777                          Case 97 '80
1778                              nUrut = nI - 1
1779                              nKeterangan = "Non Fix It Right (Q1) (%) "
1780                              nTotal = 0
1781                              If nDiFollow > 0 Then
1782                                  nTotal = Round(nNoQ1 / nDiFollow, 2) * 100
1783                              End If
1784                              If cekInsert(nUrut, nKeterangan) Then
1785                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
1786                              Else
1787                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
1788                              End If
1789                              nTotal = 0
1790                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data FIR ...")
1791                          Case 98 '81
1792                              nUrut = nI - 1
1793                              nKeterangan = "Non On Time Delivery Rate (Q2) (%) "
1794                              nTotal = 0
1795                              If nDiFollow > 0 Then
1796                                  nTotal = Round(nNoQ2 / nDiFollow, 2) * 100
1797                              End If
1798                              If cekInsert(nUrut, nKeterangan) Then
1799                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
1800                              Else
1801                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
1802                              End If
1803                              nTotal = 0
1804                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data OTD ...")
1805                          Case 99 '82
1806                              nUrut = nI - 1
1807                              nKeterangan = "Non Reasonable Time (Q3) (%) "
1808                              nTotal = 0
1809                              If nDiFollow > 0 Then
1810                                  nTotal = Round(nNoQ3 / nDiFollow, 2) * 100
1811                              End If
1812                              If cekInsert(nUrut, nKeterangan) Then
1813                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
1814                              Else
1815                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
1816                              End If
1817                              nTotal = 0
1818                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data OTD ...")
1819                          Case 100 '83
1820                              nUrut = nI - 1
1821                              nKeterangan = "Gate To Gate"
1822                              If gHrSama = True Then
1823                                  rsTemp.Find "Keterangan = 'Lead Time Gate Total 1Day'", , adSearchForward, 1
1824                              Else
1825                                  rsTemp.Find "Keterangan = 'Lead Time Gate Total '", , adSearchForward, 1
1826                              End If
1827                              If Not rsTemp.EOF Then
1828                                  If cekInsert(nUrut, nKeterangan) Then
1829                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1830                                  Else
1831                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1832                                  End If
1833                                  GTG = CDbl(rsTemp("JUMLAH"))
1834                              Else
1835                                  If cekInsert(nUrut, nKeterangan) Then
1836                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1837                                  Else
1838                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1839                                  End If
1840                                  GTG = 0
1841                              End If
1842                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Gate To Gate ...")
1843                          Case 101 '84
1844                              nUrut = nI - 1
1845                              nKeterangan = "Unit Entry"
                                  Dim cUE As Double
                                  Dim cTHS As Double
                                  Dim nUE_THS As Double
1846                              If gHrSama = True Then
1847                                  rsTemp.Find "Keterangan = 'Unit Gate Total 1Day'", , adSearchForward, 1
1848                                  If Not rsTemp.EOF Then cUE = CDbl(rsTemp("JUMLAH"))
1849                                  rsTemp.Find "Keterangan = 'Unit Gate Total 1Day THS'", , adSearchForward, 1
1850                                  If Not rsTemp.EOF Then cTHS = CDbl(rsTemp("JUMLAH"))
1851                              Else
1852                                  cUE = totalUnitEntry
1853                                  cTHS = totalUnitEntryTHS
1854                              End If
                                  
1855                              jTHS = cTHS
1856                              nUE_THS = cUE - cTHS
                                  
                                  Dim cUENotif As Double
1857                              rsTemp.Find "Keterangan = 'Unit Notif Total 1Day'", , adSearchForward, 1
1858                              If Not rsTemp.EOF Then cUENotif = CDbl(rsTemp("JUMLAH"))
                                  Dim cUENotifths As Double
1859                              rsTemp.Find "Keterangan = 'Unit Notif Total 1Day THS'", , adSearchForward, 1
1860                              If Not rsTemp.EOF Then cUENotifths = CDbl(rsTemp("JUMLAH"))
                                  
                                  
1861                              If nUE_THS > 0 Then
1862                                  If cekInsert(nUrut, nKeterangan) Then
1863                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nUE_THS, 1)
1864                                      Call objO.UpdateData(sSQL)
1865                                      sSQL = insertTabel(namaBulan, 103, nKeterangan, (cUENotif - cUENotifths), 1)
1866                                      Call objO.UpdateData(sSQL)
1867                                      sSQL = insertTabel(namaBulan, 112, "Total Unit Entry non THS", (cUENotif - cUENotifths), 1)
1868                                  Else
1869                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nUE_THS, 2)
1870                                      Call objO.UpdateData(sSQL)
1871                                      sSQL = insertTabel(namaBulan, 103, nKeterangan, (cUENotif - cUENotifths), 2)
1872                                      Call objO.UpdateData(sSQL)
1873                                      sSQL = insertTabel(namaBulan, 112, "Total Unit Entry non THS", (cUENotif - cUENotifths), 2)
1874                                  End If
1875                                  UE = nUE_THS
1876                                  UEnT = cUENotif - cUENotifths
1877                              Else
1878                                  If cekInsert(nUrut, nKeterangan) Then
1879                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1880                                      Call objO.UpdateData(sSQL)
1881                                      sSQL = insertTabel(namaBulan, 103, nKeterangan, 0, 1)
1882                                      Call objO.UpdateData(sSQL)
1883                                      sSQL = insertTabel(namaBulan, 112, "Total Unit Entry non THS", 0, 1)
1884                                  Else
1885                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1886                                      Call objO.UpdateData(sSQL)
1887                                      sSQL = insertTabel(namaBulan, 103, nKeterangan, 0, 2)
1888                                      Call objO.UpdateData(sSQL)
1889                                      sSQL = insertTabel(namaBulan, 112, "Total Unit Entry non THS", 0, 2)
1890                                  End If
1891                                  UE = 0
1892                                  UEnT = 0
1893                              End If
1894                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data All Unit Entry ...")
1895                          Case 102 '85
                              'Lead Time Gate Total 1Day
1896                              nUrut = nI - 1
                                  Dim cGate As Double
1897                              rsTemp.Find "Keterangan = 'Lead Time Gate Total 1Day'", , adSearchForward, 1
1898                              If Not rsTemp.EOF Then
1899                                  cGate = CDbl(rsTemp("JUMLAH"))
1900                              Else
1901                                  cGate = 0
1902                              End If
1903                              nKeterangan = "Jumlah LT Gate to Gate (Visiting) (menit)"
                                  Dim GTGpUE As Double
1904                              If UE = 0 Then
1905                                  GTGpUE = 0
1906                              Else
1907                                  GTGpUE = Round((cGate / UE), 0)
1908                              End If
                                  
1909                              If cekInsert(nUrut, nKeterangan) Then
1910                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, GTGpUE, 1, True)
1911                              Else
1912                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, GTGpUE, 2, True)
1913                              End If
1914                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Lead Time ...")
1915                          Case 103 '86
1916                              nUrut = nI - 1
1917                              nKeterangan = "Produksi Notifikasi"
1918                              If gHrSama = True Then
1919                                  rsTemp.Find "Keterangan = 'Lead Time Notif Total 1Day'", , adSearchForward, 1
1920                              Else
1921                                  rsTemp.Find "Keterangan = 'Lead Time Notif Total'", , adSearchForward, 1
1922                              End If
1923                              If Not rsTemp.EOF Then
1924                                  If cekInsert(nUrut, nKeterangan) Then
1925                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1926                                  Else
1927                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1928                                  End If
1929                                  PN = CDbl(rsTemp("JUMLAH"))
1930                              Else
1931                                  If cekInsert(nUrut, nKeterangan) Then
1932                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1933                                  Else
1934                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1935                                  End If
1936                                  PN = 0
1937                              End If
1938                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Lead Time ...")
1939                          Case 104 '87
1940                              nUrut = nI
                                  Dim cLead As Double
1941                              rsTemp.Find "Keterangan = 'Lead Time Notif Total 1Day'", , adSearchForward, 1
1942                              If Not rsTemp.EOF Then
1943                                  cLead = CDbl(rsTemp("JUMLAH"))
1944                              Else
1945                                  cLead = 0
1946                              End If
1947                              nKeterangan = "Jumlah Lead Time (menit)"
                                  Dim LDpUE As Double
1948                              If UE = 0 Then
1949                                  LDpUE = 0
1950                              Else
1951                                  LDpUE = Round((cLead / UEnT), 0)
1952                              End If
                                  
1953                              If cekInsert(nUrut, nKeterangan) Then
1954                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, LDpUE, 1, True)
1955                              Else
1956                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, LDpUE, 2, True)
1957                              End If
1958                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Lead Time ...")
1959                          Case 105 '88
1960                              nUrut = nI
1961                              nKeterangan = "Gate To Gate SBE"
1962                              If gHrSama = True Then
1963                                  rsTemp.Find "Keterangan = 'Lead Time Gate SBE Total 1Day'", , adSearchForward, 1
1964                              Else
1965                                  rsTemp.Find "Keterangan = 'Lead Time SBE Gate Total '", , adSearchForward, 1
1966                              End If
1967                              If Not rsTemp.EOF Then
1968                                  If cekInsert(nUrut, nKeterangan) Then
1969                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
1970                                  Else
1971                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
1972                                  End If
1973                                  GTGs = CDbl(rsTemp("JUMLAH"))
1974                              Else
1975                                  If cekInsert(nUrut, nKeterangan) Then
1976                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
1977                                  Else
1978                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
1979                                  End If
1980                                  GTGs = 0
1981                              End If
1982                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Gate To Gate SBE ...")
1983                          Case 106 '89
1984                              nUrut = nI
1985                              nKeterangan = "Unit SBE Non THS"
                                  Dim cUESBE As Double
                                  Dim jTHS1 As Double
                                  Dim nUESBE_THS As Double
                                  Dim nUESBE_THSnotif As Double
                                  
                                  
                                  Dim cUENotifSB As Double
1986                              rsTemp.Find "Keterangan = 'Unit Notif SBE Total 1Day'", , adSearchForward, 1
1987                              If Not rsTemp.EOF Then cUENotifSB = CDbl(rsTemp("JUMLAH"))
                                  
                                  Dim cUENotifSBths As Double
1988                              rsTemp.Find "Keterangan = 'Unit Notif SBE Total 1Day THS'", , adSearchForward, 1
1989                              If Not rsTemp.EOF Then cUENotifSBths = CDbl(rsTemp("JUMLAH"))
                                  
1990                              nUESBE_THSnotif = cUENotifSB - cUENotifSBths
                                      
1991                              If gHrSama = True Then
1992                                  rsTemp.Find "Keterangan = 'Unit Gate SBE Total 1Day'", , adSearchForward, 1
1993                                  If Not rsTemp.EOF Then cUESBE = CDbl(rsTemp("JUMLAH"))
1994                                  rsTemp.Find "Keterangan = 'Unit Gate SBE Total 1Day THS'", , adSearchForward, 1
1995                                  If Not rsTemp.EOF Then jTHS1 = CDbl(rsTemp("JUMLAH"))
1996                                  nUESBE_THS = cUESBE - jTHS1
1997                              Else
1998                                  cUESBE = TotalUnitSBE
1999                                  nUESBE_THS = cUESBE - jTHS
2000                              End If
                                  
2001                              If nUESBE_THS > 0 Or nUESBE_THSnotif > 0 Then
2002                                  If cekInsert(nUrut, nKeterangan) Then
2003                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nUESBE_THS, 1)
2004                                      Call objO.UpdateData(sSQL)
2005                                      sSQL = insertTabel(namaBulan, 109, nKeterangan, nUESBE_THSnotif, 1)
2006                                  Else
2007                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nUESBE_THS, 2)
2008                                      Call objO.UpdateData(sSQL)
2009                                      sSQL = insertTabel(namaBulan, 109, nKeterangan, nUESBE_THSnotif, 2)
2010                                  End If
2011                                  UEs = nUESBE_THS
2012                                  UEsTHS = nUESBE_THSnotif
2013                              Else
2014                                  If cekInsert(nUrut, nKeterangan) Then
2015                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
2016                                      Call objO.UpdateData(sSQL)
2017                                      sSQL = insertTabel(namaBulan, 109, nKeterangan, 0, 1)
2018                                  Else
2019                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
2020                                      Call objO.UpdateData(sSQL)
2021                                      sSQL = insertTabel(namaBulan, 109, nKeterangan, 0, 2)
2022                                  End If
2023                                  UEs = 0
2024                                  UEsTHS = 0
2025                              End If
2026                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data SBE Non THS ...")
2027                          Case 107 '90
2028                              nUrut = nI
                                  Dim jGTGSBE As Double
                                  Dim GTGspUEs As Double
2029                              nKeterangan = "Jumlah LT Gate to Gate SBE (menit)"
2030                              rsTemp.Find "Keterangan = 'Lead Time Gate SBE Total 1Day'", , adSearchForward, 1
2031                              If Not rsTemp.EOF Then jGTGSBE = CDbl(rsTemp("JUMLAH"))
2032                              If UEs = 0 Then
2033                                  GTGspUEs = 0
2034                              Else
2035                                  GTGspUEs = Round((jGTGSBE / UEs), 0)
2036                              End If
2037                              If cekInsert(nUrut, nKeterangan) Then
2038                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, GTGspUEs, 1, True)
2039                              Else
2040                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, GTGspUEs, 2, True)
2041                              End If
2042                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Gate To Gate Lead Time ...")
2043                          Case 108 '91
2044                              nUrut = nI
2045                              nKeterangan = "Produksi Notifikasi SBE"
2046                              If gHrSama = True Then
2047                                  rsTemp.Find "Keterangan = 'Lead Time Notif SBE Total 1Day'", , adSearchForward, 1
2048                              Else
2049                                  rsTemp.Find "Keterangan = 'Lead Time SBE Notif Total '", , adSearchForward, 1
2050                              End If
2051                              If Not rsTemp.EOF Then
2052                                  If cekInsert(nUrut, nKeterangan) Then
2053                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
2054                                  Else
2055                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
2056                                  End If
2057                                  PNs = CDbl(rsTemp("JUMLAH"))
2058                              Else
2059                                  If cekInsert(nUrut, nKeterangan) Then
2060                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
2061                                  Else
2062                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
2063                                  End If
2064                                  PNs = 0
2065                              End If
2066                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Lead Time ...")
2067                         Case 109 '92
2068                              nUrut = nI + 1
2069                              nKeterangan = "Jumlah Lead Time SBE (menit)"
                                  Dim PNspUEs As Double
                                  Dim cPNs As Double
2070                              rsTemp.Find "Keterangan = 'Lead Time Notif SBE Total 1Day'", , adSearchForward, 1
2071                              If Not rsTemp.EOF Then
2072                                  cPNs = CDbl(rsTemp("JUMLAH"))
2073                              Else
2074                                  cPNs = 0
2075                              End If
2076                              If UEsTHS = 0 Then
2077                                  PNspUEs = 0
2078                              Else
2079                                  PNspUEs = Round((cPNs / UEsTHS), 0)
2080                              End If
                                  
2081                              If cekInsert(nUrut, nKeterangan) Then
2082                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, PNspUEs, 1, True)
2083                              Else
2084                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, PNspUEs, 2, True)
2085                              End If
2086                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Lead Time ...")
2087                          Case 110 '93
2088                              nUrut = nI + 1
2089                              nKeterangan = "Time Delivery (Notifikasi)"
2090                              rsTemp.Find "Keterangan = 'Time Delivery Notifikasi '", , adSearchForward, 1
2091                              If Not rsTemp.EOF Then
2092                                  If cekInsert(nUrut, nKeterangan) Then
2093                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
2094                                  Else
2095                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
2096                                  End If
2097                                  TDN = CDbl(rsTemp("JUMLAH"))
2098                              Else
2099                                  If cekInsert(nUrut, nKeterangan) Then
2100                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
2101                                  Else
2102                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
2103                                  End If
2104                                  TDN = 0
2105                              End If
2106                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data OTD Notification ...")
2107                          Case 111 '94
2108                              nUrut = nI + 2
2109                              nKeterangan = "On Time Delivery (%)"
                                  Dim xOTD As Double
                                  'If totalUnitEntry - totalUnitEntryTHS > 0 Then
                                  '    xOTD = Round((TDN / (totalUnitEntry - totalUnitEntryTHS)) * 100, 1)
                                  'Else
                                  '    xOTD = 0
                                  'End If
2110                              xOTD = objO.FetchData_Double("select (((Select Count(A.Slisih) from " & _
                                          "(SELECT CASE WHEN DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), " & _
                                          "ISNULL(TGLPANGGIL,'')) <= 0 AND DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), ISNULL(TGLPANGGIL,'')) >= -30 THEN 1 ELSE 0 END AS SLISIH " & _
                                          "FROM JOBDTA A " & _
                                          "INNER JOIN JOBDTL B ON A.NO_WO = B.NO_WO " & _
                                          "LEFT JOIN " & Log.LogQueue & "..FINISHSERVICE C ON A.NO_POLISI = C.NOPOL AND A.INVDATE = C.TGLFINISH AND PANGGIL <> 0 " & _
                                          "WHERE NOINV <> '' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND PDS = 0 AND THS = 0 AND TPSVC <> 'W' AND " & _
                                          "MONTH(DATE) = " & i & " AND YEAR(DATE) = " & Text(0) & ") A Where A.Slisih = 1)*100) " & _
                                          "/(Select Count(A.Slisih) from " & _
                                          "(SELECT CASE WHEN DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), " & _
                                          "ISNULL(TGLPANGGIL,'')) <= 0 AND DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), ISNULL(TGLPANGGIL,'')) >= -30 THEN 1 ELSE 0 END AS SLISIH " & _
                                          "FROM JOBDTA A " & _
                                          "INNER JOIN JOBDTL B ON A.NO_WO = B.NO_WO " & _
                                          "LEFT JOIN " & Log.LogQueue & "..FINISHSERVICE C ON A.NO_POLISI = C.NOPOL AND A.INVDATE = C.TGLFINISH AND PANGGIL <> 0 " & _
                                          "WHERE NOINV <> '' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND PDS = 0 AND THS = 0 AND TPSVC <> 'W' AND " & _
                                          "MONTH(DATE) = " & i & " AND YEAR(DATE) = " & Text(0) & ") A)) as OTD")

2111                              If cekInsert(nUrut, nKeterangan) Then
2112                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, xOTD, 1, True)
2113                              Else
2114                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, xOTD, 2, True)
2115                              End If
2116                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data OTD (%) ...")
2117                          Case 112 '95
2118                              nUrut = nI + 2
2119                              nKeterangan = "Jumlah pelanggan yang diundang (aktual diundang) "
2120                              nTotal = 0
2121                              nTotal = CDbl(kpi6_b(CDbl(i), Trim(Text(0))))
2122                              If cekInsert(nUrut, nKeterangan) Then
2123                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
2124                              Else
2125                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
2126                              End If
2127                              nDiUndang = CDbl(nTotal)
2128                              nTotal = 0
2129                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Pelanggan ...")
2130                          Case 113 '96
2131                              nUrut = nI + 2
2132                              nKeterangan = "Jumlah pelanggan yang seharusnya diundang "
2133                              nTotal = 0
2134                              nTotal = CDbl(kpi6_a(CDbl(i), Trim(Text(0))))
2135                              If cekInsert(nUrut, nKeterangan) Then
2136                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
2137                              Else
2138                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
2139                              End If
2140                              nHrsUndang = CDbl(nTotal)
2141                              nTotal = 0
2142                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Pelanggan ...")
2143                          Case 114 '97
2144                              nUrut = nI + 2
2145                              nKeterangan = "Jumlah pelanggan yang datang setelah diundang "
2146                              nTotal = 0
2147                              nTotal = CDbl(kpi9_a(CDbl(i), Trim(Text(0))))
2148                              If cekInsert(nUrut, nKeterangan) Then
2149                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
2150                              Else
2151                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
2152                              End If
2153                              nDtgUndang = CDbl(nTotal)
2154                              nTotal = 0
2155                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Pelanggan ...")
2156                          Case 115 '98
2157                              nUrut = nI + 2
2158                              nKeterangan = "Total SBE SBI"
                                  Dim oSBE As Double
                                  Dim oSBI As Double
2159                              oSBI = Jumlah1K
2160                              oSBE = totalUnitEntry
                                  Dim oSBESBI As Double
2161                              oSBESBI = oSBE + oSBI
2162                              If cekInsert(nUrut, nKeterangan) Then
2163                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, oSBESBI, 1)
2164                              Else
2165                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, oSBESBI, 2)
2166                              End If
2167                              SBESBI = oSBESBI
2168                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Total SB ...")
2169                          Case 116 '99
2170                              nUrut = nI + 2
2171                              nKeterangan = "Kontribusi MRS - Come After Invited (%)"
                                  Dim nDtgUndangpSBESBI As Double
2172                              If SBESBI = 0 Then
2173                                  nDtgUndangpSBESBI = 0
2174                              Else
2175                                  nDtgUndangpSBESBI = Round((nDtgUndang / SBESBI) * 100, 0)
2176                              End If
                                  
2177                              If cekInsert(nUrut, nKeterangan) Then
2178                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nDtgUndangpSBESBI, 1, True)
2179                              Else
2180                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nDtgUndangpSBESBI, 2, True)
2181                              End If
2182                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data MRS - Come After Invited...")

2183                          Case 117 '100   'acep
                                  ''Set rsTemp = obj.KPI_5(True, i, CInt(Text(0)))
2184                              Set rsTemp = objKPI.KPI_5(True, i, CInt(Text(0)), "")
2185                              nUrut = nI + 2
2186                              nKeterangan = "Jumlah Telepon Terhubung (1) (Tt)"
                                  'rsTemp.Find "KETERANGAN = 'KPI_5A'", , adSearchForward, 1
2187                              If Not rsTemp.EOF Then
2188                                  If cekInsert(nUrut, nKeterangan) Then
2189                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("xx")), 1)
2190                                  Else
2191                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("xx")), 2)
2192                                  End If
2193                                  TTh = CDbl(rsTemp("xx"))
2194                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Kontribusi MRS ...")
2195                              Else
2196                                  If cekInsert(nUrut, nKeterangan) Then
2197                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
2198                                  Else
2199                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
2200                                  End If
2201                                  TTh = 0
2202                              End If
2203                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data MRS - Call ...")
2204                          Case 118 '101
2205                              nUrut = nI + 2
                                  ''Set rsTemp = obj.KPI_5(False, i, CInt(Text(0)))
2206                              Set rsTemp = objKPI.KPI_5(False, i, CInt(Text(0)), "")
2207                              nKeterangan = "Pelanggan yang ditelepon (SBE+SBI)"
2208                              If Not rsTemp.EOF Then
2209                                  If cekInsert(nUrut, nKeterangan) Then
2210                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("xx")), 1)
2211                                  Else
2212                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("xx")), 2)
2213                                  End If
2214                                  PTh = CDbl(rsTemp("xx"))
2215                              Else
2216                                  If cekInsert(nUrut, nKeterangan) Then
2217                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
2218                                  Else
2219                                      sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
2220                                  End If
2221                                  PTh = 0
2222                              End If
2223                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data MRS - Call (SBE & SBI) ...")
2224                          Case 119 '102
2225                              nUrut = nI + 2
2226                              nKeterangan = "Database Valid (%)"
                                  Dim TThpPTh As Double
2227                              If PTh = 0 Then
2228                                  TThpPTh = 0
2229                              Else
2230                                  TThpPTh = Round((TTh / PTh) * 100, 0)
2231                              End If
                                  
2232                              If cekInsert(nUrut, nKeterangan) Then
2233                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, TThpPTh, 1, True)
2234                              Else
2235                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, TThpPTh, 2, True)
2236                              End If
2237                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Database Valid ...")
2238                          Case 120
2239                              nUrut = nI + 1
2240                              nKeterangan = "KPI Hansa Renkei BP"
                                  
2241                              If cekInsert(nUrut, nKeterangan) Then
2242                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1, True)
2243                              Else
2244                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2, True)
2245                              End If
2246                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Leads SA To Sales ...")
2247                          Case 121
2248                              nUrut = nI + 1
2249                              nKeterangan = "Leads SA To Sales"
                                  
2250                              Call LeadSAToSales(Text(0).Text, i)
2251                              If cekInsert(nUrut, nKeterangan) Then
2252                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalLead2Sales, 1, False)
2253                              Else
2254                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalLead2Sales, 2, False)
2255                              End If
2256                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Leads SA To Sales ...")
2257                          Case 122
2258                              nUrut = nI + 1
2259                              nKeterangan = "DO SA to Sales"
                                  
2260                              Call DOSAToSales(Text(0).Text, i)
                                  
2261                              If cekInsert(nUrut, nKeterangan) Then
2262                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalDOSAtoSales, 1, False)
2263                              Else
2264                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalDOSAtoSales, 2, False)
2265                              End If
2266                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data DO SA to Sales ...")
2267                          Case 123
2268                              nUrut = nI + 1
2269                              nKeterangan = "Booking Req By Sales"
                                   
2270                              Call BookingReqBySales(Text(0).Text, i)
                                  
2271                              If cekInsert(nUrut, nKeterangan) Then
2272                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalReqBookBySales, 1, False)
2273                              Else
2274                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalReqBookBySales, 2, False)
2275                              End If
2276                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Booking Req By Sales ...")
2277                          Case 124
2278                              nUrut = nI + 2
2279                              nKeterangan = "Booked Show By Sales"
                                  
2280                              Call BookedShowBySales(Text(0).Text, i)
                                  
2281                              If cekInsert(nUrut, nKeterangan) Then
2282                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalBookBySales, 1, False)
2283                              Else
2284                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalBookBySales, 2, False)
2285                              End If
2286                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Booked Show By Sales ...")
2287                          Case 125
2288                              nUrut = nI + 2
2289                              nKeterangan = "Contact Ratio By Sales (%)"
                                  Dim nilaiContactRatio As Double
2290                              nilaiContactRatio = 0
                                  
2291                              Call ContactRatioBySales(Text(0).Text, i)
                                  
2292                              If totalReminderSales > 0 Then
2293                                  nilaiContactRatio = Round((totalReminderTerhubungSales * 100) / totalReminderSales, 2)
2294                              End If
2295                              If cekInsert(nUrut, nKeterangan) Then
2296                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nilaiContactRatio, 1, False)
2297                              Else
2298                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nilaiContactRatio, 2, False)
2299                              End If
2300                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Leads SA To Sales ...")
2301                          Case 126
2302                              nUrut = nI + 2
2303                              nKeterangan = "Come After Invited By Sales (%) "
                                  Dim nilaiCAISales As Double
                                  
2304                              Call ComeAfterInvitedBySales(Text(0).Text, i)
                                  
2305                              If totalReminderSales > 0 Then
2306                                  nilaiCAISales = Round((totalInvoiceSales * 100) / totalReminderSales, 2)
2307                              End If
                                  
2308                              If cekInsert(nUrut, nKeterangan) Then
2309                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nilaiCAISales, 1, False)
2310                              Else
2311                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nilaiCAISales, 2, False)
2312                              End If
2313                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Come After Invited By Sales ...")
2314                          Case 127
2315                              nUrut = nI + 2
2316                              nKeterangan = "WO Contribution By BP - SA"
2317                              Call WOContributionByBP(Text(0).Text, i)
                                  
2318                              If cekInsert(nUrut, nKeterangan) Then
2319                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalWOBPtoGRP, 1, False)
2320                              Else
2321                                  sSQL = insertTabel(namaBulan, nUrut, nKeterangan, totalWOBPtoGRP, 2, False)
2322                              End If
2323                              Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Leads SA To Sales ...")
                                  
2324                      End Select
2325                      nI = nI + 1
2326                      Call objO.UpdateData(sSQL)
2327                  Loop
2328              End If
2329              rsTemp.Close
              'Next
2330          nBar.Min = 0
2331          nBar.Max = 10
2332          nBar.value = 0
2333          nLab.Caption = "Get Data Report : 0 Of 10"
2334          nPro.Caption = "Process :  1 Of 2"
2335          nKet.Caption = "..."
2336          PicBar.Visible = False
2337          DoEvents
2338          Shape1.Height = 615
2339          framePenjualanPartCounter.Visible = False
2340          Call DisableExport
2341          Call setLokasiTombol
2342          Call Cetak
2343          Call Transfer_JAF(CInt(Left(cboBulan1, InStr(cboBulan1, " |") - 1)), Text(0).Text)
2344      Else
2345          MsgBox pesan, vbCritical, "KELENGKAPAN"
2346          Exit Sub
2347      End If
End Sub


Private Sub tblSimpan_Click()
cboBulan2.Text = cboBulan1.Text
Text(1).Text = Text(0).Text
'Screen.MousePointer = vbHourglass
Dim pesan As String
Dim temp As String
Dim boleh As Boolean

    boleh = True
    If Text(0) = "" Then
        pesan = pesan + "Tahun Awal. "
        boleh = False
    End If
    If Text(1) = "" Then
        pesan = pesan + "Tahun Akhir. "
        boleh = False
    End If
    If CDbl(Text(0)) <> CDbl(Text(1)) Then
        pesan = pesan + "Tahun harus sama. "
        boleh = False
    End If
    If cboBulan1 = "" Then
        pesan = pesan + "Bulan Awal. "
        boleh = False
    End If
    If cboBulan2 = "" Then
        pesan = pesan + "Bulan Akhir. "
        boleh = False
    End If
    
    If boleh Then
        Shape1.Height = 1575
        framePenjualanPartCounter.Visible = True
        cmdPilihRepo.Enabled = True
        ckFakturA.Enabled = True
        ckFakturB.Enabled = True
        ckFakturC.Enabled = True
        ckFakturA.value = vbChecked
        ckFakturB.value = vbChecked
        ckFakturC.value = vbChecked
        
    Else
        MsgBox pesan, vbCritical, "KELENGKAPAN"
        Exit Sub
    End If
End Sub

Private Sub Cetak()
Dim oRs As Object
'Dim Report As New RptRepoBulanan
    Set Report = New RptRepoBulanan

    sSQL = "SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY urut"
    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    
    Report.txtPerusahaan.SetText gPerush
    Report.txtTahun.SetText CDbl(Text(0))
    Report.txtTanggal.SetText Mid(cboBulan1, InStr(cboBulan1, "| ") + 2, Len(cboBulan1) - InStr(cboBulan1, "| ") + 2) & CDbl(Text(0))
    
    j = CDbl(Left(cboBulan1, InStr(cboBulan1, " |") - 1))
    Report.txtBln1.SetText GetNamaBulan(j)
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub CetakByModel()
Dim oRs As Object
'Dim Report As New RptRepoBulananByModel
    Set Report = New RptRepoBulananByModel
    

    sSQL = "SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY urut"
    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    
    Report.txtPerusahaan.SetText gPerush
    Report.txtTahun.SetText CDbl(Text(2))
    Report.txtTanggal.SetText Mid(cboBulan3, InStr(cboBulan3, "| ") + 2, Len(cboBulan3) - InStr(cboBulan3, "| ") + 2) & CDbl(Text(2))
    
    j = CDbl(Left(cboBulan1, InStr(cboBulan3, " |") - 1))
    'Report.txtBln1.SetText GetNamaBulan(j)
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub

Private Sub Form_Resize()
    Me.WindowState = vbMaximized
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 60 '0
    If ScaleHeight > (asxBars1.Height + 120) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 120
    End If
    If ScaleWidth > (asxPowerBanner1.Width + 180) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 180
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Function jasaPpn(Bln As Integer, Thn As Integer) As Double
Dim rsSearch As ADODB.Recordset
Dim namaTabel As String
    
    jasaPpn = 0
    namaTabel = "REVENUE"
    Call IsiTempRepo(Bln, Thn, namaTabel, "")
    
    Set rsSearch = objO.FetchData("SELECT ISNULL(SUM(ISNULL(sbi+sbe+grp+twc+lub+pdi+was,0)),0) AS TOTJASA FROM [" & namaTabel & "] WHERE TIPEINV <> 'AFFNONPPN' AND SBI = 0 AND RTJ = 0 AND PDI = 0 AND TWC = 0")
    If Not rsSearch.EOF Then
        jasaPpn = CDbl(rsSearch("TOTJASA"))
    End If
End Function

Private Function jasaNonPpn(Bln As Integer, Thn As Integer) As Double
Dim rsHlp As ADODB.Recordset
Dim rsSearch As ADODB.Recordset
Dim namaTabel As String
    
    jasaNonPpn = 0
    namaTabel = "REVENUE"
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If rsHlp.EOF Then
       Call IsiTempRepo(Bln, Thn, namaTabel, "")
    End If
    
    Set rsSearch = objO.FetchData("SELECT ISNULL(SUM(ISNULL(sbi+sbe+grp+twc+lub+pdi+was,0)),0) AS TOTJASA FROM [" & namaTabel & "] WHERE TIPEINV = 'AFFNONPPN' AND SBI = 0 AND RTJ = 0 AND PDI = 0 AND TWC = 0")
    If Not rsSearch.EOF Then
        jasaNonPpn = CDbl(rsSearch("TOTJASA"))
    End If
End Function

Private Function jasaTotal(Bln As Integer, Thn As Integer) As Double
Dim rsHlp As ADODB.Recordset
Dim rsSearch As ADODB.Recordset
Dim namaTabel As String
    
    jasaTotal = 0
    namaTabel = "REVENUE"
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If rsHlp.EOF Then
       Call IsiTempRepo(Bln, Thn, namaTabel, "")
    End If
    
    Set rsSearch = objO.FetchData("SELECT ISNULL(SUM(ISNULL(sbi+sbe+grp+twc+lub+pdi+was,0)),0) AS TOTJASA FROM [" & namaTabel & "]")
    If Not rsSearch.EOF Then
        jasaTotal = CDbl(rsSearch("TOTJASA"))
    End If
End Function

Private Sub IsiTempRepo(Bln As Integer, Thn As Integer, vNamaTabel As String, txt As String)
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
      Dim hlp6 As String
      Dim hlp7 As String
      Dim rsHlp As ADODB.Recordset
      Dim rsPro As ADODB.Recordset
      Dim rsAdd As ADODB.Recordset
      Dim nNilai As Double
      Dim nDuoInv As Boolean
      Dim subletOnly As Boolean
      Dim noInv As String
      Dim nmCust As String
      Dim nStatus As String
      'Dim vNamaTabel As  String
          
          
1         If Log.RPLHistori = 0 Then
2             Set rsPro = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A LEFT JOIN CUSTDAT B ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE A.FS_LOG_DEL = '' AND MONTH(INVDATE) = '" & Bln & "' AND YEAR(INVDATE) = '" & Thn & "' ORDER BY NOINV")
3         Else
4             Set rsPro = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A LEFT JOIN CUSTDAT B ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM C ON C.FS_LOG_DEL = '' AND A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE A.FS_LOG_DEL = '' AND MONTH(INVDATE) = '" & Bln & "' AND YEAR(INVDATE) = '" & Thn & "' ORDER BY NOINV")
5         End If
                                   ' --:: 23 AGUSTUS 2010 ; DELIV <> 'BTL' AND ; Menampilkan yg Retur ::--
6         If Not rsPro.EOF Then
7             Call BuatTabelRepoInv(vNamaTabel)
8             Do While Not rsPro.EOF
9                 subletOnly = True
10                For i = 1 To 8
11                    hlp = "REP" & i
12                    If rsPro(hlp) <> "" Then
13                        subletOnly = False
14                        hlp2 = "KLPKERJA" & i
15                        hlp3 = "RTE" & i
16                        hlp4 = "MEKANIK" & i
17                        hlp5 = "OFF" & i
18                        hlp6 = "DISC" & i
19                        hlp7 = "FREEC" & i
20                        If createTemp(vNamaTabel) Then
21                            nDuoInv = False
22                            If rsPro("NDUOINV") = 1 Then
23                                nDuoInv = True
24                            End If
                              '#########################################################'
                              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                              '#########################################################'
25                            nStatus = rsPro("STATUS")
26                            nNilai = CDbl(rsPro(hlp5)) * (1 - (CDbl(rsPro(hlp6)) / 100))
27                            noInv = Trim(rsPro("NOINV"))
28                            nmCust = Trim(rsPro("CUSTOMER"))
                              
29                            If nStatus = "AFFNONPPN" Then
30                                nNilai = nNilai * (100 / 110)
31                            End If
                              
32                            Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
33                            If Not rsHlp.EOF Then
34                                Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                               & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsPro(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsPro(hlp2) = "SBI" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "SBE = SBE + '" & IIf(rsPro(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsPro(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "TWC = TWC + '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsPro(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "PDI = PDI + '" & IIf(rsPro(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsPro(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "BOR = '" & cekDuo(CDbl(rsPro("BODY")), noInv) & "', MATERIAL = '" & cekDuo(CDbl(rsPro("MTRL")), noInv) & "', PART = '" & cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noInv) & "', " _
                                               & "OLI = '" & cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noInv) & "', SUBLET = '" & cekDuo(CDbl(rsPro("SBLT")), noInv) & "', SUBTOT = '" & cekDuo(CDbl(rsPro("SUBTOT")), noInv) & "', " _
                                               & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "', HPP = '" & cekDuo(CDbl(rsPro("HPP")), noInv) & "' " _
                                               & "WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
35                            Else
36                                Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noInv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                               & "'0','" & Replace(CDbl(rsPro(hlp3)), ",", ".") & "', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '" & IIf(rsPro(hlp2) = "SBI" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsPro(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "','" & IIf(rsPro(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsPro(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "'" & cekDuo(CDbl(rsPro("BODY")), noInv) & "', '" & cekDuo(CDbl(rsPro("MTRL")), noInv) & "', '" & cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noInv) & "', '" & cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noInv) & "', '" & cekDuo(CDbl(rsPro("SBLT")), noInv) & "', '" & cekDuo(CDbl(rsPro("SUBTOT")), noInv) & "', " _
                                               & " '" & cekDuo(CDbl(rsPro("MET")), noInv, 1) & "', '" & CDbl(rsPro("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "', '" & cekDuo(CDbl(rsPro("HPP")), noInv) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
37                            End If
                              
38                        End If
39                    End If
40                Next
41                If subletOnly Then
42                    nDuoInv = False
43                    noInv = Trim(rsPro("NOINV"))
                      
44                    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
45                    If Not rsHlp.EOF Then
46                        Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                       & "JOR = '0', TIPEINV = '" & Trim(rsPro("STATUS")) & "', CUSTOMER = '" & Trim(rsPro("CUSTOMER")) & "', BOR = '" & CDbl(rsPro("BODY")) & "', MATERIAL = '" & CDbl(rsPro("MTRL")) & "', " _
                                       & "PART = '" & CDbl(rsPro("PART_G") + rsPro("PART_NG")) & "', OLI = '" & CDbl(rsPro("OIL_G") + rsPro("OIL_NG")) & "', SUBLET = '" & CDbl(rsPro("SBLT")) & "', SUBTOT = '" & CDbl(rsPro("SUBTOT")) & "', " _
                                       & "LABOR = '" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsPro("NOFAKTUR")) & "', HPP = '" & CDbl(rsPro("HPP")) & "' WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
47                    Else
48                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noInv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                       & "'0','0', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                       & "'" & CDbl(rsPro("BODY")) & "', '" & CDbl(rsPro("MTRL")) & "', '" & CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")) & "', '" & CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")) & "', '" & CDbl(rsPro("SBLT")) & "', '" & CDbl(rsPro("SUBTOT")) & "', " _
                                       & " '" & CDbl(rsPro("MET")) & "', '" & CDbl(rsPro("TOTALL")) & "', '" & Trim(rsPro("STATUS")) & "', '" & Trim(rsPro("CUSTOMER")) & "','" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & Trim(rsPro("NOFAKTUR")) & "','" & CDbl(rsPro("HPP")) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
49                    End If
50                End If
                  
                  ' DARI ADD JOB
51                Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND FS_LOG_DEL = '" & Trim(rsPro("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsPro("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
52                If Not rsAdd.EOF Then
53                    Do While Not rsAdd.EOF
54                        hlp2 = "KLPKERJA"
55                        hlp3 = "RTE"
56                        hlp4 = "MEKANIK"
57                        hlp5 = "OFFN"
58                        hlp6 = "FREEC"
                          
59                        nDuoInv = False
60                        If rsPro("NDUOINV") = 1 Then
61                            nDuoInv = True
62                        End If
                          
                          '#########################################################'
                          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                          '#########################################################'
63                        nStatus = rsPro("STATUS")
64                        nNilai = CDbl(rsAdd(hlp5)) * (1 - (IIf(Right(Trim(rsPro("NOINV")), 2) = "-B", IIf(rsAdd("FREEC"), 100, 0), rsAdd("DISC")) / 100))
65                        noInv = Trim(rsPro("NOINV"))
66                        nmCust = Trim(rsPro("CUSTOMER"))

67                        If Trim(nStatus) = "AFFNONPPN" Then
68                            nNilai = nNilai * (100 / 110)
69                        End If
                          
70                        Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
71                        If Not rsHlp.EOF Then
72                            objO.UpdateData ("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                           & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                           & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "TWC = TWC + '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "BOR = '" & cekDuo(CDbl(rsPro("BODY")), noInv) & "', MATERIAL = '" & cekDuo(CDbl(rsPro("MTRL")), noInv) & "', PART = '" & cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noInv) & "', " _
                                           & "OLI = '" & cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noInv) & "', SUBLET = '" & cekDuo(CDbl(rsPro("SBLT")), noInv) & "', SUBTOT = '" & CDbl(rsPro("SUBTOT")) & "', HPP = '" & cekDuo(CDbl(rsPro("HPP")), noInv) & "', " _
                                           & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "' " _
                                           & "WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
73                        Else
74                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noInv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                           & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "SBE" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsPro("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "LUB" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & cekDuo(CDbl(rsPro("BODY")), noInv) & "', '" & cekDuo(CDbl(rsPro("MTRL")), noInv) & "', '" & cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noInv) & "', '" & cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noInv) & "', '" & cekDuo(CDbl(rsPro("SBLT")), noInv) & "', '" & CDbl(rsPro("SUBTOT")) & "', " _
                                           & " '" & cekDuo(CDbl(rsPro("MET")), noInv, 1) & "', '" & CDbl(rsPro("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "','" & cekDuo(CDbl(rsPro("HPP")), noInv) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
75                        End If
                          
76                        rsAdd.MoveNext
77                    Loop
78                End If
79                rsPro.MoveNext
80            Loop
81            Call IsiTempRepoRetur(Bln, Thn, vNamaTabel, txt)
82        Else
83            Exit Sub
84        End If
End Sub

Private Sub IsiTempRepoRetur(Bln As Integer, Thn As Integer, vNamaTabel As String, txt As String)
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
      Dim hlp6 As String
      Dim hlp7 As String
      Dim rsHlp As ADODB.Recordset
      Dim rsPro As ADODB.Recordset
      Dim rsAdd As ADODB.Recordset
      Dim nNilai As Double
      Dim nDuoInv As Boolean
      Dim subletOnly As Boolean
      Dim noInv As String
      Dim nmCust As String
      Dim nStatus As String
          
1         If Log.RPLHistori = 0 Then
2             Set rsPro = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A LEFT JOIN CUSTDAT B ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE DELIV = 'BTL' AND A.FS_LOG_DEL = '' AND MONTH(INVDATE) = '" & Bln & "' AND YEAR(INVDATE) = '" & Thn & "' ORDER BY NOINV")
3         Else
4             Set rsPro = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A LEFT JOIN CUSTDAT B ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM C ON C.FS_LOG_DEL = '' AND A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE DELIV = 'BTL' AND A.FS_LOG_DEL = '' AND MONTH(INVDATE) = '" & Bln & "' AND YEAR(INVDATE) = '" & Thn & "' ORDER BY NOINV")
5         End If
6         If Not rsPro.EOF Then
7             If createTemp(vNamaTabel) Then
8             Do While Not rsPro.EOF
9                 subletOnly = True
10                For i = 1 To 8
11                    hlp = "REP" & i
12                    If rsPro(hlp) <> "" Then
13                        subletOnly = False
14                        hlp2 = "KLPKERJA" & i
15                        hlp3 = "RTE" & i
16                        hlp4 = "MEKANIK" & i
17                        hlp5 = "OFF" & i
18                        hlp6 = "DISC" & i
19                        hlp7 = "FREEC" & i
20                        If createTemp(vNamaTabel) Then
21                            nDuoInv = False
22                            If rsPro("NDUOINV") = 1 Then
23                                nDuoInv = True
24                            End If
                              '#########################################################'
                              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                              '#########################################################'
25                            nStatus = rsPro("STATUS")
26                            nNilai = -1 * CDbl(rsPro(hlp5)) * (1 - (CDbl(rsPro(hlp6)) / 100))
27                            noInv = Trim(rsPro("NOINV")) & "R"
28                            nmCust = Trim(rsPro("CUSTOMER"))
                              
29                            If nStatus = "AFFNONPPN" Then
30                                nNilai = nNilai * (100 / 110)
31                            End If
                              
32                            Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
33                            If Not rsHlp.EOF Then
34                                objO.UpdateData ("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                               & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsPro(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsPro(hlp2) = "SBI" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "SBE = SBE + '" & IIf(rsPro(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsPro(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "TWC = TWC + '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsPro(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "PDI = PDI + '" & IIf(rsPro(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsPro(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "BOR = '" & -1 * cekDuo(CDbl(rsPro("BODY")), noInv) & "', MATERIAL = '" & -1 * cekDuo(CDbl(rsPro("MTRL")), noInv) & "', PART = '" & -1 * cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noInv) & "', " _
                                               & "OLI = '" & -1 * cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noInv) & "', SUBLET = '" & -1 * cekDuo(CDbl(rsPro("SBLT")), noInv) & "', SUBTOT = '" & -1 * cekDuo(CDbl(rsPro("SUBTOT")), noInv) & "', " _
                                               & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "', HPP = '" & -1 * cekDuo(CDbl(rsPro("HPP")), noInv) & "' " _
                                               & "WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
35                            Else
36                                Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noInv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                               & "'0','" & Replace(CDbl(rsPro(hlp3)), ",", ".") & "', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '" & IIf(rsPro(hlp2) = "SBI" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsPro(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "','" & IIf(rsPro(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsPro(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                               & "'" & -1 * cekDuo(CDbl(rsPro("BODY")), noInv) & "', '" & -1 * cekDuo(CDbl(rsPro("MTRL")), noInv) & "', '" & -1 * cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noInv) & "', '" & -1 * cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noInv) & "', '" & -1 * cekDuo(CDbl(rsPro("SBLT")), noInv) & "', '" & -1 * cekDuo(CDbl(rsPro("SUBTOT")), noInv) & "', " _
                                               & " '" & -1 * cekDuo(CDbl(rsPro("MET")), noInv, 1) & "', '" & -1 * CDbl(rsPro("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "', '" & -1 * cekDuo(CDbl(rsPro("HPP")), noInv) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
37                            End If
                              
38                        End If
39                    End If
40                Next
41                If subletOnly Then
42                    nDuoInv = False
43                    If rsPro("NDUOINV") = 1 Then nDuoInv = True
                      '#########################################################'
                      '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                      '#########################################################'
44                    noInv = Trim(rsPro("NOINV")) & "R"
                      
45                    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
46                    If Not rsHlp.EOF Then
47                        Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                       & "JOR = '0', TIPEINV = '" & Trim(rsPro("STATUS")) & "', CUSTOMER = '" & Trim(rsPro("CUSTOMER")) & "', BOR = '" & -1 * CDbl(rsPro("BODY")) & "', MATERIAL = '" & -1 * CDbl(rsPro("MTRL")) & "', " _
                                       & "PART = '" & -1 * CDbl(rsPro("PART_G") + rsPro("PART_NG")) & "', OLI = '" & -1 * CDbl(rsPro("OIL_G") + rsPro("OIL_NG")) & "', SUBLET = '" & -1 * CDbl(rsPro("SBLT")) & "', SUBTOT = '" & -1 * CDbl(rsPro("SUBTOT")) & "', " _
                                       & "LABOR = '" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsPro("NOFAKTUR")) & "', HPP = '" & -1 * CDbl(rsPro("HPP")) & "' WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
48                    Else
49                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noInv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                       & "'0','0', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                       & "'" & -1 * CDbl(rsPro("BODY")) & "', '" & -1 * CDbl(rsPro("MTRL")) & "', '" & -1 * CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")) & "', '" & -1 * CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")) & "', '" & -1 * CDbl(rsPro("SBLT")) & "', '" & -1 * CDbl(rsPro("SUBTOT")) & "', " _
                                       & " '" & -1 * CDbl(rsPro("MET")) & "', '" & -1 * CDbl(rsPro("TOTALL")) & "', '" & Trim(rsPro("STATUS")) & "', '" & Trim(rsPro("CUSTOMER")) & "','" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & Trim(rsPro("NOFAKTUR")) & "','" & -1 * CDbl(rsPro("HPP")) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
50                    End If
51                End If
                  
                  ' DARI ADD JOB
52                Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND FS_LOG_DEL = '" & Trim(rsPro("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsPro("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
53                If Not rsAdd.EOF Then
54                    Do While Not rsAdd.EOF
55                        hlp2 = "KLPKERJA"
56                        hlp3 = "RTE"
57                        hlp4 = "MEKANIK"
58                        hlp5 = "OFFN"
59                        hlp6 = "FREEC"
                          
60                        nDuoInv = False
61                        If rsPro("NDUOINV") = 1 Then
62                            nDuoInv = True
63                        End If

64                        nStatus = rsPro("STATUS")
65                        nNilai = -1 * CDbl(rsAdd(hlp5)) * (1 - (IIf(Right(Trim(rsPro("NOINV")), 2) = "-B", IIf(rsAdd("FREEC"), 100, 0), rsAdd("DISC")) / 100))
66                        noInv = Trim(rsPro("NOINV")) & "R"
67                        nmCust = Trim(rsPro("CUSTOMER"))
                          
68                        If Trim(nStatus) = "AFFNONPPN" Then
69                            nNilai = nNilai * (100 / 110)
70                        End If
                          
71                        Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "'")
72                        If Not rsHlp.EOF Then
73                            Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                           & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                           & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "TWC = TWC + '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                           & "BOR = '" & -1 * cekDuo(CDbl(rsPro("BODY")), noInv) & "', MATERIAL = '" & -1 * cekDuo(CDbl(rsPro("MTRL")), noInv) & "', PART = '" & -1 * cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noInv) & "', " _
                                           & "OLI = '" & -1 * cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noInv) & "', SUBLET = '" & -1 * cekDuo(CDbl(rsPro("SBLT")), noInv) & "', SUBTOT = '" & -1 * CDbl(rsPro("SUBTOT")) & "', HPP = '" & -1 * cekDuo(CDbl(rsPro("HPP")), noInv) & "', " _
                                           & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "' " _
                                           & "WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
74                        Else
75                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noInv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                           & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "SBE" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsPro("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "LUB" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & -1 * cekDuo(CDbl(rsPro("BODY")), noInv) & "', '" & -1 * cekDuo(CDbl(rsPro("MTRL")), noInv) & "', '" & -1 * cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noInv) & "', '" & -1 * cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noInv) & "', '" & -1 * cekDuo(CDbl(rsPro("SBLT")), noInv) & "', '" & -1 * CDbl(rsPro("SUBTOT")) & "', " _
                                           & " '" & -1 * cekDuo(CDbl(rsPro("MET")), noInv, 1) & "', '" & -1 * CDbl(rsPro("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "','" & -1 * cekDuo(CDbl(rsPro("HPP")), noInv) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
76                        End If
                          
77                        rsAdd.MoveNext
78                    Loop
79                End If
80                rsPro.MoveNext
81            Loop
82            End If
83        End If
End Sub

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

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset
    createTemp = False
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelRepoInv(vNamaTabel)
        createTemp = True
    End If
End Function

Private Sub BuatTabelRepoInv(vNamaTabel As String)
Dim strSql As String

    strSql = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(strSql)

    strSql = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
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
         
    strSql = strSql + "[SUBLET][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SUBTOT][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [METERAI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TOTAL][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TIPEINV][varchar](20) NOT NULL DEFAULT '', " _
         & "    [CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
         & "    [LABOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [NOFAKTUR][varchar](30) NOT NULL DEFAULT ''," _
         & "    [HPP][decimal](18,2) NOT NULL DEFAULT 0," _
         & "    [INVDATE][datetime] NOT NULL DEFAULT '01/01/3000') ON [PRIMARY]"
    Call objO.UpdateData(strSql)
End Sub

Private Function totalUnitEntry(Bln As Integer, Thn As Integer) As Integer
Dim rsSearch As ADODB.Recordset

    totalUnitEntry = 0
    Set rsSearch = objO.FetchData("SELECT COUNT(*) AS JUMLAH FROM T_INVOICE WHERE MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND DELIV <> 'BTL' AND NOINV NOT LIKE '%-B' AND FS_LOG_DEL = ''")
    If Not rsSearch.EOF Then
        totalUnitEntry = CDbl(rsSearch("JUMLAH"))
    End If
    Set rsSearch = Nothing
    
End Function

Private Sub Transfer_JAF(Bln As Integer, Thn As Integer)
Dim rsTr As New ADODB.Recordset
Set rsTr = objO.FetchData("EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'")
    If Log.LogHOConn = 1 Then
        If ConHO = False Then
            Call Saving_Log_HO("EXEC REPOTAF", "EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'", "Bulan = " & Bln & " And Tahun = " & Thn)
        Else
            If Get_ConnectionProfil_Bengkel99 = False Then
                Call Saving_Log_HO("EXEC REPOTAF", "EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'", "Bulan = " & Bln & " And Tahun = " & Thn)
            Else
                Call objHO.OpenTransactionLog_NRM
                Call objHO.AddNew_WO("Delete Bengkel99..Jam_TAF WHere Bulan = " & Bln & " And Tahun = " & Thn & " And Cabang = '" & Log.LogNamaNas & "'")
                rsTr.MoveFirst
                Do Until rsTr.EOF
                    Call objHO.AddNew_WO("Insert Into Bengkel99..Jam_TAF Values (" & rsTr.Fields(0).value & ", " & rsTr.Fields(1).value & ",'" & rsTr.Fields(2).value & "','" & rsTr.Fields(3).value & "'," & RepRp(rsTr.Fields(4).value) & ")")
                    rsTr.MoveNext
                Loop
                Call objHO.CloseTransactionLog_NRM
                If errHO = True Then Call Saving_Log_HO("EXEC REPOTAF", "EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'", "Bulan = " & Bln & " And Tahun = " & Thn)
            End If
        End If
    Else
        Call Saving_Log_HO("EXEC REPOTAF", "EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'", "Bulan = " & Bln & " And Tahun = " & Thn)
    End If
End Sub

Private Sub Saving_Log_HO(nStatus As String, isSQl As String, nData As String)
    Call objHO.AddNew_LogHO(nData, "EXEC", nStatus, isSQl)
End Sub

Private Sub cmdPrintS_Click()
    Dim objSt As New clsStall
    Dim nOpt As Integer
    nOpt = IIf(Opt(0).value = True, 0, 1)
    Dim nTotalUnit(31) As Double
    
    Dim rsSt As ADODB.Recordset
    Dim nStNP As Double
    Dim nStEM As Double
    Dim nStGR As Double
    
    Dim nTot(31) As Double
    Dim nJA(31) As Double
    Dim nJF(31) As Double
    Dim nRPS(31) As Double
    Dim nLTr(31) As Double
    Dim nOPS(31) As Double
    Dim nDay(31) As Double
    Dim sSQLs As String
    Dim C As Integer
    
    PicBarEM.Visible = True
    nBarEM.Min = 0
    nBarEM.Max = Day(DateSerial(Val(cboThnS.Text), Val(cboBlnS.Text) + 1, 1 - 1))
    nBarEM.value = 0
    DoEvents
    
    For C = 1 To Day(DateSerial(Val(cboThnS.Text), Val(cboBlnS.Text) + 1, 1 - 1))
    
        nLabEM.Caption = "Colecting Data " & IIf(nOpt = 0, "NP,", "EM,") & "   Periode " & C & "/" & cboBlnS.Text & "/" & cboThnS.Text
        nBarEM.value = C - 1
        DoEvents
        
        Set rsSt = objSt.FetchQtyStall_aDay(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text))
        If rsSt.RecordCount = 0 Then
            nStNP = 0
            nStEM = 0
            nStGR = 0
        Else
            nStNP = rsSt.Fields(2).value
            nStEM = rsSt.Fields(1).value
            nStGR = rsSt.Fields(3).value
        End If
        nTot(C) = objSt.FetchTotalUnit_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt)
        nJA(C) = Round((objSt.FetchTotalJA_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt) / 60), 2) ' Dalam Jam
        nJF(C) = Round((objSt.FetchTotalJF_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt) / 60), 2) ' Dalam Jam
        If rsSt.RecordCount = 0 Or nTot(C) = 0 Then
            nRPS(C) = 0
        Else
            If nOpt = 0 Then
                If Val(rsSt.Fields(2).value) <> 0 Then
                    nRPS(C) = Round(nTot(C) / Val(rsSt.Fields(2).value), 2)
                Else
                    nRPS(C) = 0
                End If
            Else   'EM
                If Val(rsSt.Fields(1).value) <> 0 Then
                    nRPS(C) = Round(nTot(C) / Val(rsSt.Fields(1).value), 2)
                Else
                    nRPS(C) = 0
                End If
            End If
        End If
        If nTot(C) = 0 Or nJA(C) = 0 Then
            nLTr(C) = 0
        Else
            nLTr(C) = Round((nJA(C) * 60) / nTot(C), 2)
        End If
        If nJA(C) = 0 Or nJF(C) = 0 Then
            nOPS(C) = 0
        Else
            nOPS(C) = (Round(nJA(C) / nJF(C), 4) * 100)
        End If
        Select Case Weekday(DateSerial(Val(cboThnS.Text), Val(cboBlnS.Text), C))
        '1 = Minggu, 2 = Senin, 3 = Selasa, 4 = Rabu, 5 = Kamis, 6 = Jum’at, 7 = Sabtu
            Case 2, 3, 4, 5, 6
                nDay(C) = 7.25
            Case 7
                nDay(C) = 4.25
            Case 1
                nDay(C) = 0
        End Select
        If nOpt = 0 Then
            sSQLs = "Select * from LapNP_HO Where Tahun = " & Val(cboThnS.Text) & " And Bulan = " & Val(cboBlnS.Text) & " And Tanggal = " & C & ""
        Else
            sSQLs = "Select * from LapEM_HO Where Tahun = " & Val(cboThnS.Text) & " And Bulan = " & Val(cboBlnS.Text) & " And Tanggal = " & C & ""
        End If
        
        If objO.FetchData_Boolean(sSQLs) = True Then
            Call objSt.Update_Sumary_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt, nTot(C), nJA(C), nJF(C), nRPS(C), nLTr(C), nOPS(C), nDay(C), nStEM, nStNP, nStGR)
        Else
            Call objSt.AddNew_Sumary_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt, nTot(C), nJA(C), nJF(C), nRPS(C), nLTr(C), nOPS(C), nDay(C), nStEM, nStNP, nStGR)
        End If
    Next C
    PicBarEM.Visible = False
    DoEvents
    Call Print_NP_EM(nOpt, Val(cboThnS.Text), Val(cboBlnS.Text))
End Sub

Private Sub Print_NP_EM(isEM As Integer, isThn As Double, isBln As Double)
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New Rpt_NP_EM_HO
        
        sSQL = "select Tahun,Bulan,Tanggal," & IIf(isEM = 0, "StallNP", "StallEM") & " as Stall,StallGR,Tot,jA,JF,RPS,LTr,OPS,Jam from " & IIf(isEM = 0, "LapNP_HO", "LapEM_HO") & " Where " & _
               "Tahun = " & isThn & " And Bulan  = " & isBln & " Order By Tanggal ASC"

    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.Text2.SetText "----oo000oo----     SUMMARY  HARIAN  KPI  " & IIf(isEM = 0, "NP", "EM") & "     ----oo000oo----"
    Report.TxtAlamat.SetText gAlamat
    Report.Text13.SetText "Periode : " & isBln & " / " & isThn
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub


Private Sub cmdPrintNEA_Click()
    If dt0.Text = "__/__/____" Or dt1.Text = "__/__/____" Then Exit Sub
    nUnit = 0
    nUnitR = 0
    nINV = 0
    nInvR = 0
    noInv = ""
    Dim rsList As ADODB.Recordset
    Set rsList = objO.FetchData_LapNEA(dt0, dt1)
    createTbl_NEA True
    If rsList.RecordCount > 0 Then
        PicNEA.Visible = True
        nNEA.Caption = "Please Wait ... Processing Data NEA ..."
        BarNEA.Max = rsList.RecordCount + 1
        BarNEA.Min = 0
        BarNEA.value = 0
        C = 0
        DoEvents
        Do Until rsList.EOF
                C = C + 1
                BarNEA.value = C
                DoEvents
                rsNEA.AddNew
                rsNEA!noInv = Trim(rsList.Fields(0).value)
                
                Select Case Right(Trim(rsList.Fields(0).value), 1)
                    Case "R"
                        nInvR = nInvR + 1
                    Case Else
                        nINV = nINV + 1
                End Select
                
                If noInv <> Replace(Replace(Trim(rsList.Fields(0).value), "A", ""), "B", "") Then
                    noInv = Replace(Replace(Trim(rsList.Fields(0).value), "A", ""), "B", "")
                    If Right(Replace(Replace(Trim(rsList.Fields(0).value), "A", ""), "B", ""), 1) = "R" Then
                        nUnitR = nUnitR + 1
                    Else
                        nUnit = nUnit + 1
                    End If
                End If
                
                rsNEA!Tgl_inv = Left(Trim(rsList.Fields(1).value), Len(Trim(rsList.Fields(1).value)) - 3)
                rsNEA!no_wo = Trim(rsList.Fields(2).value)
                rsNEA!Tgl_WO = Left(Trim(rsList.Fields(3).value), Len(Trim(rsList.Fields(3).value)) - 3)
                rsNEA!NO_POLISI = Trim(rsList.Fields(4).value)
                rsNEA!Type = Trim(rsList.Fields(5).value)
                rsNEA!Problem_Keluhan = Replace(Replace(Replace(Trim(rsList.Fields(6).value), ", , ", ""), ", ,", ""), ",, ", ", ")
                
                Dim rsAdd As ADODB.Recordset
                Set rsAdd = objO.FetchData("select Rep from JobADD2 Where NO_WO = '" & Trim(rsList.Fields(2).value) & "'")
                If rsAdd.RecordCount > 0 Then
                    Dim nADD As String
                    nADD = Trim(rsList.Fields(7).value)
                    rsAdd.MoveFirst
                    Do Until rsAdd.EOF
                        nADD = nADD & ", " & Trim(rsAdd.Fields(0).value)
                        rsAdd.MoveNext
                    Loop
                    rsNEA!Solusi_Masalah = nADD
                Else
                    rsNEA!Solusi_Masalah = Trim(rsList.Fields(7).value)
                End If
                
                Dim rsSbl As ADODB.Recordset
                Set rsSbl = objO.FetchData("select namabarang from JobSublet Where NO_WO = '" & Trim(rsList.Fields(2).value) & "'")
                If rsSbl.RecordCount > 0 Then
                    Dim nSbl As String
                    nSbl = rsNEA!Solusi_Masalah
                    rsSbl.MoveFirst
                    Do Until rsSbl.EOF
                        nSbl = nSbl & ", Sblt:" & Trim(rsSbl.Fields(0).value)
                        rsSbl.MoveNext
                    Loop
                    rsNEA!Solusi_Masalah = nSbl
                End If
                rsNEA!Solusi_Masalah = Replace(Replace(Replace(rsNEA!Solusi_Masalah, ", , ", ""), ", ,", ""), ",, ", ", ")
                rsNEA!PIC = Trim(rsList.Fields(8).value)
                rsNEA!Mech = Trim(rsList.Fields(14).value)
                rsNEA!jasa = Trim(rsList.Fields(9).value)
                rsNEA!part = Trim(rsList.Fields(10).value)
                rsNEA!SUBLET = Trim(rsList.Fields(11).value)
                rsNEA!Revenue = Trim(rsList.Fields(12).value)
                rsNEA!Remark = Trim(rsList.Fields(13).value)
                rsNEA!namaCust = Trim(rsList.Fields(15).value)
                rsNEA.Update
            rsList.MoveNext
        Loop
        
        nNEA.Caption = "Please Wait ... Load Report NEA ..."
        C = C + 1
        BarNEA.value = C
        DoEvents
        PicNEA.Visible = False
        Call LoadReportNea
    Else
        MsgBox "Data Tidak Tersedia !", vbInformation + vbOKOnly, "Laporan NEA"
    End If
End Sub

Private Sub createTbl_NEA(Optional Edit As Boolean = False)
    If rsNEA.State Then rsNEA.Close
    Set rsNEA = Nothing
    With rsNEA
        .Fields.Append "Noinv", adVarChar, 25
        .Fields.Append "Tgl_inv", adVarChar, 20
        .Fields.Append "No_WO", adVarChar, 25
        .Fields.Append "Tgl_WO", adVarChar, 20
        .Fields.Append "No_Polisi", adVarChar, 15
        .Fields.Append "TYPE", adVarChar, 100
        .Fields.Append "Problem_Keluhan", adVarChar, 750
        .Fields.Append "Solusi_Masalah", adVarChar, 1000
        .Fields.Append "PIC", adVarChar, 10
        .Fields.Append "Mech", adVarChar, 10
        .Fields.Append "JASA", adDouble
        .Fields.Append "PART", adDouble
        .Fields.Append "Sublet", adDouble
        .Fields.Append "Revenue", adDouble
        .Fields.Append "Remark", adVarChar, 100
        .Fields.Append "NamaCust", adVarChar, 200
        .Open
    End With
End Sub

Private Sub LoadReportNea()
    Screen.MousePointer = vbHourglass
    Set Report = New RptNEA
    Report.database.SetDataSource rsNEA
    Report.txtPerusahaan.SetText gPerush
    Report.TxtAlamat.SetText gAlamat
    Report.Text13.SetText "Periode Invoice : " & dt0.Text & " - " & dt1.Text
    Report.txtKbg.SetText Log.LogKbgName
    
    Report.txtInv.SetText nINV
    Report.txtInvR.SetText nInvR
    Report.txtInvOK.SetText (nINV - nInvR)
    Report.txtUnit.SetText nUnit
    Report.txtunitR.SetText nUnitR
    Report.txtUnitOK.SetText (nUnit - nUnitR)
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub


Private Sub cmdPrintPS_Click()
1        On Error GoTo cmdPrintPS_Click_Error

2         If dtP0.Text = "__/__/____" Or dtP1.Text = "__/__/____" Then Exit Sub
3         nUnit = 0
4         nUnitR = 0
5         nINV = 0
6         nInvR = 0
7         noInv = ""
          Dim rsList As ADODB.Recordset
8         Set rsList = objO.FetchData_LapPS(dtP0, dtP1)
9         createTbl_PS True
10        If rsList.RecordCount > 0 Then
11            PicPS.Visible = True
12            nPS.Caption = "Please Wait ... Processing Data Profit Sharing ..."
13            BarPS.Max = rsList.RecordCount + 1
14            BarPS.Min = 0
15            BarPS.value = 0
16            C = 0
17            DoEvents
18            Do Until rsList.EOF
19                    C = C + 1
20                    BarPS.value = C
21                    DoEvents
22                    rsPS.AddNew
              
23                    rsPS!NO_POLISI = Trim(rsList.Fields(0).value)
24                    rsPS!Nama_Pelanggan = Trim(rsList.Fields(1).value)
25                    rsPS!Type = Trim(rsList.Fields(2).value)
26                    rsPS!no_wo = Trim(rsList.Fields(3).value)
27                    rsPS!noInv = Trim(rsList.Fields(4).value)

28                    Select Case Right(Trim(rsList.Fields(4).value), 1)
                          Case "R"
29                            nInvR = nInvR + 1
30                        Case Else
31                            nINV = nINV + 1
32                    End Select
                      
33                    If noInv <> Replace(Replace(Trim(rsList.Fields(4).value), "A", ""), "B", "") Then
34                        noInv = Replace(Replace(Trim(rsList.Fields(4).value), "A", ""), "B", "")
35                        If Right(Replace(Replace(Trim(rsList.Fields(4).value), "A", ""), "B", ""), 1) = "R" Then
36                            nUnitR = nUnitR + 1
37                        Else
38                            nUnit = nUnit + 1
39                        End If
40                    End If
                      
41                    rsPS!Tgl_inv = Left(Trim(rsList.Fields(5).value), (Len(Trim(rsList.Fields(5).value)) - 3))
                      
                      Dim rsAdd As ADODB.Recordset
42                    Set rsAdd = objO.FetchData("select Rep from JobADD2 Where NO_WO = '" & Trim(rsList.Fields(3).value) & "'")
43                    If rsAdd.RecordCount > 0 Then
                          Dim nADD As String
44                        nADD = Trim(rsList.Fields(6).value)
45                        rsAdd.MoveFirst
46                        Do Until rsAdd.EOF
47                            nADD = nADD & ", " & Trim(rsAdd.Fields(0).value)
48                            rsAdd.MoveNext
49                        Loop
50                        rsPS!Solusi_Masalah = nADD
51                    Else
52                        rsPS!Solusi_Masalah = Trim(rsList.Fields(6).value)
53                    End If
                      
                      Dim rsSbl As ADODB.Recordset
54                    Set rsSbl = objO.FetchData("select namabarang from JobSublet Where NO_WO = '" & Trim(rsList.Fields(3).value) & "'")
55                    If rsSbl.RecordCount > 0 Then
                          Dim nSbl As String
56                        nSbl = rsPS!Solusi_Masalah
57                        rsSbl.MoveFirst
58                        Do Until rsSbl.EOF
59                            nSbl = nSbl & ", Sblt:" & Trim(rsSbl.Fields(0).value)
60                            rsSbl.MoveNext
61                        Loop
62                        rsPS!Solusi_Masalah = nSbl
63                    End If
64                    rsPS!Solusi_Masalah = Replace(Replace(Replace(rsPS!Solusi_Masalah, ", , ", ""), ", ,", ""), ",, ", ", ")

65                    rsPS!Tgl_masuk = Trim(rsList.Fields(7).value)
66                    rsPS!Tgl_Mulai = Trim(objO.FetchData_String("select top 1 Tanggal + Jam from progress Where noOrder = '" & Trim(rsList.Fields(3).value) & "' and Flag = 'S' Order By Tanggal, Jam ASC"))
67                    rsPS!Tgl_Mulai = Left(rsPS!Tgl_Mulai, Len(rsPS!Tgl_Mulai) - 3)
68                    rsPS!Tgl_Selesai = Trim(objO.FetchData_String("select top 1 Tanggal + Jam from progress Where noOrder = '" & Trim(rsList.Fields(3).value) & "' and Flag = 'F' Order By Tanggal, Jam DESC"))
69                    rsPS!Tgl_Selesai = Left(rsPS!Tgl_Selesai, Len(rsPS!Tgl_Selesai) - 3)
                      
                      ''2015-11-18 || ganti koma mjd titik
70                    rsPS!jasa = Replace(Trim(rsList.Fields(8).value), ",", ".")
71                    rsPS!part = Replace(Trim(rsList.Fields(9).value), ",", ".")
72                    rsPS!SUBLET = Replace(Trim(rsList.Fields(10).value), ",", ".")
73                    rsPS!Revenue = Replace(Trim(rsList.Fields(11).value), ",", ".")
                      
74                    rsPS.Update
75                rsList.MoveNext
76            Loop
              
77            nPS.Caption = "Please Wait ... Load Report Profit Sharing ..."
78            C = C + 1
79            BarPS.value = C
80            DoEvents
81            PicPS.Visible = False
82            Call LoadReportPS
83        Else
84            MsgBox "Data Tidak Tersedia !", vbInformation + vbOKOnly, "Laporan PS"
85        End If

         
86       Exit Sub

cmdPrintPS_Click_Error:

87        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdPrintPS_Click of Form frmRepoBulanan")
End Sub

Private Sub createTbl_PS(Optional Edit As Boolean = False)
    If rsPS.State Then rsPS.Close
    Set rsPS = Nothing
    With rsPS
        .Fields.Append "No_Polisi", adVarChar, 15
        .Fields.Append "Nama_Pelanggan", adVarChar, 200
        .Fields.Append "TYPE", adVarChar, 100
        .Fields.Append "No_WO", adVarChar, 25
        .Fields.Append "Noinv", adVarChar, 25
        .Fields.Append "Tgl_inv", adVarChar, 50
        .Fields.Append "Solusi_Masalah", adVarChar, 1000
        .Fields.Append "Tgl_masuk", adVarChar, 50
        .Fields.Append "Tgl_Mulai", adVarChar, 50
        .Fields.Append "Tgl_Selesai", adVarChar, 50
        .Fields.Append "JASA", adDouble
        .Fields.Append "PART", adDouble
        .Fields.Append "Sublet", adDouble
        .Fields.Append "Revenue", adDouble
        .Open
    End With
End Sub

Private Sub LoadReportPS()
    Screen.MousePointer = vbHourglass
    Set Report = New RptProfitSharing
    Report.database.SetDataSource rsPS
    Report.txtPerusahaan.SetText gPerush
    Report.TxtAlamat.SetText gAlamat
    Report.Text13.SetText "Periode Invoice : " & dtP0.Text & " - " & dtP1.Text
    Report.txtKbg.SetText Log.LogKbgName
        
    Report.txtInv.SetText nINV
    Report.txtInvR.SetText nInvR
    Report.txtInvOK.SetText (nINV - nInvR)
    Report.txtUnit.SetText nUnit
    Report.txtunitR.SetText nUnitR
    Report.txtUnitOK.SetText (nUnit - nUnitR)
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
    Call DisableExport
    Call setLokasiTombol
End Sub

Private Sub cmdPrintP_Click()
    Screen.MousePointer = vbHourglass
    nWaitPart.Visible = True
    DoEvents
    Dim rsPart As ADODB.Recordset
    Set rsPart = objR.FetchLapPart(cboThnP.Text, cboBlnP.Text)
    Set Report = New RptBulPart
    Report.database.SetDataSource rsPart
    Report.txtPerusahaan.SetText gPerush
    Report.TxtAlamat.SetText gAlamat
    Report.Text2.SetText "Periode Laporan : " & GetNamaBulan(cboBlnP.Text) & " " & cboThnP.Text
        
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
    nWaitPart.Visible = False
    Call DisableExport
    Call setLokasiTombol
End Sub


Private Sub cmdClose_Click()
    FrRate.Visible = False
End Sub

Private Sub kosongNilai()
    'total sbe
    nTotSBE1 = 0
    nTotSBE2 = 0
    nTotSBE3 = 0
    nTotSBE4 = 0
    nTotSBE5 = 0
    nTotSBE6 = 0
    nTotSBE7 = 0
    nTotSBE8 = 0
    nTotSBE9 = 0
    nTotSBE10 = 0
    nTotSBE11 = 0
    nTotSBE12 = 0
    nTotSBE13 = 0
    nTotSBE14 = 0
    nTotSBE15 = 0
    nTotSBE16 = 0
    nTotSBE17 = 0
    nTotSBE18 = 0
    nTotSBE19 = 0
    nTotSBE20 = 0
    nTotSBE21 = 0
    
    'total grp
    nTotGR1 = 0
    nTotGR2 = 0
    nTotGR3 = 0
    nTotGR4 = 0
    nTotGR5 = 0
    nTotGR6 = 0
    nTotGR7 = 0
    nTotGR8 = 0
    nTotGR9 = 0
    nTotGR10 = 0
    nTotGR11 = 0
    nTotGR12 = 0
    nTotGR13 = 0
    nTotGR14 = 0
    nTotGR15 = 0
    nTotGR16 = 0
    nTotGR17 = 0
    nTotGR18 = 0
    nTotGR19 = 0
    nTotGR20 = 0
    nTotGR21 = 0
    
    totCPUS1 = 0
    totCPUS2 = 0
    totCPUS3 = 0
    totCPUS4 = 0
    totCPUS5 = 0
    totCPUS6 = 0
    totCPUS7 = 0
    totCPUS8 = 0
    totCPUS9 = 0
    totCPUS10 = 0
    totCPUS11 = 0
    totCPUS12 = 0
    totCPUS13 = 0
    totCPUS14 = 0
    totCPUS15 = 0
    totCPUS16 = 0
    totCPUS17 = 0
    totCPUS18 = 0
    totCPUS19 = 0
    totCPUS20 = 0
    totCPUS21 = 0
    
    nonCPUS1 = 0
    nonCPUS2 = 0
    nonCPUS3 = 0
    nonCPUS4 = 0
    nonCPUS5 = 0
    nonCPUS6 = 0
    nonCPUS7 = 0
    nonCPUS8 = 0
    nonCPUS9 = 0
    nonCPUS10 = 0
    nonCPUS11 = 0
    nonCPUS12 = 0
    nonCPUS13 = 0
    nonCPUS14 = 0
    nonCPUS15 = 0
    nonCPUS16 = 0
    nonCPUS17 = 0
    nonCPUS18 = 0
    nonCPUS19 = 0
    nonCPUS20 = 0
    nonCPUS21 = 0
    
End Sub


Private Sub XPButton1_Click()
Screen.MousePointer = vbHourglass
Dim pesan As String
Dim temp As String
Dim boleh As Boolean
Dim nI As Integer
Dim nTotal As Double
Dim nTotal2 As Double
Dim nSubTotal As Double
Dim namaBulan As String
Dim nUrut As Integer
Dim nKeterangan As String
''========================''

Dim nilJasaPPn As Double
Dim nilJasaNonPPn As Double
Dim nilJasaTotal As Double

''========================''
Dim nHrsFollow As Integer
Dim nDiFollow As Integer
Dim nNoQ1 As Integer
Dim nNoQ2 As Integer
Dim nNoQ3 As Integer
Dim nRTJ As Integer

''========================''
Dim nHrsUndang  As Integer
Dim nDiUndang   As Integer
Dim nDtgUndang  As Integer

''========================''
Dim GTG         As Double
Dim UE          As Double
Dim UEnT        As Double
Dim PN          As Double
Dim GTGs        As Double
Dim UEs         As Double
Dim UEsTHS      As Double
Dim PNs         As Double
Dim TDN         As Double
Dim SBESBI      As Double
Dim PTh         As Double
Dim TTh         As Double
Dim jTHS        As Double
Dim TotalSBE    As Double
Dim totalUnitEntryTHS As Double
Dim totalUnitEntry As Double
Dim TotalUnitSBE As Double
Dim Jumlah1K As Double
Dim i As Integer
Dim strSql As String
Dim hasil As String

    boleh = True
    
    If Text(2) = "" Then
        pesan = pesan + "Tahun Report. "
        boleh = False
    End If
    
    If cboBulan3 = "" Then
        pesan = pesan + "Bulan Report. "
        boleh = False
    End If
    
    If boleh Then
        Shape1.Height = 1575
    
        Call kosongNilai
        
        sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
             & " drop table [tempdb]..[" & vNamaTabel & "]"
        Call objO.UpdateData(sSQL)
        
        temp = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ([URUT][NUMERIC](9) NOT NULL DEFAULT 0, [KETERANGAN][varchar](100) NOT NULL DEFAULT '', [BOLD][char](1) NOT NULL DEFAULT '0' "
        For i = Left(cboBulan1, InStr(cboBulan1, " |") - 1) To Left(cboBulan2, InStr(cboBulan2, " |") - 1)
            temp = temp + ", [" & GetNamaBulan(i) & "][money] NOT NULL DEFAULT 0"
        Next
        temp = temp & " ,[AGYA][MONEY]NOT NULL DEFAULT 0, " _
                    & " [ALPHARD][MONEY]NOT NULL DEFAULT 0, " _
                    & " [AVANZA][MONEY]NOT NULL DEFAULT 0, " _
                    & " [CAMRY][MONEY]NOT NULL DEFAULT 0, " _
                    & " [COROLLA][MONEY]NOT NULL DEFAULT 0, " _
                    & " [DYNA][MONEY]NOT NULL DEFAULT 0, " _
                    & " [ETIOS][MONEY]NOT NULL DEFAULT 0, " _
                    & " [FJCRUISER][MONEY]NOT NULL DEFAULT 0, " _
                    & " [FORTUNER][MONEY]NOT NULL DEFAULT 0, " _
                    & " [FT86][MONEY]NOT NULL DEFAULT 0, " _
                    & " [HILUX][MONEY]NOT NULL DEFAULT 0, " _
                    & " [INNOVA][MONEY]NOT NULL DEFAULT 0, " _
                    & " [LANDCRUISER][MONEY]NOT NULL DEFAULT 0, " _
                    & " [LIMO][MONEY]NOT NULL DEFAULT 0, " _
                    & " [NAV1][MONEY]NOT NULL DEFAULT 0, " _
                    & " [PRIUS][MONEY]NOT NULL DEFAULT 0, " _
                    & " [RUSH][MONEY]NOT NULL DEFAULT 0, " _
                    & " [HIACE][MONEY]NOT NULL DEFAULT 0, " _
                    & " [VIOS][MONEY]NOT NULL DEFAULT 0, " _
                    & " [YARIS][MONEY]NOT NULL DEFAULT 0," _
                    & " [OTHERS][MONEY]NOT NULL DEFAULT 0,"
        temp = temp + ") ON [PRIMARY]"
        Call objO.UpdateData(temp)
        
        '6 Agustus 2012
        nBarByModel.Min = 0
        nBarByModel.Max = 16
        nBarByModel.value = 0
        nLabByModel.Caption = "Get Data Report : 0 Of 16"
        nProByModel.Caption = "Process :  1 Of 2"
        picBarByModel.Visible = True
        DoEvents
        
        'For i = Left(cboBulan1, InStr(cboBulan1, " |") - 1) To Left(cboBulan2, InStr(cboBulan2, " |") - 1)
            i = Left(cboBulan3, InStr(cboBulan3, " |") - 1)
            'Set rsTemp = objO.FetchData("EXEC REPOBULANAN3 " & Left(cboBulan1, InStr(cboBulan1, " |") - 1) & ", " & Text(0) & "")
            
            Call objO.DeleteData("Delete FROM BatchBulananByModel Where Bulan = '" & i & "' and Tahun = '" & Text(2).Text & "'")
            
           '6 Agustus 2012
            Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3A_ByModel " & i & "," & Text(2) & "")
            Call BarProsesByModel(1, 16, 1, 2, 1, "Get Data Unit Entry ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3B_ByModel " & i & "," & Text(2) & "")
            Call BarProsesByModel(2, 16, 1, 2, 2, "Get Data Unit Entry ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3C_ByModel " & i & "," & Text(2) & "")
            Call BarProsesByModel(3, 16, 1, 2, 3, "Get Data Unit Entry ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3D_ByModel " & i & "," & Text(2) & "")
            Call BarProsesByModel(4, 16, 1, 2, 4, "Get Data Revenue ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3E_ByModel " & i & "," & Text(2) & "")
            Call BarProsesByModel(5, 16, 1, 2, 5, "Get Data Revenue ...")
            DoEvents
            
            Set rsTemp = objO.FetchData("Select Keterangan, SUM(JUMLAH) AS JUMLAH, Bulan, TAHUN,KODE_KATEGORY from BatchBulananBymodel  Where Bulan = '" & i & "' and Tahun = '" & Text(2).Text & "' group by BULAN, Tahun, KETERANGAN,KODE_KATEGORY")
            Call BarProsesByModel(15, 16, 1, 2, 15, "Get Data Report ...")
            DoEvents
            
            If Not rsTemp.EOF Then
                nTotal = 0
                nTotal2 = 0
                nSubTotal = 0
                nI = 1
                namaBulan = GetNamaBulan(Left(cboBulan1, InStr(cboBulan1, " |") - 1))

                nBar.Min = 0
                nBar.Max = 35
                nBar.value = 0
                nLab.Caption = "Calculation Data : 0 Of 30"
                nPro.Caption = "Process :  2 Of 3"
                DoEvents
                Dim rsPencarian As ADODB.Recordset
                Do While nI <= 30
                    Select Case nI
                        Case 1
                            nUrut = nI
                            Set rsPencarian = rsTemp.Clone
                            nKeterangan = "10K saja "
                            rsPencarian.filter = "KETERANGAN = '10'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil 'Trim(Left(hasil, InStr(hasil, "|") - 1))
                                       ' nTotal = CDbl(Mid(hasil, InStr(hasil, "|") + 1, Len(hasil) - InStr(hasil, "|")))
                                        
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                'nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
'
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 10 K ...")
                        Case 2
                            nUrut = nI
                            Set rsPencarian = rsTemp.Clone
                            nKeterangan = "20K saja "
                            rsPencarian.filter = "KETERANGAN = '20'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 20 K ...")
                        Case 3
                            nUrut = nI
                            nKeterangan = "30K saja "
                            Set rsPencarian = rsTemp.Clone
                            nKeterangan = "10K saja "
                            rsPencarian.filter = "KETERANGAN = '30'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 30 K ...")
                        Case 4
                            nUrut = nI
                            nKeterangan = "40K saja "
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = '40'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 40 K ...")
                        Case 5
                            nUrut = nI
                            nKeterangan = "50K saja "
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = '50'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 50 K ...")
                        Case 6
                            nUrut = nI
                            nKeterangan = "60K saja "
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = '60'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 60 K ...")
                        Case 7
                            nUrut = nI
                            nKeterangan = "70K saja "
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = '70'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 70 K ...")
                        Case 8
                            nUrut = nI
                            nKeterangan = "80K saja "
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = '80'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 80 K ...")
                        Case 9
                            nUrut = nI
                            nKeterangan = "90K saja "
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = '90'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 90 K ...")
                        Case 10
                            nUrut = nI
                            nKeterangan = "100K saja "
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = '100'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data 100 K ...")
                        Case 11
                            nUrut = nI
                            nKeterangan = "Kelipatan 5K dibawah 100K "
                          '  rsTemp.Find "KETERANGAN = 'Kelipatan 5K'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'Kelipatan 5K'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data Kelipatan 5 K < 100 K ...")
                        Case 12
                            nUrut = nI
                            nKeterangan = "SBE lainnya diatas 100K "
                           ' rsTemp.Find "KETERANGAN = 'Lainnya Diatas 100K'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'Lainnya Diatas 100K'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotSBE21 = nTotSBE21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "SBE")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data > 100 K ...")
                        Case 13
                            ' TOTAL SBE
                            nUrut = nI
                            nKeterangan = "  TOTAL SERVIS BERKALA EKSTERNAL "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE21), 2, True)
                                Call objO.UpdateData(sSQL)
                     
                            Else
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(nTotSBE21), 2, True)
                                Call objO.UpdateData(sSQL)
                            End If
                            nSubTotal = nSubTotal + nTotal
                            TotalSBE = nSubTotal
                            TotalUnitSBE = TotalSBE
                            nTotal = 0
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data Total SBE ...")
                        Case 14
                            nUrut = nI
                            nKeterangan = "Unit Ganti Oli Mesin "
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'Unit Ganti Oli Mesin'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotGR21 = nTotGR21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "GRP")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data Ganti Oli Mesin ...")
                        Case 15
                            nUrut = nI
                            nKeterangan = "General Repair Lainnya "
                            'rsTemp.Find "KETERANGAN = 'Lainnya General Repair1'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'Lainnya General Repair1'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nTotGR21 = nTotGR21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "GRP")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data Other GR ...")
                        
                        Case 17
                            ' SUB TOTAL GRP
                            nUrut = nI - 1
                            nKeterangan = "  SUBTOTAL GENERAL REPAIR "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(nTotGR2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(nTotGR4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nTotGR8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(nTotGR9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(nTotGR10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(nTotGR11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nTotGR13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(nTotGR14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(nTotGR15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(nTotGR17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(nTotGR18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR21), 2, True)
                                Call objO.UpdateData(sSQL)
                     
                            Else
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(nTotGR2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(nTotGR4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nTotGR8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(nTotGR9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(nTotGR10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(nTotGR11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(nTotGR12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nTotGR13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(nTotGR14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(nTotGR15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(nTotGR17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(nTotGR18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(nTotGR21), 2, True)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProsesByModel(nUrut, 120, 2, 2, nUrut, "Proses Data Sub Total GR ...")
                        Case 18
                            ' TOTAL CPUS
                            nUrut = nI - 1
                            totCPUS1 = nTotSBE1 + nTotGR1
                            totCPUS2 = nTotSBE2 + nTotGR2
                            totCPUS3 = nTotSBE3 + nTotGR3
                            totCPUS4 = nTotSBE4 + nTotGR4
                            totCPUS5 = nTotSBE5 + nTotGR5
                            totCPUS6 = nTotSBE6 + nTotGR6
                            totCPUS7 = nTotSBE7 + nTotGR7
                            totCPUS8 = nTotSBE8 + nTotGR8
                            totCPUS9 = nTotSBE9 + nTotGR9
                            totCPUS10 = nTotSBE10 + nTotGR10
                            totCPUS11 = nTotSBE11 + nTotGR11
                            totCPUS12 = nTotSBE12 + nTotGR12
                            totCPUS13 = nTotSBE13 + nTotGR13
                            totCPUS14 = nTotSBE14 + nTotGR14
                            totCPUS15 = nTotSBE15 + nTotGR15
                            totCPUS16 = nTotSBE16 + nTotGR16
                            totCPUS17 = nTotSBE17 + nTotGR17
                            totCPUS18 = nTotSBE18 + nTotGR18
                            totCPUS19 = nTotSBE19 + nTotGR19
                            totCPUS20 = nTotSBE20 + nTotGR20
                            totCPUS21 = nTotSBE21 + nTotGR21
                                                    
                            
                            nKeterangan = "  TOTAL CPUS "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(totCPUS2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(totCPUS4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(totCPUS8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(totCPUS9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(totCPUS10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(totCPUS11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(totCPUS13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(totCPUS14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(totCPUS15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(totCPUS17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(totCPUS18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS21), 2, True)
                                Call objO.UpdateData(sSQL)
                     
                            Else
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(totCPUS2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(totCPUS4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(totCPUS8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(totCPUS9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(totCPUS10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(totCPUS11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(totCPUS12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(totCPUS13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(totCPUS14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(totCPUS15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(totCPUS17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(totCPUS18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(totCPUS21), 2, True)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Total CPUS ...")
                        Case 19
                            nUrut = nI - 1
                            nKeterangan = "1K Penjualan Outlet Sama "
                            'rsTemp.Find "KETERANGAN = 'SBI 1 dealer'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'SBI 1 dealer'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        'nTotGR21 = nTotGR21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "GRP")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 1K Local ...")
                        Case 20
                            nUrut = nI - 1
                            nKeterangan = "1K Total "
                            'rsTemp.Find "KETERANGAN = 'SBI beda dealer'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'SBI beda dealer'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nonCPUS21 = nonCPUS21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "NONCPUS")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data 1K Total ...")
                        Case 21
                            nUrut = nI - 1
                            nKeterangan = "Warranty "
                            'rsTemp.Find "KETERANGAN = 'nilai WARRANTY'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'nilai WARRANTY'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nonCPUS21 = nonCPUS21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "NONCPUS")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Warranty ...")
                        Case 22
                            nUrut = nI - 1
                            nKeterangan = "Return Job "
                            'rsTemp.Find "KETERANGAN = 'Return Job'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'Return Job'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nonCPUS21 = nonCPUS21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "NONCPUS")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data RTJ ...")
                        Case 23
                            nUrut = nI - 1
                            nKeterangan = "Pre Delivery Servis "
                           ' rsTemp.Find "KETERANGAN = 'PDS'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'PDS'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        nonCPUS21 = nonCPUS21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "NONCPUS")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data PDS ...")
                        Case 24
                            ' SUBTOTAL NONCPUS
                            nUrut = nI - 1
                            nKeterangan = "  SUBTOTAL NON CPUS "
                             If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS21), 2, True)
                                Call objO.UpdateData(sSQL)
                     
                            Else
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(nonCPUS21), 2, True)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Sub Total Non CPUS ...")
                        Case 25
                            nUrut = nI - 1
                            nKeterangan = "  JUMLAH TOTAL UNIT "
                            totUE1 = totCPUS1 + nonCPUS1
                            totUE2 = totCPUS2 + nonCPUS2
                            totUE3 = totCPUS3 + nonCPUS3
                            totUE4 = totCPUS4 + nonCPUS4
                            totUE5 = totCPUS5 + nonCPUS5
                            totUE6 = totCPUS6 + nonCPUS6
                            totUE7 = totCPUS7 + nonCPUS7
                            totUE8 = totCPUS8 + nonCPUS8
                            totUE9 = totCPUS1 + nonCPUS9
                            totUE10 = totCPUS10 + nonCPUS10
                            totUE11 = totCPUS11 + nonCPUS11
                            totUE12 = totCPUS12 + nonCPUS12
                            totUE13 = totCPUS13 + nonCPUS13
                            totUE14 = totCPUS14 + nonCPUS14
                            totUE15 = totCPUS15 + nonCPUS15
                            totUE16 = totCPUS16 + nonCPUS16
                            totUE17 = totCPUS17 + nonCPUS17
                            totUE18 = totCPUS18 + nonCPUS18
                            totUE19 = totCPUS19 + nonCPUS19
                            totUE20 = totCPUS20 + nonCPUS20
                            totUE21 = totCPUS21 + nonCPUS21
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(totUE1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(totUE2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(totUE3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(totUE4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(totUE5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(totUE6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(totUE7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(totUE8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(totUE9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(totUE10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(totUE11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(totUE12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(totUE13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(totUE14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(totUE15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(totUE16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(totUE17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(totUE18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(totUE19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(totUE20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(totUE21), 2, True)
                                Call objO.UpdateData(sSQL)
                     
                            Else
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(totUE1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(totUE2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(totUE3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(totUE4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(totUE5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(totUE6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(totUE7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(totUE8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(totUE9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(totUE10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(totUE11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(totUE12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(totUE13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(totUE14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(totUE15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(totUE16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(totUE17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(totUE18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(totUE19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(totUE20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(totUE21), 2, True)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data THS ...")
                        Case 26
                            ' JUMLAH TOTAL UNIT
                            nUrut = nI - 1
                            nKeterangan = "Labor "
                            'rsTemp.Find "KETERANGAN = 'Revenue Jasa Total'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'Revenue Jasa Total'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        labor21 = labor21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "LABOR")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Revenue Jasa ...")
                        
                        Case 27
                            nUrut = nI - 1
                            nKeterangan = "Part Toyota & Oli TGMO (CPUS + Non CPUS)"
                            'rsTemp.Find "KETERANGAN = 'Revenue Part dan Oli TGMO Total'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'Revenue Part dan Oli TGMO Total'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        labor21 = labor21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "LABOR")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part & Oli Genuine ...")
                       
                        Case 28
                            nUrut = nI - 1
                            nKeterangan = "Oli / Bahan Non Toyota (CPUS + Non CPUS)"
                            'rsTemp.Find "KETERANGAN = 'Revenue Oli dan Bahan Non Toyota Total'", , adSearchForward, 1
                             Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'Revenue Oli dan Bahan Non Toyota Total'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        labor21 = labor21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "LABOR")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Part & Oli Non Genuine ...")
                        Case 29
                            nUrut = nI - 1
                            nKeterangan = "Pekerjaan Luar (CPUS + Non CPUS)"
                            'rsTemp.Find "KETERANGAN = 'Revenue Pekerjaan Luar Total'", , adSearchForward, 1
                            Set rsPencarian = rsTemp.Clone
                            rsPencarian.filter = "KETERANGAN = 'Revenue Pekerjaan Luar Total'"
                            If Not rsPencarian.EOF Then
                                rsPencarian.MoveFirst
                                Do While Not rsPencarian.EOF
                                    If IsNull(rsPencarian("KODE_KATEGORY")) Then
                                        If cekInsert(nUrut, nKeterangan) Then
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 1)
                                        Else
                                            sSQL = insertTabelByModel("Others", namaBulan, nUrut, nKeterangan, CDbl(rsPencarian("JUMLAH")), 2)
                                        End If
                                        labor21 = labor21 + CDbl(rsPencarian("JUMLAH"))
                                    Else
                                        hasil = cariSQLByModel(Trim(rsPencarian("KODE_KATEGORY")), CDbl(rsPencarian("JUMLAH")), namaBulan, nUrut, nKeterangan, nTotal, "LABOR")
                                        sSQL = hasil
                                    End If
                                    
                                    rsPencarian.MoveNext
                                    Call objO.UpdateData(sSQL)
                                Loop
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 1)
                                Else
                                    sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(0), 2)
                                End If
                                nTotal = nTotal + CDbl(0)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Sublet ...")
                        Case 30
                            ' SUB TOTAL REVENUE CPUS
                            nUrut = nI - 1
                            nKeterangan = "  SUBTOTAL REVENUE CPUS "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(labor1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(labor2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(labor3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(labor4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(labor5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(labor6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(labor7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(labor8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(labor9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(labor10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(labor11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(labor12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(labor13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(labor14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(labor15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(labor16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(labor17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(labor18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(labor19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(labor20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(labor21), 2, True)
                                Call objO.UpdateData(sSQL)
                     
                            Else
                                sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(labor1), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(labor2), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(labor3), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(labor4), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(labor5), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(labor6), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(labor7), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(labor8), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(labor9), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(labor10), 2, True)
                                Call objO.UpdateData(sSQL)
                   
                                sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(labor11), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(labor12), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(labor13), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(labor14), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(labor15), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(labor16), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(labor17), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(labor18), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(labor19), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(labor20), 2, True)
                                Call objO.UpdateData(sSQL)
                                
                                sSQL = insertTabelByModel("OTHERS", namaBulan, nUrut, nKeterangan, CDbl(labor21), 2, True)
                                Call objO.UpdateData(sSQL)
                            End If
                            Call BarProses(nUrut, 120, 2, 2, nUrut, "Proses Data Revenue CPUS ...")
                        
                    End Select
                    nI = nI + 1
                    
                Loop
            End If
            rsTemp.Close
        'Next
        nBar.Min = 0
        nBar.Max = 10
        nBar.value = 0
        nLab.Caption = "Get Data Report : 0 Of 10"
        nPro.Caption = "Process :  1 Of 2"
        nKet.Caption = "..."
        picBarByModel.Visible = False
        DoEvents
        Shape1.Height = 615
        Call CetakByModel
       ' Call Transfer_JAF(CInt(Left(cboBulan1, InStr(cboBulan1, " |") - 1)), Text(0).Text)
    Else
        MsgBox pesan, vbCritical, "KELENGKAPAN"
        Exit Sub
    End If
End Sub

Private Function cariSQLByModel(kendaraan As String, JUMLAH As Double, namaBulan As String, nUrut As Integer, nKeterangan As String, nTotal As Double, tipe As String) As String
    Select Case kendaraan
    Case "AGYA"
        If tipe = "SBE" Then
            nTotSBE1 = nTotSBE1 + CDbl(JUMLAH)
         '   JUMLAH = nTotSBE1
        ElseIf tipe = "GRP" Then
            nTotGR1 = nTotGR1 + CDbl(JUMLAH)
         '   JUMLAH = nTotGR1
        ElseIf tipe = "NONCPUS" Then
            nonCPUS1 = nonCPUS1 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor1 = labor1 + CDbl(JUMLAH)
        End If
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("AGYA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "ALPHARD"
        If tipe = "SBE" Then
            nTotSBE2 = nTotSBE2 + CDbl(JUMLAH)
         '   JUMLAH = nTotSBE2
        ElseIf tipe = "GRP" Then
            nTotGR2 = nTotGR2 + CDbl(JUMLAH)
         '   JUMLAH = nTotGR2
        ElseIf tipe = "NONCPUS" Then
            nonCPUS2 = nonCPUS2 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor2 = labor2 + CDbl(JUMLAH)
        End If
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("ALPHARD", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "AVANZA"
        If tipe = "SBE" Then
            nTotSBE3 = nTotSBE3 + CDbl(JUMLAH)
         '   JUMLAH = nTotSBE3
        ElseIf tipe = "GRP" Then
            nTotGR3 = nTotGR3 + CDbl(JUMLAH)
        ElseIf tipe = "NONCPUS" Then
            nonCPUS3 = nonCPUS3 + CDbl(JUMLAH)
         '   JUMLAH = nTotGR3
        ElseIf tipe = "LABOR" Then
            labor3 = labor3 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("AVANZA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "CAMRY"
        If tipe = "SBE" Then
            nTotSBE4 = nTotSBE4 + CDbl(JUMLAH)
         '   JUMLAH = nTotSBE4
        ElseIf tipe = "GRP" Then
            nTotGR4 = nTotGR4 + CDbl(JUMLAH)
        ElseIf tipe = "NONCPUS" Then
            nonCPUS4 = nonCPUS4 + CDbl(JUMLAH)
         '   JUMLAH = nTotGR4
        ElseIf tipe = "LABOR" Then
            labor4 = labor4 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("CAMRY", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "COROLLA"
        If tipe = "SBE" Then
            nTotSBE5 = nTotSBE5 + CDbl(JUMLAH)
        '    JUMLAH = nTotSBE5
        ElseIf tipe = "GRP" Then
            nTotGR5 = nTotGR5 + CDbl(JUMLAH)
        '    JUMLAH = nTotGR5
        ElseIf tipe = "NONCPUS" Then
            nonCPUS5 = nonCPUS5 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor5 = labor5 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("COROLLA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "DYNA"
        If tipe = "SBE" Then
            nTotSBE6 = nTotSBE6 + CDbl(JUMLAH)
        '    JUMLAH = nTotSBE6
        ElseIf tipe = "GRP" Then
            nTotGR6 = nTotGR6 + CDbl(JUMLAH)
        '    JUMLAH = nTotGR6
        ElseIf tipe = "NONCPUS" Then
            nonCPUS6 = nonCPUS6 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor6 = labor6 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("DYNA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "ETIOS"
        If tipe = "SBE" Then
            nTotSBE7 = nTotSBE7 + CDbl(JUMLAH)
        '    JUMLAH = nTotSBE7
        ElseIf tipe = "GRP" Then
            nTotGR7 = nTotGR7 + CDbl(JUMLAH)
        '    JUMLAH = nTotGR7
        ElseIf tipe = "NONCPUS" Then
            nonCPUS7 = nonCPUS7 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor7 = labor7 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "ETIOS")
            sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
         
       ' nTotal = nTotal + CDbl(JUMLAH)
    Case "ETIOS TAXI"
        If tipe = "SBE" Then
            nTotSBE7 = nTotSBE7 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE7
        ElseIf tipe = "GRP" Then
            nTotGR7 = nTotGR7 + CDbl(JUMLAH)
            'JUMLAH = nTotGR7
        ElseIf tipe = "NONCPUS" Then
            nonCPUS7 = nonCPUS7 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor7 = labor7 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "ETIOS")
            sSQL = insertTabelByModel("ETIOS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
       ' nTotal = nTotal + CDbl(JUMLAH)
    Case "FJCRUISER"
        If tipe = "SBE" Then
            nTotSBE8 = nTotSBE8 + CDbl(JUMLAH)
           ' JUMLAH = nTotSBE8
        ElseIf tipe = "GRP" Then
            nTotGR8 = nTotGR8 + CDbl(JUMLAH)
            'JUMLAH = nTotGR8
        ElseIf tipe = "NONCPUS" Then
            nonCPUS8 = nonCPUS8 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor8 = labor8 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("FJCRUISER", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
       
       ' nTotal = nTotal + CDbl(JUMLAH)
    Case "FORTUNER BENSIN"
        If tipe = "SBE" Then
            nTotSBE9 = nTotSBE9 + CDbl(JUMLAH)
           ' JUMLAH = nTotSBE9
        ElseIf tipe = "GRP" Then
            nTotGR9 = nTotGR9 + CDbl(JUMLAH)
           ' JUMLAH = nTotGR9
        ElseIf tipe = "NONCPUS" Then
            nonCPUS9 = nonCPUS9 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor9 = labor9 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "FORTUNER")
            sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "FORTUNER DIESEL"
        If tipe = "SBE" Then
            nTotSBE9 = nTotSBE9 + CDbl(JUMLAH)
        '    JUMLAH = nTotSBE9
        ElseIf tipe = "GRP" Then
            nTotGR9 = nTotGR9 + CDbl(JUMLAH)
        '    JUMLAH = nTotGR9
        ElseIf tipe = "NONCPUS" Then
            nonCPUS9 = nonCPUS9 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor9 = labor9 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "FORTUNER")
            sSQL = insertTabelByModel("FORTUNER", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "FT 86"
        If tipe = "SBE" Then
            nTotSBE10 = nTotSBE10 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE10
        ElseIf tipe = "GRP" Then
            nTotGR10 = nTotGR10 + CDbl(JUMLAH)
            'JUMLAH = nTotGR10
        ElseIf tipe = "NONCPUS" Then
            nonCPUS10 = nonCPUS10 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor10 = labor10 + CDbl(JUMLAH)
        End If
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("FT86", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "HILUX"
        If tipe = "SBE" Then
            nTotSBE11 = nTotSBE11 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE11
        ElseIf tipe = "GRP" Then
            nTotGR11 = nTotGR11 + CDbl(JUMLAH)
            'JUMLAH = nTotGR11
        ElseIf tipe = "NONCPUS" Then
            nonCPUS11 = nonCPUS11 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor11 = labor11 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("HILUX", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "INNOVA BENSIN" ' Or "INNOVA DIESEL"
        If tipe = "SBE" Then
            nTotSBE12 = nTotSBE12 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE12
        ElseIf tipe = "GRP" Then
            nTotGR12 = nTotGR12 + CDbl(JUMLAH)
            'JUMLAH = nTotGR12
        ElseIf tipe = "NONCPUS" Then
            nonCPUS12 = nonCPUS12 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor12 = labor12 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "INNOVA")
            sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "INNOVA DIESEL"
        If tipe = "SBE" Then
            nTotSBE12 = nTotSBE12 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE12
        ElseIf tipe = "GRP" Then
            nTotGR12 = nTotGR12 + CDbl(JUMLAH)
            'JUMLAH = nTotGR12
        ElseIf tipe = "NONCPUS" Then
            nonCPUS12 = nonCPUS12 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor12 = labor12 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "INNOVA")
            sSQL = insertTabelByModel("INNOVA", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "LAND CRUISER"
        If tipe = "SBE" Then
            nTotSBE13 = nTotSBE13 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE13
        ElseIf tipe = "GRP" Then
            nTotGR13 = nTotGR13 + CDbl(JUMLAH)
            'JUMLAH = nTotGR13
        ElseIf tipe = "NONCPUS" Then
            nonCPUS13 = nonCPUS13 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor13 = labor13 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("LANDCRUISER", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
       
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "LIMO" 'Or "LIMO CKD"
        If tipe = "SBE" Then
            nTotSBE14 = nTotSBE14 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE14
        ElseIf tipe = "GRP" Then
            nTotGR14 = nTotGR14 + CDbl(JUMLAH)
            'JUMLAH = nTotGR14
        ElseIf tipe = "NONCPUS" Then
            nonCPUS14 = nonCPUS14 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor14 = labor14 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "LIMO")
            sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
       '; nTotal = nTotal + CDbl(JUMLAH)
    Case "LIMO CKD"
        If tipe = "SBE" Then
            nTotSBE14 = nTotSBE14 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE14
        ElseIf tipe = "GRP" Then
            nTotGR14 = nTotGR14 + CDbl(JUMLAH)
            'JUMLAH = nTotGR14
        ElseIf tipe = "NONCPUS" Then
            nonCPUS14 = nonCPUS14 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor14 = labor14 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "LIMO")
            sSQL = insertTabelByModel("LIMO", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
       
        'nTotal = nTotal + CDbl(JUMLAH)
   Case "NAV 1"
        If tipe = "SBE" Then
            nTotSBE15 = nTotSBE15 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE15
        ElseIf tipe = "GRP" Then
            nTotGR15 = nTotGR15 + CDbl(JUMLAH)
            'JUMLAH = nTotGR15
        ElseIf tipe = "NONCPUS" Then
            nonCPUS15 = nonCPUS15 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor15 = labor15 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("NAV1", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
       
        'nTotal = nTotal + CDbl(JUMLAH)
   Case "PRIUS"
        If tipe = "SBE" Then
            nTotSBE16 = nTotSBE16 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE16
        ElseIf tipe = "GRP" Then
            nTotGR16 = nTotGR16 + CDbl(JUMLAH)
            'JUMLAH = nTotGR16
        ElseIf tipe = "NONCPUS" Then
            nonCPUS16 = nonCPUS16 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor16 = labor16 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("PRIUS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
   Case "RUSH"
        If tipe = "SBE" Then
            nTotSBE17 = nTotSBE17 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE17
        ElseIf tipe = "GRP" Then
            nTotGR17 = nTotGR17 + CDbl(JUMLAH)
            'JUMLAH = nTotGR17
        ElseIf tipe = "NONCPUS" Then
            nonCPUS17 = nonCPUS17 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor17 = labor17 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("RUSH", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
       ' nTotal = nTotal + CDbl(JUMLAH)
   Case "TOYOTA HI ACE"
        If tipe = "SBE" Then
            nTotSBE18 = nTotSBE18 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE18
        ElseIf tipe = "GRP" Then
            nTotGR18 = nTotGR18 + CDbl(JUMLAH)
            'JUMLAH = nTotGR18
        ElseIf tipe = "NONCPUS" Then
            nonCPUS18 = nonCPUS18 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor18 = labor18 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            sSQL = insertTabelByModel("HIACE", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
        
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "VIOS" 'Or "VIOS CKD"
        If tipe = "SBE" Then
            nTotSBE19 = nTotSBE19 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE19
        ElseIf tipe = "GRP" Then
            nTotGR19 = nTotGR19 + CDbl(JUMLAH)
            'JUMLAH = nTotGR19
        ElseIf tipe = "NONCPUS" Then
            nonCPUS19 = nonCPUS19 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor19 = labor19 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "VIOS")
            sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
      
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "VIOS CKD"
        If tipe = "SBE" Then
            nTotSBE19 = nTotSBE19 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE19
        ElseIf tipe = "GRP" Then
            nTotGR19 = nTotGR19 + CDbl(JUMLAH)
            'JUMLAH = nTotGR19
        ElseIf tipe = "NONCPUS" Then
            nonCPUS19 = nonCPUS19 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor19 = labor19 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "VIOS")
            sSQL = insertTabelByModel("VIOS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
       
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "YARIS" 'Or "YARIS CKD"
        If tipe = "SBE" Then
            nTotSBE20 = nTotSBE20 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE20
        ElseIf tipe = "GRP" Then
            nTotGR20 = nTotGR20 + CDbl(JUMLAH)
            'JUMLAH = nTotGR20
        ElseIf tipe = "NONCPUS" Then
            nonCPUS20 = nonCPUS20 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor20 = labor20 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "YARIS")
            sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
         
        'nTotal = nTotal + CDbl(JUMLAH)
    Case "YARIS CKD"
        If tipe = "SBE" Then
            nTotSBE20 = nTotSBE20 + CDbl(JUMLAH)
            'JUMLAH = nTotSBE20
        ElseIf tipe = "GRP" Then
            nTotGR20 = nTotGR20 + CDbl(JUMLAH)
            'JUMLAH = nTotGR20
        ElseIf tipe = "NONCPUS" Then
            nonCPUS20 = nonCPUS20 + CDbl(JUMLAH)
        ElseIf tipe = "LABOR" Then
            labor20 = labor20 + CDbl(JUMLAH)
        End If
        
        If cekInsert(nUrut, nKeterangan) Then
            sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 1)
        Else
            JUMLAH = JUMLAH + cekNilaiKolom(nUrut, nKeterangan, "YARIS")
            sSQL = insertTabelByModel("YARIS", namaBulan, nUrut, nKeterangan, CDbl(JUMLAH), 2)
        End If
         
        'nTotal = nTotal + CDbl(JUMLAH)
    End Select
    cariSQLByModel = sSQL '& "|" & nTotal
End Function


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton2.Top = CRViewer1.Top + 345
    XPButton2.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = XPButton2.Left + XPButton2.Width + 30
    
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
        XPButton2.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton2.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton2.Visible = True
        
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

Public Sub LeadSAToSales(Tahun As Integer, Bulan As Integer)
    tglAwal = Format(Bulan, "00") & "/01/" & Tahun
    tglAkhir = "dateadd(day,-1,DATEADD(month,1,'" & tglAwal & "'))"
    sSQL = "select  count(distinct b.Kodepr) as jml,'GR 1. Lead SA GR to Sales'"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..prospectmst b with (nolock)"
    sSQL = sSQL & " Left Join"
    sSQL = sSQL & " ("
    sSQL = sSQL & "     select x.kodepr, x.KodeProspect, X.Status, x.tglfollow from [" & dbReplikasiHO & "]..ProspectFollow X with (nolock)"
    sSQL = sSQL & "     inner join ("
    sSQL = sSQL & "     select kodepr, KodeProspect, MAX(tglfollow) as maxFollow"
    sSQL = sSQL & "     from [" & dbReplikasiHO & "]..ProspectFollow  with (nolock)"
    sSQL = sSQL & "     where tglfollow between '" & tglAwal & " 00:00:00' and " & tglAkhir & "+' 23:59:59' and kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & "     group by kodepr, KodeProspect) Y on x.KodePr = Y.KodePr and x.TglFollow = Y.maxFollow and x.KodeProspect = Y.KodeProspect"
    sSQL = sSQL & " ) F on b.KodePr = F.KodePr"
    sSQL = sSQL & " where b.TglEntry < dateadd(MONTH, 1, '" & tglAwal & "') and kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " and f.TglFollow  between '" & tglAwal & " 00:00:00' and " & tglAkhir & "+' 23:59:59' and b.sumber = '17'"
    
    ''sSQL = "select count(*) total from [" & dbReplikasiHO & "]..prospectkary where sa = 1 and year(TglEntry) = " & Tahun & " and MONTH(TglEntry) = " & Bulan
    totalLead2Sales = objO.FetchData_Double(sSQL)
End Sub

Public Sub DOSAToSales(Tahun As Integer, Bulan As Integer)
    tglAwal = Format(Bulan, "00") & "/01/" & Tahun
    tglAkhir = "dateadd(day,-1,DATEADD(month,1,'" & tglAwal & "'))"
    
    sSQL = "select COUNT(b.kodepr) as total "
    sSQL = sSQL & " From [SPK" & Right(gDb, 2) & "]..Summary a"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..FollowSPK b on a.SPK_No = b.NoSpk"
    sSQL = sSQL & " where TipeTr ='ND' and Status = 1"
    sSQL = sSQL & " and month(Tanggal) = " & Bulan & " and YEAR(tanggal) = '" & Tahun & "'"
    sSQL = sSQL & " and b.KodePr in (select KodePr from [" & dbReplikasiHO & "]..ProspectKary where SA = 1 and month(TglEntry) = " & Bulan & " and YEAR(TglEntry) = '" & Tahun & "' and KodeCabang = '" & Log.kodeCabang & "' )"

    sSQL = "select count(distinct b.Kodepr) as jml,'GR 2. DO Sa GR to SAles' "
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..prospectmst b  with (nolock)"
    sSQL = sSQL & " Left Join"
    sSQL = sSQL & " ("
    sSQL = sSQL & "     select x.kodepr, x.KodeProspect, X.KodePlan, X.Status, x.tglfollow from [" & dbReplikasiHO & "]..ProspectFollow X  with (nolock)"
    sSQL = sSQL & "     inner join ("
    sSQL = sSQL & "     select kodepr, KodeProspect, MAX(tglfollow) as maxFollow"
    sSQL = sSQL & "     from [" & dbReplikasiHO & "]..ProspectFollow  with (nolock)"
    sSQL = sSQL & "     where tglfollow between '" & tglAwal & " 00:00:00' and " & tglAkhir & "+' 23:59:59' and kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & "     group by kodepr, KodeProspect) Y on x.KodePr = Y.KodePr and x.TglFollow = Y.maxFollow and x.KodeProspect = Y.KodeProspect"
    sSQL = sSQL & " ) F on b.KodePr = F.KodePr"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on f.KodePr = c.KodePr and f.KodeProspect = c.KodeProspect and f.KodePlan = c.KodePlan"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..FollowSPK d on c.KodePr = d.KodePr and c.KodeProspect = d.KodeProspect and c.KodePlan = d.KodePlan"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..Summary e on d.NoSpk = e.No_bukti"
    sSQL = sSQL & " where e.tipetr = 'SPK' and e.status = '1' and b.sumber = '17' and b.kodecabang = '" & Log.kodeCabang & "' --and b.SalesPr = '0105-1132'"
    sSQL = sSQL & " and b.TglEntry < dateadd(MONTH, 1, " & tglAwal & "'')"
    sSQL = sSQL & " and F.TglFollow between '" & tglAwal & " 00:00:00' and " & tglAkhir & "+' 23:59:59'"
    sSQL = sSQL & " and e.Tanggal between '" & tglAwal & " 00:00:00' and " & tglAkhir & "+' 23:59:59'"
    totalDOSAtoSales = objO.FetchData_Double(sSQL)
End Sub

Public Sub BookingReqBySales(Tahun As Integer, Bulan As Integer)
    tglAwal = Format(Bulan, "00") & "/01/" & Tahun
    tglAkhir = "dateadd(day,-1,DATEADD(month,1,'" & tglAwal & "'))"
    
    sSQL = "select count(b.SalesPr) as BOC"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect and c.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " where a.kodecabang = '" & Log.kodeCabang & "' and a.TypePlan = 3 and"
    sSQL = sSQL & " (year(TglPlan) = '" & Tahun & "' and month(TglPlan) = '" & Bulan & "' )"
    sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and"
    sSQL = sSQL & " a.NoRangka in (select no_chassis from BOOK where FS_LOG_DEL = '' and year(DATE_BOOK) = '" & Tahun & "' and month(DATE_BOOK) = '" & Bulan & "') and"
    sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.kodeCabang & "' and KetPlan like '%reminder service%')"
    
    sSQL = "select COUNT(a.kodepr), 'GR 3. BOOKING REQUEST BY SALES'"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a with (nolock)"
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..ProspectFollow b with (nolock) on a.KodePr = b.KodePr and a.KodeProspect = b.KodeProspect and a.KodePlan = b.KodePlan and a.KodeCabang = b.KodeCabang"
    sSQL = sSQL & " left join SingleIDOri..UnitCustomer c with (nolock) on a.NoRangka = c.NoRangka"
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..BookingOnline d with (nolock) on c.NoPolisi = d.nopol and d.dealer = '" & gKdCbg & "'"
    sSQL = sSQL & " where a.TypePlan = 3 and b.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " and a.TglPlan between '" & tglAwal & "' and " & tglAkhir & " and a.KodePr <> ''"
    sSQL = sSQL & " and b.TglFollow between '" & tglAwal & " 00:00' and " & tglAkhir & "+' 23:59' and RetensiKet = 'Terhubung' and BoC = 'boc' and fproses = 1"
    sSQL = sSQL & " and noBookCarsWorld <> '' and LEN(replace(left(noBookCarsWorld,15),'sales:','')) = 9 and LEFT(nobookingapps,3)= 'BSV'"


    totalReqBookBySales = objO.FetchData_Double(sSQL)
End Sub

Public Sub BookedShowBySales(Tahun As Integer, Bulan As Integer)
    tglAwal = Format(Bulan, "00") & "/01/" & Tahun
    tglAkhir = "dateadd(day,-1,DATEADD(month,1,'" & tglAwal & "'))"
    
    sSQL = "select count(b.SalesPr) as WO"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect and c.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " where a.kodecabang = '" & Log.kodeCabang & "' and a.TypePlan = 3 and"
    sSQL = sSQL & " (year(TglPlan) = '" & Tahun & "' and month(tglplan) = '" & Bulan & "' )"
    sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and"
    sSQL = sSQL & " a.NoRangka in (select no_chassis from BOOK where FS_LOG_DEL = '' and year(DATE_BOOK) = '" & Tahun & "' and month(DATE_BOOK) = '" & Bulan & "') and"
    sSQL = sSQL & " a.NoRangka in (select no_chassis from JOBDTA where FS_LOG_DEL = '' and year(DATE) = '" & Tahun & "' and month(DATE) = '" & Bulan & "' and NOFAK <> '') and"
    sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.kodeCabang & "' and KetPlan like '%reminder service%')"
    
    
    sSQL = "select COUNT(a.kodepr),'GR 4. BOOKED SHOW BY SALES'"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a with (nolock)"
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..ProspectFollow b with (nolock) on a.KodePr = b.KodePr and a.KodeProspect = b.KodeProspect and a.KodePlan = b.KodePlan and a.KodeCabang = b.KodeCabang"
    sSQL = sSQL & " left join SingleIDOri..UnitCustomer c with (nolock) on a.NoRangka = c.NoRangka"
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..BookingOnline d with (nolock) on c.NoPolisi = d.nopol"
    sSQL = sSQL & " left join [" & gDb & "]..jobdta e with (nolock) on e.no_polisi = d.nopol and nofak <> '' and deliv <> 'BTL'"
    sSQL = sSQL & " where a.TypePlan = 3 and b.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " and a.TglPlan between '" & tglAwal & "' and " & tglAkhir & " and a.KodePr <> ''"
    sSQL = sSQL & " and b.TglFollow between '" & tglAwal & " 00:00' and " & tglAkhir & "+' 23:59' and RetensiKet = 'Terhubung' and BoC = 'boc' and fproses = 1"
    sSQL = sSQL & " and noBookCarsWorld <> '' and LEN(replace(left(noBookCarsWorld,15),'sales:','')) = 9 and LEFT(nobookingapps,3)= 'BSV'"
    sSQL = sSQL & " and e.date >= '09/01/2020'"

    totalBookBySales = objO.FetchData_Double(sSQL)
End Sub

Public Sub ContactRatioBySales(Tahun As Integer, Bulan As Integer)
    tglAwal = Format(Bulan, "00") & "/01/" & Tahun
    tglAkhir = "dateadd(day,-1,DATEADD(month,1,'" & tglAwal & "'))"
    
    sSQL = "select count(b.SalesPr) as reminder"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " where a.kodecabang = '" & Log.kodeCabang & "' and TypePlan = 3 and"
    sSQL = sSQL & " (year(TglPlan) = '" & Tahun & "' and month(tglplan) = '" & Bulan & "' ) and"
    sSQL = sSQL & " idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.kodeCabang & "' and KetPlan like '%reminder service%')"
    
    sSQL = "select COUNT(a.KodePr) Reminder"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a with (nolock)"
    sSQL = sSQL & " where a.TypePlan = 3 and a.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " and a.TglPlan between '" & tglAwal & "' and " & tglAkhir & " and a.KodePr <> ''"
    totalReminderSales = objO.FetchData_Double(sSQL)
    
    sSQL = "select count(b.SalesPr) as terfollow"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect and c.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " where a.kodecabang = '" & Log.kodeCabang & "' and a.TypePlan = 3 and"
    sSQL = sSQL & " (year(TglPlan) = '" & Tahun & "' and month(tglplan) = '" & Bulan & "' ) and"
    sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.kodeCabang & "' and KetPlan like '%reminder service%')"
    
    sSQL = "select COUNT(a.kodepr) FollowUp"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a  with (nolock)"
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..ProspectFollow b with (nolock) on a.KodePr = b.KodePr and a.KodeProspect = b.KodeProspect and a.KodePlan = b.KodePlan and a.KodeCabang = b.KodeCabang"
    sSQL = sSQL & " where a.TypePlan = 3 and b.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " and a.TglPlan between '" & tglAwal & "' and " & tglAkhir & " and a.KodePr <> ''"
    sSQL = sSQL & " and b.TglFollow between '" & tglAwal & " 00:00' and " & tglAkhir & "+' 23:59'"
    totalReminderTerhubungSales = objO.FetchData_Double(sSQL)
    
End Sub

Public Sub ComeAfterInvitedBySales(Tahun As Integer, Bulan As Integer)
    tglAwal = Format(Bulan, "00") & "/01/" & Tahun
    tglAkhir = "dateadd(day,-1,DATEADD(month,1,'" & tglAwal & "'))"
    
    sSQL = "select count(b.SalesPr) as WO"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect and c.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " where a.kodecabang = '" & Log.kodeCabang & "' and a.TypePlan = 3 and"
    sSQL = sSQL & " (year(TglPlan) = '" & Tahun & "' and month(tglplan) = '" & Bulan & "' )"
    sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and"
    sSQL = sSQL & " a.NoRangka in (select no_chassis from BOOK where FS_LOG_DEL = '' and year(DATE_BOOK) = '" & Tahun & "' and month(DATE_BOOK) = '" & Bulan & "') and"
    sSQL = sSQL & " a.NoRangka in (select no_chassis from JOBDTA where FS_LOG_DEL = '' AND noinv <> '' and year(invdate) = '" & Tahun & "' and month(invdate) = '" & Bulan & "' and NOFAK <> '') and"
    sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.kodeCabang & "' and KetPlan like '%reminder service%')"
    
    sSQL = "select COUNT(a.kodepr)"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a with (nolock)"
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..ProspectFollow b with (nolock) on a.KodePr = b.KodePr and a.KodeProspect = b.KodeProspect and a.KodePlan = b.KodePlan and a.KodeCabang = b.KodeCabang"
    sSQL = sSQL & " left join SingleIDOri..UnitCustomer c with (nolock) on a.NoRangka = c.NoRangka"
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..BookingOnline d with (nolock) on c.NoPolisi = d.nopol"
    sSQL = sSQL & " left join jobdta e with (nolock) on e.no_polisi = d.nopol and nofak <> '' and deliv <> 'BTL'"
    sSQL = sSQL & " where a.TypePlan = 3 and b.kodecabang = '" & Log.kodeCabang & "'"
    sSQL = sSQL & " and a.TglPlan between '" & tglAwal & "' and " & tglAkhir & " and a.KodePr <> ''"
    sSQL = sSQL & " and b.TglFollow between '" & tglAwal & " 00:00' and " & tglAkhir & "+' 23:59' and RetensiKet = 'Terhubung' and BoC = 'boc' and fproses = 1"
    sSQL = sSQL & " and noBookCarsWorld <> '' and LEN(replace(left(noBookCarsWorld,15),'sales:','')) = 9 and LEFT(nobookingapps,3)= 'BSV'"
    sSQL = sSQL & " and e.date >= '09/01/2020'"
    totalInvoiceSales = objO.FetchData_Double(sSQL)
End Sub

Public Sub WOContributionByBP(Tahun As Integer, Bulan As Integer)
    tglAwal = Format(Bulan, "00") & "/01/" & Tahun
    tglAkhir = "dateadd(day,-1,DATEADD(month,1,'" & tglAwal & "'))"
    
    sSQL = "select COUNT(c.NO_WO) as total"
    sSQL = sSQL & " from BookingOnline a"
    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..BookingOnline b on a.id = b.id"
    sSQL = sSQL & " inner join JOBDTA c on a.NoFak = c.NOFAK and c.FS_LOG_DEL = ''"
    sSQL = sSQL & " where a.NoFak <> '' and B.nobookingapps like 'TGR%' and MONTH(c.date) = '" & i & "' and YEAR(c.date) = '" & Text(0).Text & "'"
    
    
    sSQL = "select COUNT(nobookingapps)"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..BookingOnline a with (nolock)"
    sSQL = sSQL & " left join jobdta b with (nolock) on a.no_booking = b.nofak"
    sSQL = sSQL & " where noBookCarsWorld <> '' and LEN(replace(left(noBookCarsWorld,14),'SABP:','')) = 9 and LEFT(nobookingapps,3)= 'TGR'"
    sSQL = sSQL & " and fProses not in (0,3) and no_booking <> ''"
    sSQL = sSQL & " and tgl_entry between '" & tglAwal & "' and " & tglAkhir & "+' 23:59' and dealer = '" & gKdCbg & "'"
    totalWOBPtoGRP = objO.FetchData_Double(sSQL)
End Sub
