VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmHOPeriode 
   Caption         =   "Table Periode Label Sent"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   3615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin TDBNumber6Ctl.TDBNumber nilaiMin1 
         Height          =   300
         Left            =   2760
         TabIndex        =   2
         Top             =   840
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":0000
         Caption         =   "frmHOPeriode.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":008C
         Keys            =   "frmHOPeriode.frx":00AA
         Spin            =   "frmHOPeriode.frx":00F4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber nilaiMax1 
         Height          =   300
         Left            =   3960
         TabIndex        =   3
         Top             =   840
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":011C
         Caption         =   "frmHOPeriode.frx":013C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":01A8
         Keys            =   "frmHOPeriode.frx":01C6
         Spin            =   "frmHOPeriode.frx":0210
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber nilaiMin2 
         Height          =   300
         Left            =   2760
         TabIndex        =   5
         Top             =   1200
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":0238
         Caption         =   "frmHOPeriode.frx":0258
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":02C4
         Keys            =   "frmHOPeriode.frx":02E2
         Spin            =   "frmHOPeriode.frx":032C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber nilaiMax2 
         Height          =   300
         Left            =   3960
         TabIndex        =   6
         Top             =   1200
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":0354
         Caption         =   "frmHOPeriode.frx":0374
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":03E0
         Keys            =   "frmHOPeriode.frx":03FE
         Spin            =   "frmHOPeriode.frx":0448
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber nilaiMin3 
         Height          =   300
         Left            =   2760
         TabIndex        =   8
         Top             =   1560
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":0470
         Caption         =   "frmHOPeriode.frx":0490
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":04FC
         Keys            =   "frmHOPeriode.frx":051A
         Spin            =   "frmHOPeriode.frx":0564
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber nilaiMax3 
         Height          =   300
         Left            =   3960
         TabIndex        =   9
         Top             =   1560
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":058C
         Caption         =   "frmHOPeriode.frx":05AC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":0618
         Keys            =   "frmHOPeriode.frx":0636
         Spin            =   "frmHOPeriode.frx":0680
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber nilaiMin4 
         Height          =   300
         Left            =   2760
         TabIndex        =   11
         Top             =   1920
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":06A8
         Caption         =   "frmHOPeriode.frx":06C8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":0734
         Keys            =   "frmHOPeriode.frx":0752
         Spin            =   "frmHOPeriode.frx":079C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber nilaiMax4 
         Height          =   300
         Left            =   3960
         TabIndex        =   12
         Top             =   1920
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":07C4
         Caption         =   "frmHOPeriode.frx":07E4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":0850
         Keys            =   "frmHOPeriode.frx":086E
         Spin            =   "frmHOPeriode.frx":08B8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber nilaiMin5 
         Height          =   300
         Left            =   2760
         TabIndex        =   14
         Top             =   2280
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":08E0
         Caption         =   "frmHOPeriode.frx":0900
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":096C
         Keys            =   "frmHOPeriode.frx":098A
         Spin            =   "frmHOPeriode.frx":09D4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber nilaiMax5 
         Height          =   300
         Left            =   3960
         TabIndex        =   15
         Top             =   2280
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   529
         Calculator      =   "frmHOPeriode.frx":09FC
         Caption         =   "frmHOPeriode.frx":0A1C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOPeriode.frx":0A88
         Keys            =   "frmHOPeriode.frx":0AA6
         Spin            =   "frmHOPeriode.frx":0AF0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label8 
         BackColor       =   &H0068ECE6&
         Caption         =   "NILAI MAX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H0068ECE6&
         Caption         =   "NILAI MIN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H0068ECE6&
         Caption         =   ".:: SETTING PERIODE LS ::."
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
         Left            =   1560
         TabIndex        =   16
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0068ECE6&
         Caption         =   "Minggu 5 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0068ECE6&
         Caption         =   "Minggu 4 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0068ECE6&
         Caption         =   "Minggu 3 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0068ECE6&
         Caption         =   "Minggu 2 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0068ECE6&
         Caption         =   "Minggu 1 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   4635
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   8176
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "   Tabel Periode Label Sent"
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
      Top             =   0
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   3840
      TabIndex        =   19
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdKeluar 
      Height          =   360
      Left            =   4800
      TabIndex        =   20
      Top             =   4080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Kembali Ke &Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmHOPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsPeriode As New ADODB.Recordset

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
          Dim boleh       As Boolean
          Dim pembilang   As Integer
          Dim penyebut    As Integer
          Dim pesan       As String
          Dim isBegin     As Boolean
          Dim persenFoll  As Integer
          
1        On Error GoTo cmdSimpan_Click_Error
          
2         isBegin = False
3         boleh = True
          
4         pembilang = 0
5         penyebut = 0
          
6         If nilaiMin1 > 0 Then
7             pembilang = pembilang + nilaiMin1
8             penyebut = penyebut + 1
9         End If
          
10        If nilaiMin2 > 0 Then
11            pembilang = pembilang + nilaiMin2
12            penyebut = penyebut + 1
13        End If
          
14        If nilaiMin3 > 0 Then
15            pembilang = pembilang + nilaiMin3
16            penyebut = penyebut + 1
17        End If
          
18        If nilaiMin4 > 0 Then
19            pembilang = pembilang + nilaiMin4
20            penyebut = penyebut + 1
21        End If
          
22        If nilaiMin5 > 0 Then
23            pembilang = pembilang + nilaiMin5
24            penyebut = penyebut + 1
25        End If
          
26        If penyebut > 0 Then
27            Set rsPeriode = objO.FetchDataHO("SELECT PERSENFOLLOW FROM T_PROFIL WITH (NOLOCK)")
28            If Not rsPeriode.EOF Then
29                persenFoll = CInt(rsPeriode!persenfollow)
30            End If
31            If (pembilang / penyebut) < persenFoll Then
32                MsgBox "Nilai total 1 bulan masih belum mencukupi syarat minimal", vbInformation
33                Exit Sub
34            End If
35        End If
          
36        If nilaiMin1 > nilaiMax1 Then
37            boleh = False
38            pesan = pesan & " Nilai min minggu 1 lebih besar dari nilai max,"
39        ElseIf nilaiMin2 > nilaiMax2 Then
40            boleh = False
41            pesan = pesan & " Nilai min minggu 2 lebih besar dari nilai max,"
42        ElseIf nilaiMin3 > nilaiMax3 Then
43            boleh = False
44            pesan = pesan & " Nilai min minggu 3 lebih besar dari nilai max,"
45        ElseIf nilaiMin4 > nilaiMax4 Then
46            boleh = False
47            pesan = pesan & " Nilai min minggu 4 lebih besar dari nilai max,"
48        ElseIf nilaiMin5 > nilaiMax5 Then
49            boleh = False
50            pesan = pesan & " Nilai min minggu 5 lebih besar dari nilai max,"
51        End If
          
52        If boleh = True Then
53            Call objHT.OpenTransactionHO
54            isBegin = True
55            Set rsPeriode = objHT.FetchDataHO("SELECT [WEEK],MINLS,MAXLS FROM MASTER_SETTING WITH (NOLOCK)")
56            If Not rsPeriode.EOF Then
57                Call objHT.InsertData_Ho("DELETE FROM MASTER_SETTING")
58            End If
              
59            Call objHT.InsertData_Ho("INSERT INTO MASTER_SETTING ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('1','" & nilaiMin1 & "','" & nilaiMax1 & "','" & gSA & "',GETDATE())")
60            Call objHT.InsertData_Ho("INSERT INTO master_settinglog ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('1','" & nilaiMin1 & "','" & nilaiMax1 & "','" & gSA & "',GETDATE())")
              
61            Call objHT.InsertData_Ho("INSERT INTO MASTER_SETTING ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('2','" & nilaiMin2 & "','" & nilaiMax2 & "','" & gSA & "',GETDATE())")
62            Call objHT.InsertData_Ho("INSERT INTO master_settinglog ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('2','" & nilaiMin2 & "','" & nilaiMax2 & "','" & gSA & "',GETDATE())")
              
63            Call objHT.InsertData_Ho("INSERT INTO MASTER_SETTING ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('3','" & nilaiMin3 & "','" & nilaiMax3 & "','" & gSA & "',GETDATE())")
64            Call objHT.InsertData_Ho("INSERT INTO master_settinglog ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('3','" & nilaiMin3 & "','" & nilaiMax3 & "','" & gSA & "',GETDATE())")

65            Call objHT.InsertData_Ho("INSERT INTO MASTER_SETTING ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('4','" & nilaiMin4 & "','" & nilaiMax4 & "','" & gSA & "',GETDATE())")
66            Call objHT.InsertData_Ho("INSERT INTO master_settinglog ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('4','" & nilaiMin4 & "','" & nilaiMax4 & "','" & gSA & "',GETDATE())")

67            Call objHT.InsertData_Ho("INSERT INTO MASTER_SETTING ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('5','" & nilaiMin5 & "','" & nilaiMax5 & "','" & gSA & "',GETDATE())")
68            Call objHT.InsertData_Ho("INSERT INTO master_settinglog ([week],[minLs],[maxLS],[userEntry],[tglEntry]) VALUES ('5','" & nilaiMin5 & "','" & nilaiMax5 & "','" & gSA & "',GETDATE())")

69            Call objHT.InsertData_Ho("INSERT INTO TEMPKIRIM VALUES ('PERIODE', 'ALL','S')")
              
70            Call objHT.CloseTransactionHO
71            MsgBox "Data berhasil disimpan", vbInformation
72            Call refreshData
73        Else
74            MsgBox "Terjadi kesalahan." & vbCrLf & pesan, vbInformation
75        End If

76       On Error GoTo 0
77       Exit Sub

cmdSimpan_Click_Error:
78        If isBegin Then
79            Call objHT.CloseTransactionHO_err
80        End If
81        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmHOPeriode"
          
End Sub

Private Sub Form_Load()
    Call refreshData
End Sub

Private Sub refreshData()
    Set rsPeriode = objO.FetchDataHO("SELECT [WEEK],MINLS,MAXLS FROM MASTER_SETTING WITH (NOLOCK)")
    nilaiMax1 = 0
    nilaiMax2 = 0
    nilaiMax3 = 0
    nilaiMax4 = 0
    nilaiMax5 = 0
    
    nilaiMin1 = 0
    nilaiMin2 = 0
    nilaiMin3 = 0
    nilaiMin4 = 0
    nilaiMin5 = 0
    
    If Not rsPeriode.EOF Then
        rsPeriode.MoveFirst
        Do While Not rsPeriode.EOF
            If NulSID(rsPeriode!Week) = 1 Then
                nilaiMin1 = NulSID(rsPeriode!MINLS)
                nilaiMax1 = NulSID(rsPeriode!MAXLS)
            ElseIf NulSID(rsPeriode!Week) = 2 Then
                nilaiMin2 = NulSID(rsPeriode!MINLS)
                nilaiMax2 = NulSID(rsPeriode!MAXLS)
            ElseIf NulSID(rsPeriode!Week) = 3 Then
                nilaiMin3 = NulSID(rsPeriode!MINLS)
                nilaiMax3 = NulSID(rsPeriode!MAXLS)
            ElseIf NulSID(rsPeriode!Week) = 4 Then
                nilaiMin4 = NulSID(rsPeriode!MINLS)
                nilaiMax4 = NulSID(rsPeriode!MAXLS)
            ElseIf NulSID(rsPeriode!Week) = 5 Then
                nilaiMin5 = NulSID(rsPeriode!MINLS)
                nilaiMax5 = NulSID(rsPeriode!MAXLS)
            End If
            rsPeriode.MoveNext
        Loop
        
    End If
End Sub
