VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form frmHOTFlateRate 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Master Flate Rate"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   11850
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11880
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Height          =   6135
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   12375
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Upload File"
         Height          =   1695
         Left            =   2400
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   8895
         Begin ARProgBarCtrl.ARProgressBar ProgBar 
            Height          =   615
            Left            =   240
            TabIndex        =   43
            Top             =   960
            Visible         =   0   'False
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   1085
            ForeColor       =   65280
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionForeColor=   255
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFF00&
            Caption         =   "PROSES IMPORT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   5520
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton browseDisc 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Browse File DiscRge"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   8
            Top             =   600
            Width           =   2415
         End
         Begin VB.CommandButton browseFrm 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Browse File FrmGene"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   240
            Width           =   2415
         End
         Begin AIFCmp1.asxPowerButton asxPowerButton1 
            Height          =   600
            Left            =   7320
            TabIndex        =   39
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1058
            BackColor       =   16761087
            BorderStyle     =   3
            Caption         =   "BATAL"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "DiscRge : "
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "Frmgene : "
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Height          =   2895
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   11295
         Begin VB.TextBox txtYearLama 
            Height          =   285
            Left            =   2760
            TabIndex        =   42
            Text            =   "Text3"
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtVdsLama 
            Height          =   285
            Left            =   2760
            TabIndex        =   41
            Text            =   "Text3"
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtOpenoLama 
            Height          =   285
            Left            =   2760
            TabIndex        =   40
            Text            =   "Text3"
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin AIFCmp1.asxPowerButton cmdKeluar 
            Height          =   360
            Left            =   9240
            TabIndex        =   15
            Top             =   2400
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   635
            BackColor       =   16744576
            BorderStyle     =   3
            Caption         =   "Kembali Ke &Menu"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
            TextColor       =   16777215
         End
         Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
            Height          =   255
            Left            =   1440
            TabIndex        =   21
            Top             =   960
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   450
            Calculator      =   "frmHOTFlateRate.frx":0000
            Caption         =   "frmHOTFlateRate.frx":0020
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmHOTFlateRate.frx":008C
            Keys            =   "frmHOTFlateRate.frx":00AA
            Spin            =   "frmHOTFlateRate.frx":00F4
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
            MaxValue        =   9999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   "."
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin AIFCmp1.asxPowerButton cmdSimpan 
            Height          =   360
            Left            =   5400
            TabIndex        =   12
            Top             =   2400
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   65280
            BorderStyle     =   3
            Caption         =   "&Tambah"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin AIFCmp1.asxPowerButton cmdUbah 
            Height          =   360
            Left            =   6360
            TabIndex        =   13
            Top             =   2400
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   65535
            BorderStyle     =   3
            Caption         =   "&Ubah"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin AIFCmp1.asxPowerButton cmdHapus 
            Height          =   360
            Left            =   7320
            TabIndex        =   14
            Top             =   2400
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   255
            BorderStyle     =   3
            Caption         =   "&Hapus"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   16
            Top             =   600
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   18
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   2
            Left            =   1440
            TabIndex        =   22
            Top             =   1320
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   ">CCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   3
            Left            =   1440
            TabIndex        =   24
            Top             =   1680
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   29
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   1425
            TabIndex        =   26
            Top             =   2040
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   5
            Left            =   6345
            TabIndex        =   28
            Top             =   240
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   ">CC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   6
            Left            =   6345
            TabIndex        =   30
            Top             =   600
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">C"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   7
            Left            =   6345
            TabIndex        =   32
            Top             =   960
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   ">CC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   6345
            TabIndex        =   34
            Top             =   1320
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   48
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   9
            Left            =   6345
            TabIndex        =   36
            Top             =   1680
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   48
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin AIFCmp1.asxPowerButton cmdImport 
            Height          =   360
            Left            =   8280
            TabIndex        =   38
            Top             =   2400
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   16761087
            BorderStyle     =   3
            Caption         =   "Import"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESC B : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5610
            TabIndex        =   37
            Top             =   1740
            Width           =   705
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESC A : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5610
            TabIndex        =   35
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LANGUAGE : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5310
            TabIndex        =   33
            Top             =   1020
            Width           =   1020
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FRNC : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5775
            TabIndex        =   31
            Top             =   660
            Width           =   555
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6000
            TabIndex        =   29
            Top             =   300
            Width           =   330
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WMI : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   885
            TabIndex        =   27
            Top             =   2100
            Width           =   525
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CLM TYPE : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   25
            Top             =   1740
            Width           =   930
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HOUR : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   810
            TabIndex        =   23
            Top             =   1380
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "YEAR : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   855
            TabIndex        =   20
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VDS : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   660
            Width           =   1305
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OPENO : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   1275
         End
      End
      Begin TrueOleDBGrid70.TDBGrid GridJobOrder 
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   3240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   14933984
         RowDividerColor =   14933984
         RowSubDividerColor=   14933984
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H77EAE1&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Named:id=33:Normal"
         _StyleDefs(41)  =   ":id=33,.parent=0"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(44)  =   ":id=34,.wraptext=-1"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=37:Caption"
         _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(51)  =   "Named:id=38:HighlightRow"
         _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=39:EvenRow"
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
         _StyleDefs(61)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(62)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(63)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(64)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin VB.Label lblJumlah 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "JUMLAH DATA = 0"
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
         Left            =   8760
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   6195
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   10927
      StartColor      =   49152
      EndColor        =   12648447
      FormatString    =   "    Table Master Flate Rate"
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
End
Attribute VB_Name = "frmHOTFlateRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOTFlateRate
' DateTime  : 27/07/2015 11:18
' Author    : RUDI KRISTANTO
' Purpose   :
'---------------------------------------------------------------------------------------

Dim rsGrid As New ADODB.Recordset       ' Recordset Grid
Dim rsTemp As ADODB.Recordset           ' Recordset Proses
Dim cols As TrueOleDBGrid70.Columns
Dim boleh As Boolean
Dim pesan As String
Dim gAtur As Boolean
'Dim sTampung As String

Private Sub ber_sih()
    For i = 0 To 9
        Text(i) = ""
    Next
    TDBNumber1.value = 0
    txtOpenoLama.Text = ""
    txtVdsLama.Text = ""
    txtYearLama.Text = ""
End Sub

Private Sub buka_Tutup(Optional Ket As Boolean = True, Optional stat As Boolean = True)
    Text(0).Enabled = Ket
    Text(1).Enabled = stat
    Text(2).Enabled = stat
    Text(3).Enabled = stat
    Text(4).Enabled = stat
    Text(5).Enabled = stat
    Text(6).Enabled = stat
    Text(7).Enabled = stat
    Text(8).Enabled = stat
    Text(9).Enabled = stat
    TDBNumber1.Enabled = stat
    
End Sub

Private Sub refreshGrid()
    gridJobOrder.ClearFields
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = objO.FetchDataHO("SELECT * FROM MASTER_flaterate WITH(NOLOCK) where ishapus = '0' ORDER BY openo asc ")
    lblJumlah.Caption = "JUMLAH = " & rsGrid.RecordCount
    If Not rsGrid.EOF Then
        gridJobOrder.DataSource = rsGrid
        gridJobOrder.Columns(0).Width = 800
        gridJobOrder.Columns(1).Width = 500
        gridJobOrder.Columns(2).Width = 3500
        gridJobOrder.Columns(3).Width = 3500
        gridJobOrder.Columns(4).Width = 3500
    End If
End Sub

Private Sub cboKM_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub chSBE_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub cmdCetak_Click()
    frmHOTFlateRateCtk.Show
End Sub

Private Sub asxPowerButton1_Click()
    Text1.Text = ""
    Text2.Text = ""
    Call objHT.CloseTransactionHO_err
    ProgBar.Visible = False
    Frame2.Visible = False
End Sub

Private Sub cmdHapus_Click()
    If Trim(Text(0)) <> "" Then
        temp = MsgBox("Yakin akan menghapus data Openo = " & Trim(Text(0)) & " ?", vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE master_FLATERATE SET ishapus = '1', tglupdate = getdate() WHERE [OPENO] = '" & Trim(Text(0)) & "' and VDS = '" & Trim(Text(1)) & "' AND YEAR = '" & TDBNumber1.value & "' and ishapus = '0'")
            Call ber_sih
            Call refreshGrid
        End If
    Else
        MsgBox "Pilih Data yang akan dihapus"
        Exit Sub
    End If
End Sub

Private Sub cmdImport_Click()
    If Frame2.Visible = False Then
        Frame2.Visible = True
    Else
        Frame2.Visible = False
    End If
End Sub


Function pathOfFile(FileName As String) As String
    Dim posn As Integer
    posn = InStrRev(FileName, "\")
    If posn > 0 Then
        pathOfFile = Left$(FileName, posn)
    Else
        pathOfFile = ""
    End If
End Function

Private Sub browseDisc_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.InitDir = App.Path
    CommonDialog1.filter = "Text File (*.txt)|*.txt"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Text2.Text = CommonDialog1.FileName
    End If
End Sub

Private Sub browseFrm_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.InitDir = App.Path
    CommonDialog1.filter = "Text File (*.txt)|*.txt"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Text1.Text = CommonDialog1.FileName
    End If
End Sub

Private Sub cmdImportDiscRge_Click()
              Dim CnExel As New ADODB.Connection
              Dim RcExcel As New ADODB.Recordset
              Dim Jml As Integer
              Dim koneksi As String
              Dim kumpulan, OPENO, id, frnc, lang, desc_a, desc_b As String
              Dim konter As Double
              
1             konter = 0
2             ProgBar.Min = 0
              
3             If Text2.Text <> "" Then
4                 If CnExel.State Then CnExel.Close
                  
                  Dim strConn As String
5                 strConn = "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
                  " Dbq= " & pathOfFile(Text2.Text) & ";Extensions=asc,csv,tab,txt;HDR=NO;Persist Security Info=False"
6                 CnExel.CursorLocation = adUseClient
7                 CnExel.Open strConn
8                 koneksi = "select * from [" & Trim(Mid(Text2.Text, InStrRev(Text2.Text, "\") + 1, Len(Text2.Text))) & "]"
9                 GoSub Proses
10            Else
11                Exit Sub
12            End If
Proses:
13            On Error GoTo cmdImportDiscRge_Click_Error
14            If RcExcel.State Then RcExcel.Close
15            RcExcel.Open koneksi, CnExel, adOpenStatic, adLockReadOnly, adCmdText
16            If RcExcel.RecordCount > 0 Then
17                RcExcel.MoveFirst
18                ProgBar.Visible = True
19                ProgBar.Min = 1
20                ProgBar.Max = RcExcel.RecordCount + 1
21                konter = 1
22                Call objHT.OpenTransactionHO
23                Do While Not RcExcel.EOF
24                    kumpulan = ""
25                    OPENO = ""
26                    id = ""
27                    frnc = ""
28                    lang = ""
29                    desc_a = ""
30                    desc_b = ""
31                    kumpulan = Trim(Nul(RcExcel.Fields(0).value))
                      
32                    id = Trim(Mid(kumpulan, 1, 2))
33                    frnc = Trim(Mid(kumpulan, 3, 1))
34                    OPENO = Trim(Mid(kumpulan, 4, 10))
35                    lang = Trim(Mid(kumpulan, 14, 2))
36                    desc_a = Trim(Mid(kumpulan, 16, 50))
37                    desc_b = Trim(Mid(kumpulan, 66, 50))
                      
38                    Set rsTemp = objHT.FetchDataHO("SELECT * FROM MASTER_FLATERATE WHERE openo = '" & OPENO & "'")
39                    If Not rsTemp.EOF Then
40                        rsTemp.MoveFirst
41                        Do While Not rsTemp.EOF
42                            Call objHT.InsertData_Ho("UPDATE MASTER_FLATERATE SET id = '" & id & "',frnc = '" & frnc & "',language = '" & lang & "',desc_a = '" & desc_a & "',desc_b = '" & desc_b & "' WHERE OPENO = '" & OPENO & "'")
43                            rsTemp.MoveNext
44                        Loop
      '                Else
      '                   Call objHT.InsertData_Ho("INSERT INTO MASTER_FLATERATE VALUES('" & openo & "','','','','','','" & id & "','" & frnc & "','" & lang & "','" & Replace(desc_a, "'", "") & "','" & Replace(desc_b, "'", "") & "',getdate(),0)")
45                    End If
                      
46                    RcExcel.MoveNext
47                    konter = konter + 1
48                    ProgBar.value = konter
49                    DoEvents
                      
50                Loop
51                Call objHT.CloseTransactionHO
52                MsgBox "Data berhasil diimport", vbInformation
53                ProgBar.Visible = False
54            End If
55       On Error GoTo 0
56       Exit Sub

cmdImportDiscRge_Click_Error:
57        Call objHT.CloseTransactionHO_err
58        ProgBar.Visible = False
59        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdImportDiscRge_Click of Form frmHOTFlateRate")
End Sub

Private Function cmdImportFrmGene_Click() As Boolean
              Dim CnExel As New ADODB.Connection
              Dim RcExcel As New ADODB.Recordset
              Dim Jml As Integer
              Dim koneksi As String
              Dim kumpulan, OPENO, VDS, year, hour, clm_type, wmi As String
              Dim konter As Double
              
1             konter = 0
2             ProgBar.Min = 0
              
3             If Text1.Text <> "" Then
4                 If CnExel.State Then CnExel.Close
                  
                  Dim strConn As String
5                 strConn = "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
                  " Dbq= " & pathOfFile(Text1.Text) & ";Extensions=asc,csv,tab,txt;HDR=NO;Persist Security Info=False"
6                 CnExel.CursorLocation = adUseClient
7                 CnExel.Open strConn
8                 koneksi = "select * from [" & Trim(Mid(Text1.Text, InStrRev(Text1.Text, "\") + 1, Len(Text1.Text))) & "]"
9                 GoSub Proses
10            Else
11                Exit Function
12            End If
Proses:
              'On Error GoTo cmdImportFrmGene_Click_Error
13            If RcExcel.State Then RcExcel.Close
14            RcExcel.Open koneksi, CnExel, adOpenStatic, adLockReadOnly, adCmdText
15            If RcExcel.RecordCount > 0 Then
16                RcExcel.MoveFirst
17                ProgBar.Visible = True
18                ProgBar.Min = 1
19                ProgBar.Max = RcExcel.RecordCount + 1
20                konter = 1
21                Call objHT.OpenTransactionHO
22                Do While Not RcExcel.EOF
23                    kumpulan = ""
24                    OPENO = ""
25                    VDS = ""
26                    year = ""
27                    hour = ""
28                    clm_type = ""
29                    wmi = ""
30                    kumpulan = Trim(Nul(RcExcel.Fields(0).value))
                      
31                    OPENO = Trim(Mid(kumpulan, 4, 10))
32                    VDS = Trim(Mid(kumpulan, 14, 6))
33                    year = Trim(Mid(kumpulan, 20, 4))
34                    hour = Trim(Mid(kumpulan, 24, 4))
35                    clm_type = Trim(Mid(kumpulan, 28, 28))
36                    wmi = Trim(Mid(kumpulan, 64, 3))
                      
37                    Set rsTemp = objHT.FetchDataHO("SELECT * FROM MASTER_FLATERATE WHERE openo = '" & OPENO & "' and vds = '" & VDS & "' and year = '" & year & "' and hour = '" & hour & "'")
38                    If rsTemp.EOF Then
      '                If Not rsTemp.EOF Then
      '                    rsTemp.MoveFirst
      '                    Do While Not rsTemp.EOF
      '                        Call objHT.InsertData_Ho("UPDATE ")
      '                        rsTemp.MoveNext
      '                    Loop
      '                Else
39                       Call objHT.InsertData_Ho("INSERT INTO MASTER_FLATERATE VALUES('" & OPENO & "','" & Trim(VDS) & "','" & Trim(year) & "','" & Trim(hour) & "','" & Trim(clm_type) & "','" & Trim(wmi) & "','','','','','',getdate()," & Replace((CDbl(IIf(hour <> "", hour, 0)) / 10), ",", ".") & ",0,'')")
40                       Call objHT.UpdateData_Ho("INSERT INTO TEMPKIRIM VALUES ('M_FRATE','" & OPENO & "|" & Trim(VDS) & Trim(year) & "|" & "','S')")
                         
41                    End If
                      
42                    RcExcel.MoveNext
43                    konter = konter + 1
44                    ProgBar.value = konter
45                    DoEvents
                      
46                Loop
47                Call objHT.CloseTransactionHO
48                MsgBox "Data berhasil diimport", vbInformation
49                ProgBar.Visible = False
50                cmdImportFrmGene_Click = True
51                Exit Function
52            End If
53            cmdImportFrmGene_Click = True
54            Exit Function
55       On Error GoTo 0
56       Exit Function

cmdImportFrmGene_Click_Error:
57        Call objHT.CloseTransactionHO_err
58        ProgBar.Visible = False
59        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdImportFrmGene_Click of Form frmHOTFlateRate")
60        cmdImportFrmGene_Click = False
End Function

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    If cmdSimpan.Caption = "&Tambah" Then
        Call ber_sih
        
        Call buka_Tutup(True, True)
        Text(0).SetFocus
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
        Me.Tag = "baru"
    ElseIf cmdSimpan.Caption = "&Simpan" Then
        boleh = True
        
        ' CHECK DATA KOSONG
        pesan = ""
        
        If Trim(Text(0)) = "" Then
            boleh = False
            pesan = pesan + "Openo. "
        End If
        
        If Trim(Text(1)) = "" Then
            boleh = False
            pesan = pesan + "Vds. "
        End If
        
        If Trim(Text(2)) = "" Then
            boleh = False
            pesan = pesan + "Hour. "
        ElseIf IsNumeric(Trim(Text(2))) = False Then
            boleh = False
            pesan = pesan + "Hour harus angka. "
        End If
        
        If Trim(Text(4)) = "" Then
            boleh = False
            pesan = pesan + "WMI. "
        End If
        
        If Trim(Text(5)) = "" Then
            boleh = False
            pesan = pesan + "ID. "
        End If
        
        If Trim(Text(8)) = "" Then
            boleh = False
            pesan = pesan + "Desc A. "
        End If
        
        
        ' JIKA DATA VALID
        If boleh Then
           
            ' PROSES SIMPAN
            If Trim(Text(0)) = Trim(txtOpenoLama.Text) And Trim(Text(1)) = Trim(txtVdsLama.Text) And Trim(TDBNumber1.value) = Trim(txtYearLama.Text) Then
                Call objO.UpdateDataHO("UPDATE [MASTER_FLATERATE] SET TGLUPDATE = GETDATE(), [OPENO] = '" & Trim(Text(0)) & "',[VDS] = '" & Trim(Text(1)) & "',[YEAR] = '" & TDBNumber1.value & "',[HOUR] = '" & Trim(Text(2)) & "',[CLM_TYPE] = '" & Trim(Text(3)) & "',[WMI] = '" & Trim(Text(4)) & "',[ID] = '" & Trim(Text(5)) & "',[FRNC] = '" & Trim(Text(6)) & "',[LANGUAGE] = '" & Trim(Text(7)) & "',[DESC_A] = '" & Trim(Text(8)) & "',[DESC_B] = '" & Trim(Text(9)) & "',[rate] = " & Replace(IIf(Trim(Text(2)) <> "", (Text(2) / 10), 0), ",", ".") & " Where openo = '" & Trim(Text(0)) & "' AND VDS = '" & Trim(Text(1)) & "' AND YEAR = '" & TDBNumber1.value & "'")
            Else
                If Trim(txtOpenoLama.Text) = "" Or Trim(txtVdsLama.Text) = "" Or Trim(txtYearLama.Text) = "" Then
                    'insert data
                    Call objO.UpdateDataHO("INSERT INTO [MASTER_FLATERATE]([OPENO],[VDS],[YEAR],[HOUR],[CLM_TYPE],[WMI],[ID],[FRNC],[LANGUAGE],[DESC_A],[DESC_B],[TGLENTRY],[rate],tglupdate) VALUES ('" & Trim(Text(0)) & "','" & Trim(Text(1)) & "','" & Trim(TDBNumber1.value) & "','" & Trim(Text(2)) & "','" & Trim(Text(3)) & "','" & Trim(Text(4)) & "','" & Trim(Text(5)) & "','" & Trim(Text(6)) & "','" & Trim(Text(7)) & "','" & Trim(Text(8)) & "','" & Trim(Text(9)) & "',getdate()," & Replace(IIf(Text(2) = "", 0, (Text(2) / 10)), ",", ".") & ",'01/01/1900')")
                Else
                    Call objO.UpdateDataHO("UPDATE [MASTER_FLATERATE] SET TGLUPDATE = GETDATE(), [OPENO] = '" & Trim(Text(0)) & "',[VDS] = '" & Trim(Text(1)) & "',[YEAR] = '" & TDBNumber1.value & "',[HOUR] = '" & Trim(Text(2)) & "',[CLM_TYPE] = '" & Trim(Text(3)) & "',[WMI] = '" & Trim(Text(4)) & "',[ID] = '" & Trim(Text(5)) & "',[FRNC] = '" & Trim(Text(6)) & "',[LANGUAGE] = '" & Trim(Text(7)) & "',[DESC_A] = '" & Trim(Text(8)) & "',[DESC_B] = '" & Trim(Text(9)) & "',[rate] = " & Replace(IIf(Trim(Text(2)) <> "", (Text(2) / 10), 0), ",", ".") & " Where openo = '" & Trim(txtOpenoLama.Text) & "' AND VDS = '" & Trim(txtVdsLama.Text) & "' AND YEAR = '" & txtYearLama.Text & "'")
                End If
            End If
            Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('M_FRATE','" & Trim(Text(0)) & "|" & Trim(Text(1)) & Trim(TDBNumber1.value) & "|" & "','S')")
            Call ber_sih
            Call buka_Tutup(False, False)
            Call refreshGrid
            cmdSimpan.Caption = "&Tambah"
            cmdUbah.Caption = "&Ubah"
        Else
            MsgBox pesan, vbExclamation, "KELENGKAPAN"
        End If
    End If
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        Call buka_Tutup(True, True)
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
        Me.Tag = "ubah"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        cmdSimpan.Caption = "&Tambah"
        cmdUbah.Caption = "&Ubah"
    End If
End Sub

Private Sub Command1_Click()
    If Trim(Text1.Text) = "" And Trim(Text2.Text) = "" Then
        MsgBox "Tidak ada data yang diimport !", vbInformation
        Exit Sub
    Else
        If Text1.Text <> "" Then
            If cmdImportFrmGene_Click = True Then
                Call cmdImportDiscRge_Click
            End If
        Else
            Call cmdImportDiscRge_Click
        End If
        Call refreshGrid
        Text1.Text = ""
        Text2.Text = ""
        Call objHT.CloseTransactionHO_err
        ProgBar.Visible = False
        Frame2.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    
    
    gridJobOrder.Width = Frame1.Width - 200
    gridJobOrder.Height = Frame1.Height - Frame3.Height - 400
    'cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
   
    
    'ProgBar.Top = cmdKeluar.Top
    lblJumlah.Left = Frame1.Left + Frame1.Width - lblJumlah.Width - 1000
    'cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
   
End Sub

Private Sub GridJobOrder_FilterChange()
    On Error GoTo errHandler
    Set cols = gridJobOrder.Columns
    Dim C As Integer
    
    C = gridJobOrder.col
    gridJobOrder.HoldFields
    rsGrid.filter = getFilter()
    gridJobOrder.col = C
    gridJobOrder.EditActive = True
    Exit Sub
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        End If
    Next col
    getFilter = tmp
End Function

Private Sub GridJobOrder_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     If rsGrid.State <> adStateClosed Then
        If Not rsGrid.EOF Then
            Text(0) = rsGrid.Fields(0).value
            txtOpenoLama.Text = rsGrid.Fields(0).value
            Text(1) = rsGrid.Fields(1).value
            txtVdsLama.Text = rsGrid.Fields(1).value
            TDBNumber1.value = rsGrid.Fields(2).value
            txtYearLama.Text = rsGrid.Fields(2).value
            Text(2) = rsGrid.Fields(3).value
            Text(3) = rsGrid.Fields(4).value
            Text(4) = rsGrid.Fields(5).value
            Text(5) = rsGrid.Fields(6).value
            Text(6) = rsGrid.Fields(7).value
            Text(7) = rsGrid.Fields(8).value
            Text(8) = rsGrid.Fields(9).value
            Text(9) = rsGrid.Fields(10).value
            
        End If
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
        
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
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCariHO.Show (vbModal)
            Case 1
                If Text(0) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = Text(0)
                    frmCariHO.Show (vbModal)
                Else
                    MsgBox "Isi Tipe Job Order", vbCritical, "KELENGKAPAN"
                    Text(0).SetFocus
                    Exit Sub
                End If
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub Text_LostFocus(Index As Integer)
    If Index = 0 Then
        If Me.Tag = "baru" Then
            Set rsTemp = objO.FetchDataHO("SELECT ID FROM MASTER_T1 WHERE ISHAPUS = '0' AND ID = '" & Trim(Text(0)) & "'")
            If Not rsTemp.EOF Then
                MsgBox "Id = " & Trim(Text(0)) & " sudah ada silahkan ganti dengan yang lain.", vbInformation
                Text(0) = ""
                Exit Sub
            End If
        End If
    End If
End Sub

