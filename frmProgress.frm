VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmProgress 
   Caption         =   "Progress Pekerjaan"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2490
      TabIndex        =   64
      Top             =   135
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   420
      TabIndex        =   46
      Top             =   2895
      Width           =   11430
      Begin VB.CommandButton TblHisPart 
         Caption         =   "History &Part"
         Height          =   390
         Left            =   3780
         TabIndex        =   93
         Top             =   4200
         Width           =   1890
      End
      Begin VB.CommandButton TbLStart 
         Caption         =   "Star&t"
         Height          =   390
         Left            =   5640
         TabIndex        =   17
         Top             =   4200
         Width           =   1890
      End
      Begin VB.CommandButton TblPause 
         Caption         =   "&Pause"
         Height          =   390
         Left            =   240
         TabIndex        =   18
         Top             =   4200
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.CommandButton TblSelesai 
         Caption         =   "&Selesai"
         Height          =   390
         Left            =   7560
         TabIndex        =   19
         Top             =   4200
         Width           =   1890
      End
      Begin VB.CommandButton TblKeluar 
         Caption         =   "&Kembali ke Menu"
         Height          =   390
         Left            =   9450
         TabIndex        =   20
         Top             =   4200
         Width           =   1890
      End
      Begin VB.Frame Frame5 
         Height          =   3930
         Left            =   3660
         TabIndex        =   61
         Top             =   195
         Width           =   7680
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   37
            Left            =   1200
            TabIndex        =   15
            Top             =   3120
            Width           =   2235
            _ExtentX        =   3942
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   38
            Left            =   1200
            TabIndex        =   16
            Top             =   3435
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   32
            Left            =   3330
            TabIndex        =   9
            Top             =   555
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
            Index           =   33
            Left            =   120
            TabIndex        =   8
            Top             =   555
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   34
            Left            =   5475
            TabIndex        =   12
            Top             =   555
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   35
            Left            =   6420
            TabIndex        =   13
            Top             =   555
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Mask            =   ">CCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   270
            Left            =   3930
            TabIndex        =   75
            Top             =   570
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   270
            Left            =   3225
            TabIndex        =   76
            Top             =   570
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid DGrid 
            Height          =   2025
            Left            =   105
            TabIndex        =   14
            Top             =   930
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   3572
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            ScrollBars      =   0
            SelectionMode   =   1
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   4050
            TabIndex        =   10
            Top             =   555
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   270
            Left            =   4650
            TabIndex        =   89
            Top             =   570
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
            Index           =   15
            Left            =   4755
            TabIndex        =   11
            Top             =   555
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   270
            Left            =   5355
            TabIndex        =   91
            Top             =   570
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Group"
            Height          =   255
            Left            =   4815
            TabIndex        =   92
            Top             =   345
            Width           =   420
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mech"
            Height          =   255
            Left            =   4110
            TabIndex        =   90
            Top             =   345
            Width           =   420
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kondisi"
            Height          =   255
            Left            =   6660
            TabIndex        =   80
            Top             =   345
            Width           =   570
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Rate"
            Height          =   255
            Left            =   5670
            TabIndex        =   79
            Top             =   345
            Width           =   405
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
            Height          =   255
            Left            =   690
            TabIndex        =   78
            Top             =   330
            Width           =   1395
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "JP"
            Height          =   255
            Left            =   3390
            TabIndex        =   77
            Top             =   345
            Width           =   270
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Catatan"
            Height          =   255
            Left            =   135
            TabIndex        =   63
            Top             =   3180
            Width           =   945
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saran"
            Height          =   255
            Left            =   135
            TabIndex        =   62
            Top             =   3480
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3930
         Left            =   105
         TabIndex        =   47
         Top             =   195
         Width           =   3510
         Begin VB.Frame Frame4 
            Caption         =   "Estimasi Labor, Parts & Material"
            Height          =   1650
            Left            =   90
            TabIndex        =   48
            Top             =   2175
            Width           =   3315
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   39
               Left            =   1215
               TabIndex        =   49
               Top             =   300
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   40
               Left            =   1215
               TabIndex        =   33
               Top             =   615
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   41
               Left            =   1215
               TabIndex        =   50
               Top             =   1245
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   45
               Left            =   1215
               TabIndex        =   87
               Top             =   930
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   ">"
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Material"
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   975
               Width           =   945
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   1290
               Width           =   945
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Part"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   660
               Width           =   945
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Labour"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   345
               Width           =   945
            End
         End
         Begin VB.ComboBox cboStatus 
            Height          =   315
            ItemData        =   "frmProgress.frx":0000
            Left            =   1260
            List            =   "frmProgress.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   210
            Width           =   735
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   42
            Left            =   1245
            TabIndex        =   54
            Top             =   1815
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   43
            Left            =   2325
            TabIndex        =   55
            Top             =   1800
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Index           =   44
            Left            =   2805
            TabIndex        =   56
            Top             =   1800
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Index           =   30
            Left            =   1260
            TabIndex        =   7
            Top             =   1185
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
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
            Mask            =   ">CCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   31
            Left            =   1260
            TabIndex        =   5
            Top             =   870
            Width           =   570
            _ExtentX        =   1005
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
            Mask            =   ">CCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   1260
            TabIndex        =   6
            Top             =   555
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Index           =   13
            Left            =   1260
            TabIndex        =   32
            Top             =   1500
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   270
            Left            =   1860
            TabIndex        =   86
            Top             =   885
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "T SA."
            Height          =   255
            Left            =   180
            TabIndex        =   74
            Top             =   1545
            Width           =   945
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "GRP"
            Height          =   255
            Left            =   165
            TabIndex        =   73
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inv. No."
            Height          =   255
            Left            =   165
            TabIndex        =   60
            Top             =   1845
            Width           =   945
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   255
            Left            =   180
            TabIndex        =   59
            Top             =   285
            Width           =   945
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "K. Meter"
            Height          =   255
            Left            =   165
            TabIndex        =   58
            Top             =   1245
            Width           =   945
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "MCH"
            Height          =   255
            Left            =   165
            TabIndex        =   57
            Top             =   915
            Width           =   945
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   420
      TabIndex        =   35
      Top             =   480
      Width           =   11430
      Begin VB.Frame Frame6 
         Height          =   2175
         Left            =   8565
         TabIndex        =   65
         Top             =   120
         Width           =   2760
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   66
            Top             =   360
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Mask            =   ">99-99-9999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   6
            Left            =   1200
            TabIndex        =   68
            Top             =   675
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Mask            =   ">99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   7
            Left            =   1200
            TabIndex        =   1
            Top             =   990
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
            Mask            =   ">99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   1200
            TabIndex        =   2
            Top             =   1305
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
            Mask            =   ">99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   9
            Left            =   1200
            TabIndex        =   3
            Top             =   1620
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
            Mask            =   ">99:99"
            PromptChar      =   ">"
         End
         Begin VB.Label Label16 
            Caption         =   "(HH:MM)"
            Height          =   240
            Left            =   1965
            TabIndex        =   85
            Top             =   1665
            Width           =   720
         End
         Begin VB.Label Label15 
            Caption         =   "(HH:MM)"
            Height          =   240
            Left            =   1965
            TabIndex        =   84
            Top             =   1365
            Width           =   690
         End
         Begin VB.Label Label20 
            Caption         =   "(HH:MM)"
            Height          =   240
            Left            =   1965
            TabIndex        =   83
            Top             =   1050
            Width           =   660
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Actual"
            Height          =   255
            Left            =   165
            TabIndex        =   72
            Top             =   1665
            Width           =   945
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Selesai"
            Height          =   255
            Left            =   165
            TabIndex        =   71
            Top             =   1350
            Width           =   945
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai"
            Height          =   255
            Left            =   165
            TabIndex        =   70
            Top             =   1035
            Width           =   945
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Terima"
            Height          =   255
            Left            =   165
            TabIndex        =   69
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   255
            Left            =   150
            TabIndex        =   67
            Top             =   405
            Width           =   945
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   21
         Top             =   315
         Width           =   3300
         _ExtentX        =   5821
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
         Index           =   2
         Left            =   960
         TabIndex        =   22
         Top             =   630
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   3
         Left            =   960
         TabIndex        =   23
         Top             =   945
         Width           =   3300
         _ExtentX        =   5821
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
         Left            =   960
         TabIndex        =   24
         Top             =   1260
         Width           =   3300
         _ExtentX        =   5821
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
         Index           =   10
         Left            =   960
         TabIndex        =   25
         Top             =   1575
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   11
         Left            =   960
         TabIndex        =   26
         Top             =   1890
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   5700
         TabIndex        =   27
         Top             =   315
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   17
         Left            =   5700
         TabIndex        =   28
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   18
         Left            =   5700
         TabIndex        =   30
         Top             =   1590
         Width           =   2250
         _ExtentX        =   3969
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
         Index           =   25
         Left            =   5700
         TabIndex        =   31
         Top             =   1905
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   24
         Left            =   7545
         TabIndex        =   29
         Top             =   945
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   7425
         TabIndex        =   81
         Top             =   960
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
         Index           =   19
         Left            =   5700
         TabIndex        =   82
         Top             =   1275
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
         Height          =   255
         Left            =   4605
         TabIndex        =   45
         Top             =   1950
         Width           =   945
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         Height          =   255
         Left            =   4605
         TabIndex        =   44
         Top             =   1635
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Left            =   4605
         TabIndex        =   43
         Top             =   990
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         Height          =   255
         Left            =   4605
         TabIndex        =   42
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disc."
         Height          =   255
         Left            =   420
         TabIndex        =   41
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         Height          =   255
         Left            =   435
         TabIndex        =   40
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telp. No."
         Height          =   255
         Left            =   -135
         TabIndex        =   39
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   255
         Left            =   -135
         TabIndex        =   38
         Top             =   990
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   255
         Left            =   -135
         TabIndex        =   37
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   -135
         TabIndex        =   36
         Top             =   360
         Width           =   945
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1470
      TabIndex        =   0
      Top             =   120
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
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
      Mask            =   ">CCCCCCCC"
      PromptChar      =   ">"
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7560
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13335
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Progress Pekerjaan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. WO"
      Height          =   255
      Left            =   375
      TabIndex        =   34
      Top             =   165
      Width           =   945
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Baris As Integer
Dim RateCPI As Double
Dim AdaJP As Boolean
Dim OrReturn As Boolean
Dim Simpan As Boolean
Dim JamTemp As Boolean
Const MaxBrs = 7

Private Function TotalJasa() As Double
    Dim Temp As Double
    
    Temp = 0
    With DGrid
        For i = 1 To 7
            If .TextMatrix(i, 1) = "" Then
                Exit For
            Else
                Temp = Temp + Koma2Titik(.TextMatrix(i, 5)) * RateCPI
            End If
        Next
    End With
    TotalJasa = Temp
End Function

Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
    Dim aa As String
    Dim xHasil As String
    xHasil = ""
    If xUpdate Then
        xHasil = "UPDATE " + xTabel + " SET "
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC", "RTE1", "RTE2", "RTE3", "RTE4", "RTE5", "RTE6", "RTE7", "RTE8", "LABOR", "PART", "MTRL", "TOTAL", "MET", "OIL", "SBLT", "BODY", "DISCT"
                xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
            Case Else
                xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
            End Select
            i = i + 1
        Loop
    Else
        xHasil = "INSERT INTO " + xTabel + " ("
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            xHasil = xHasil + gsSql1(i)
            i = i + 1
        Loop
        xHasil = xHasil + ") VALUES ("
        
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC", "RTE1", "RTE2", "RTE3", "RTE4", "RTE5", "RTE6", "RTE7", "RTE8", "LABOR", "PART", "MTRL", "TOTAL", "MET", "OIL", "SBLT", "BODY", "DISCT"
                xHasil = xHasil + gsSql2(i)
            Case Else
                xHasil = xHasil + "'" + gsSql2(i) + "'"
            End Select
            i = i + 1
        Loop
        xHasil = xHasil + ")"
    End If
    GensSql = xHasil
End Function

Private Sub hapusForm()
    For i = 1 To 45
        Select Case i
            Case 14, 21, 22, 23, 20, 26, 27, 28, 29, 32, 36
            Case 5
                Text(i) = ">>->>->>>>"
            Case 6, 7, 8, 9
                Text(i) = ">>:>>"
            Case Else
                Text(i) = ""
        End Select
    Next
'    cboStatus.Text = ""
    
    Baris = 1
    For i = 1 To 7
        DGrid.TextMatrix(i, 1) = ""
        DGrid.TextMatrix(i, 2) = ""
        DGrid.TextMatrix(i, 3) = ""
        DGrid.TextMatrix(i, 4) = ""
        DGrid.TextMatrix(i, 5) = ""
        DGrid.TextMatrix(i, 6) = ""
    Next
End Sub


Private Sub cboStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombolCombo(KeyCode)
End Sub

Private Sub DGrid_DblClick()
    If Baris > 1 And DGrid.TextMatrix(DGrid.row, 1) <> "" Then
        A = DGrid.row
        'AMBIL
        With DGrid
            Text(33) = .TextMatrix(A, 1)
            Text(32) = .TextMatrix(A, 2)
            Text(14) = .TextMatrix(A, 3)
            Text(15) = .TextMatrix(A, 4)
            Text(34) = .TextMatrix(A, 5)
            Text(35) = .TextMatrix(A, 6)
        End With
        'HAPUS
        For i = A To Baris
            With DGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
                .TextMatrix(i, 5) = .TextMatrix(i + 1, 5)
                .TextMatrix(i, 6) = .TextMatrix(i + 1, 6)
            End With
        Next
        Baris = Baris - 1
        Text(33).SetFocus
        Text(39) = TotalJasa
        Call HitungTotal
    End If
End Sub

Private Sub dGrid_GotFocus()
    If Trim(Text(33)) <> "" Then
        If MaxBrs >= Baris Then
            With DGrid
                .TextMatrix(Baris, 1) = Text(33)
                .TextMatrix(Baris, 2) = Text(32)
                .TextMatrix(Baris, 3) = Text(14)
                .TextMatrix(Baris, 4) = Text(15)
                .TextMatrix(Baris, 5) = Format(Koma2Titik(Text(34)), "#,##0.00")
                .TextMatrix(Baris, 6) = Text(35)
            End With
            Baris = Baris + 1
            Text(33) = ""
            Text(32) = ""
            Text(14) = ""
            Text(15) = ""
            Text(34) = ""
            Text(35) = ""
            Text(33).SetFocus
            Text(39) = TotalJasa
            Call HitungTotal
        Else
            MsgBox "Maksimal 7 !"
            Text(33).SetFocus
        End If
    End If
End Sub

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call DGrid_DblClick
    End If
End Sub

Private Sub Form_Load()
    Baris = 1
    DGrid.Rows = 11
    DGrid.Cols = 7
    DGrid.ColWidth(0) = 0
    DGrid.ColWidth(1) = 3100
    DGrid.ColWidth(2) = 750
    DGrid.ColWidth(3) = 750
    DGrid.ColWidth(4) = 750
    DGrid.ColWidth(5) = 900
    DGrid.ColWidth(6) = 1100
    DGrid.ColAlignment(1) = flexAlignLeftCenter
    DGrid.ColAlignment(2) = flexAlignLeftCenter
    DGrid.ColAlignment(3) = flexAlignLeftCenter
    DGrid.ColAlignment(4) = flexAlignLeftCenter
    DGrid.ColAlignment(5) = flexAlignRightCenter
    DGrid.ColAlignment(6) = flexAlignLeftCenter
    DGrid.TextMatrix(0, 1) = "Keterangan"
    DGrid.TextMatrix(0, 2) = "JP"
    DGrid.TextMatrix(0, 3) = "Mech"
    DGrid.TextMatrix(0, 4) = "Group"
    DGrid.TextMatrix(0, 5) = "Rate"
    DGrid.TextMatrix(0, 6) = "Kondisi"
    Call hapusForm
    Simpan = True
End Sub

Private Sub TblHisPart_Click()
    If Trim(Text(0)) <> "" Then
        gNoPolisi = Text(16)
        gCust = Text(1)
            frmHistoryPart.Show (vbModal)
    Else
        MsgBox "Masukkan No. WO !", vbOKOnly, "Perhatian"
    End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub SaveOrderPekerjaan(kondisinya As String, NomorWO As String, NomorPolisi As String)
    Dim KaliRate As Double
    
    Select Case Text(24)
        Case "COM"
            KaliRate = gRateCOM
        Case "PAS"
            KaliRate = gRatePAS
        Case "IRC"
            KaliRate = gRateIRC
        Case Else
            KaliRate = 0
    End Select
    
    sSQL = "Select * from JOBDTA where NO_WO ='" & NomorWO & "' and NO_POLISI ='" & NomorPolisi & "' and DELIV<>'BTL' AND RPR_COND <> 'K'"
    oRs.Open sSQL, oConn, adOpenStatic
    
    Call ClearsSql
    gsSql1(0) = "KM"
    gsSql2(0) = Text(30)
    gsSql1(1) = "TIME_M"
    gsSql2(1) = Text(7)
    gsSql1(2) = "TIME_S"
    gsSql2(2) = Text(8)
    gsSql1(3) = "MECH"
    gsSql2(3) = Text(31)
    gsSql1(4) = "[GROUP]"
    gsSql2(4) = Text(12)
    gsSql1(5) = "STATUS"
    gsSql2(5) = cboStatus.Text
    gsSql1(6) = "REMARK1"
    gsSql2(6) = Text(37)
    gsSql1(7) = "SUGGEST"
    gsSql2(7) = Text(38)
    gsSql1(8) = "RPR_COND"
    gsSql2(8) = Trim(kondisinya)
        
    gsSql1(9) = "REP1"
    gsSql2(9) = DGrid.TextMatrix(1, 1)
    gsSql1(10) = "REP2"
    gsSql2(10) = DGrid.TextMatrix(2, 1)
    gsSql1(11) = "REP3"
    gsSql2(11) = DGrid.TextMatrix(3, 1)
    gsSql1(12) = "REP4"
    gsSql2(12) = DGrid.TextMatrix(4, 1)
    gsSql1(13) = "REP5"
    gsSql2(13) = DGrid.TextMatrix(5, 1)
    gsSql1(14) = "REP6"
    gsSql2(14) = DGrid.TextMatrix(6, 1)
    gsSql1(15) = "REP7"
    gsSql2(15) = DGrid.TextMatrix(7, 1)
    gsSql1(16) = "RTE1"
    gsSql2(16) = SimpanKoma2Titik(DGrid.TextMatrix(1, 5))
    gsSql1(17) = "RTE2"
    gsSql2(17) = SimpanKoma2Titik(DGrid.TextMatrix(2, 5))
    gsSql1(18) = "RTE3"
    gsSql2(18) = SimpanKoma2Titik(DGrid.TextMatrix(3, 5))
    gsSql1(19) = "RTE4"
    gsSql2(19) = SimpanKoma2Titik(DGrid.TextMatrix(4, 5))
    gsSql1(20) = "RTE5"
    gsSql2(20) = SimpanKoma2Titik(DGrid.TextMatrix(5, 5))
    gsSql1(21) = "RTE6"
    gsSql2(21) = SimpanKoma2Titik(DGrid.TextMatrix(6, 5))
    gsSql1(22) = "RTE7"
    gsSql2(22) = SimpanKoma2Titik(DGrid.TextMatrix(7, 5))
    
    gsSql1(23) = "KLPKERJA1"
    gsSql2(23) = DGrid.TextMatrix(1, 2)
    gsSql1(24) = "KLPKERJA2"
    gsSql2(24) = DGrid.TextMatrix(2, 2)
    gsSql1(25) = "KLPKERJA3"
    gsSql2(25) = DGrid.TextMatrix(3, 2)
    gsSql1(26) = "KLPKERJA4"
    gsSql2(26) = DGrid.TextMatrix(4, 2)
    gsSql1(27) = "KLPKERJA5"
    gsSql2(27) = DGrid.TextMatrix(5, 2)
    gsSql1(28) = "KLPKERJA6"
    gsSql2(28) = DGrid.TextMatrix(6, 2)
    gsSql1(29) = "KLPKERJA7"
    gsSql2(29) = DGrid.TextMatrix(7, 2)
    gsSql1(30) = "COND1"
    gsSql2(30) = DGrid.TextMatrix(1, 6)
    gsSql1(31) = "COND2"
    gsSql2(31) = DGrid.TextMatrix(2, 6)
    gsSql1(32) = "COND3"
    gsSql2(32) = DGrid.TextMatrix(3, 6)
    gsSql1(33) = "COND4"
    gsSql2(33) = DGrid.TextMatrix(4, 6)
    gsSql1(34) = "COND5"
    gsSql2(34) = DGrid.TextMatrix(5, 6)
    gsSql1(35) = "COND6"
    gsSql2(35) = DGrid.TextMatrix(6, 6)
    gsSql1(36) = "COND7"
    gsSql2(36) = DGrid.TextMatrix(7, 6)
    gsSql1(37) = "TIME_A"
    gsSql2(37) = Text(9)
    gsSql1(38) = "MEKANIK1"
    gsSql2(38) = DGrid.TextMatrix(1, 3)
    gsSql1(39) = "MEKANIK2"
    gsSql2(39) = DGrid.TextMatrix(2, 3)
    gsSql1(40) = "MEKANIK3"
    gsSql2(40) = DGrid.TextMatrix(3, 3)
    gsSql1(41) = "MEKANIK4"
    gsSql2(41) = DGrid.TextMatrix(4, 3)
    gsSql1(42) = "MEKANIK5"
    gsSql2(42) = DGrid.TextMatrix(5, 3)
    gsSql1(43) = "MEKANIK6"
    gsSql2(43) = DGrid.TextMatrix(6, 3)
    gsSql1(44) = "MEKANIK7"
    gsSql2(44) = DGrid.TextMatrix(7, 3)
    gsSql1(45) = "GRP1"
    gsSql2(45) = DGrid.TextMatrix(1, 4)
    gsSql1(46) = "GRP2"
    gsSql2(46) = DGrid.TextMatrix(2, 4)
    gsSql1(47) = "GRP3"
    gsSql2(47) = DGrid.TextMatrix(3, 4)
    gsSql1(48) = "GRP4"
    gsSql2(48) = DGrid.TextMatrix(4, 4)
    gsSql1(49) = "GRP5"
    gsSql2(49) = DGrid.TextMatrix(5, 4)
    gsSql1(50) = "GRP6"
    gsSql2(50) = DGrid.TextMatrix(6, 4)
    gsSql1(51) = "GRP7"
    gsSql2(51) = DGrid.TextMatrix(7, 4)
        
    If Not oRs.BOF Then
        If oRs("FN_RETURN") = 1 Then
            gsSql1(52) = "OFF1"
            gsSql2(52) = 0
            gsSql1(53) = "OFF2"
            gsSql2(53) = 0
            gsSql1(54) = "OFF3"
            gsSql2(54) = 0
            gsSql1(55) = "OFF4"
            gsSql2(55) = 0
            gsSql1(56) = "OFF5"
            gsSql2(56) = 0
            gsSql1(57) = "OFF6"
            gsSql2(57) = 0
            gsSql1(58) = "OFF7"
            gsSql2(58) = 0
        Else
            gsSql1(52) = "OFF1"
            gsSql2(52) = SimpanKoma2Titik(Round(Koma2Titik(DGrid.TextMatrix(1, 5)) * KaliRate, 2))
            gsSql1(53) = "OFF2"
            gsSql2(53) = SimpanKoma2Titik(Round(Koma2Titik(DGrid.TextMatrix(2, 5)) * KaliRate, 2))
            gsSql1(54) = "OFF3"
            gsSql2(54) = SimpanKoma2Titik(Round(Koma2Titik(DGrid.TextMatrix(3, 5)) * KaliRate, 2))
            gsSql1(55) = "OFF4"
            gsSql2(55) = SimpanKoma2Titik(Round(Koma2Titik(DGrid.TextMatrix(4, 5)) * KaliRate, 2))
            gsSql1(56) = "OFF5"
            gsSql2(56) = SimpanKoma2Titik(Round(Koma2Titik(DGrid.TextMatrix(5, 5)) * KaliRate, 2))
            gsSql1(57) = "OFF6"
            gsSql2(57) = SimpanKoma2Titik(Round(Koma2Titik(DGrid.TextMatrix(6, 5)) * KaliRate, 2))
            gsSql1(58) = "OFF7"
            gsSql2(58) = SimpanKoma2Titik(Round(Koma2Titik(DGrid.TextMatrix(7, 5)) * KaliRate, 2))
        End If
        
        sSQL = GensSql("JOBDTA", True) + " where NO_WO ='" & NomorWO & "' and NO_POLISI ='" & NomorPolisi & "' and DELIV<>'BTL' AND RPR_COND <> 'K'"
    Else
        sSQL = GensSql("JOBDTA", False)
    End If
    oConn.Execute sSQL
    oRs.Close
    Simpan = True
End Sub

Private Sub SimpanData(Kode As Byte, Kondisi As String)
    Dim boleh As Boolean
    Dim pesan As String
    
    boleh = True
    
    If (CekJam(Text(7)) = False Or Text(7) = ">>:>>") And Kode = 1 Then
        pesan = pesan + " Jam Mulai"
        boleh = False
        If vbYes = MsgBox("Jam Mulai masih kosong, diisi jam sekarang? ", vbYesNo, "Peringatan") Then
            Text(7) = Format(Time, "HH:NN")
            pesan = ""
            boleh = True
        End If
    End If
    If (CekJam(Text(8)) = False Or Text(8) = ">>:>>") And Kode = 3 Then
        pesan = pesan + " Jam Selesai"
        boleh = False
        If vbYes = MsgBox("Jam Selesai masih kosong, diisi jam sekarang? ", vbYesNo, "Peringatan") Then
            Text(8) = Format(Time, "HH:NN")
            pesan = ""
            boleh = True
        End If
    End If
    If Trim(Text(9)) = "" Then
        pesan = pesan + " Jam Actual"
        boleh = False
    End If
    If Trim(Text(31)) = "" Then
        pesan = pesan + " MECH"
        boleh = False
    End If
    If Trim(Text(30)) = "" Then
        pesan = pesan + " K Meter"
        boleh = False
    End If
    
    If boleh = True Then
        Call SaveOrderPekerjaan(Kondisi, Text(0), Text(16))
        Call hapusForm
        Text(0) = ""
        Text(0).SetFocus
    Else
        MsgBox "Masukkan" + pesan + "!"
    End If
End Sub

Private Sub TblPause_Click()
    Call SimpanData(2, "B")
End Sub

Private Sub TblSelesai_Click()
    Call SimpanData(3, "S")
End Sub

Private Sub TbLStart_Click()
    Call SimpanData(1, "P")
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
    
    Select Case Index
        Case 14, 15, 34, 35
            Call ItemRegular
    End Select
        
End Sub

Private Sub ItemRegular()
    If gIERC = True Then
        If Text(32) = "IRC" Or Text(32) = "ERC" Then
            sSQL = "Select COUNT(*) AS JML from T_JOBLIST where FS_ITEM ='" & Trim(Text(33)) & "'"
            oRs.Open sSQL, oConn, adOpenStatic
            If oRs!JML = 0 Then
                MsgBox "IRC & ERC harus sesuai tabel tekan F12 untuk memilih !", vbOKOnly, "Kesalahan"
                Text(33).SetFocus
            End If
            oRs.Close
        End If
    End If
End Sub
Private Sub PanggilDataOrder(NomorWO As String)
    On Error Resume Next
    Dim rsHlp2 As New ADODB.Recordset
    
    JamTemp = False
    Text(7).BackColor = &H80000005
    sSQL = "Select * from JOBDTA where NO_WO ='" & NomorWO & "' and DELIV<>'BTL'"
    rsHlp2.Open sSQL, oConn, adOpenStatic
    If Not rsHlp2.BOF Then
        'Data Customer & Order
        Text(1) = rsHlp2("CUSTOMER")
        Text(2) = rsHlp2("CUST_ADRES")
        Text(3) = rsHlp2("KOTA")
        Text(4) = rsHlp2("TELPON")
        Text(16) = rsHlp2("NO_POLISI")
        Text(17) = rsHlp2("MDL")
        Text(24) = rsHlp2("MODEL")
        Text(18) = rsHlp2("COLOR")
        Text(13) = rsHlp2("FRONT")
        Text(37) = rsHlp2("REMARK1")
        Text(10) = rsHlp2("PPN")
        Text(11) = rsHlp2("DISC")
        Text(12) = rsHlp2("GROUP")
        Text(19) = rsHlp2("FMDL")
        Text(45) = rsHlp2("MTRL")
        
        Text(42) = rsHlp2("NOINV")
        Text(31) = rsHlp2("MECH")
        Text(30) = rsHlp2("KM")
        Text(38) = rsHlp2("SUGGEST")
        Text(39) = rsHlp2("LABOR")
        Text(40) = rsHlp2("PART")
        Text(41) = rsHlp2("TOTAL")
        Text(44) = rsHlp2("PAID")
        Text(43) = rsHlp2("DELIV")
        cboStatus.Text = rsHlp2("STATUS")
'        Text(32) = rsHlp2("JOB")
        Text(5) = TglDDMMYYYY(rsHlp2("DATE"))
        Text(6) = rsHlp2("TIME")
        
        If rsHlp2("TIME_M") <> "" Then
            Text(7) = rsHlp2("TIME_M")
            If Format(Text(7), "HH:NN") > Format(Time, "HH:NN") Then
                JamTemp = True
                Text(7).BackColor = &HC0E0FF
            End If
        End If
        If rsHlp2("TIME_S") <> "" Then
            Text(8) = rsHlp2("TIME_S")
        End If
        If rsHlp2("TIME_A") <> "" Then
            Text(9) = rsHlp2("TIME_A")
        End If
        
        
''Format(Time, "HH:NN")

        Baris = 1
        For i = 1 To 7
            DGrid.TextMatrix(i, 1) = rsHlp2("REP" & Trim(str(i)))
            DGrid.TextMatrix(i, 2) = rsHlp2("KLPKERJA" & Trim(str(i)))
            DGrid.TextMatrix(i, 3) = rsHlp2("MEKANIK" & Trim(str(i)))
            DGrid.TextMatrix(i, 4) = rsHlp2("GRP" & Trim(str(i)))
            DGrid.TextMatrix(i, 5) = Format(rsHlp2("RTE" & Trim(str(i))), "#,##0.00")
            DGrid.TextMatrix(i, 6) = rsHlp2("COND" & Trim(str(i)))
        Next
        
        For i = 1 To 7
            If DGrid.TextMatrix(i, 1) <> "" Then
                Baris = Baris + 1
            End If
        Next

    End If
    rsHlp2.Close
    
    sSQL = "select * from MODEL where MDL ='" & Text(17) & "'"
    rsHlp2.Open sSQL, oConn, adOpenStatic
    If rsHlp2.BOF Then
        gGroupModel = ""
    Else
        gGroupModel = rsHlp2("FS_GRP_MDL")
    End If
    rsHlp2.Close
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "PROGRESS"
                frmCari.Show (vbModal)
            Case 14
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                frmCari.Show (vbModal)
            Case 15
                gMainKeyCari = Text(Index)
                gSearchOption = "GROUP_MEKANIK"
                frmCari.Show (vbModal)
            Case 17
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.Show (vbModal)
            Case 31
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                frmCari.Show (vbModal)
            Case 32
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCari.Show (vbModal)
            Case 33
                gMainKeyCari = Text(Index)
                gSearchOption = "JOBLIST"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
    Select Case Index
        Case 0
        Case 34
            KeyAscii = CekNumerik(KeyAscii)
            Simpan = False
        Case Else
            Simpan = False
    End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Dim NoWO As String
    Dim boleh As Boolean
    
    Select Case Index
        Case 0
            OrReturn = False
            If Trim(Text(Index)) <> "" Then
                boleh = True
                If Simpan = False Then
                    If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                        boleh = False
                    End If
                End If
                If boleh Then
                    NoWO = Text(0)
                    sSQL = "Select * from JOBDTA where NO_WO ='" & NoWO & "' and DELIV<>'BTL' AND RPR_COND <> 'K'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.BOF Then
                        Cancel = True
                        gMainKeyCari = Text(Index)
                        gSearchOption = "PROGRESS"
                        frmCari.Show (vbModal)
                    Else
                        Call PanggilDataOrder(NoWO)
                    End If
                    oRs.Close
                    gMDL = Text(17)
                    Call HitungNilaiKaliRate
                    Call HitungTotal
                End If
            Else
                Call hapusForm
            End If
            
'            If Trim(Text(Index)) <> "" Then
'                NoWO = Text(0)
'                sSQL = "Select * from JOBDTA where NO_WO ='" & NoWO & "' and DELIV='PRC' AND RPR_COND <> 'K'"
'                oRs.Open sSQL, oConn, adOpenStatic
'                If oRs.BOF Then
'                    Cancel = True
'                    gMainKeyCari = Text(Index)
'                    gSearchOption = "PROGRESS"
'                    frmCari.Show (vbModal)
'                Else
'                    Call PanggilDataOrder(NoWO)
'                End If
'                oRs.Close
'                gMDL = Text(17)
'                Call HitungNilaiKaliRate
'                Call HitungTotal
'            End If
        Case 7
            If JamTemp = False Then
                If CekJam(Text(Index)) = False Then
                    Text(Index) = Format(Time, "HH:NN")
                Else
                    If Format(Text(7), "HH:NN") > Format(Time, "HH:NN") Then
                        Text(7) = ">>:>>"
                    End If
                End If
            End If
        Case 8
           If CekJam(Text(Index)) = False Then
                Text(Index) = Format(Time, "HH:NN")
            Else
                If Format(Text(8), "HH:NN") > Format(Time, "HH:NN") Then
                    Text(8) = ">>:>>"
'                    Text(8) = Format(Time, "HH:NN")
                End If
            End If
            Text(9) = KurangM1M2(Text(7), Text(8))
            
        Case 14
            If Trim(Text(Index)) <> "" Then
                sSQL = "Select * from MEKANIK where MECH ='" & Text(Index) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MEKANIK"
                    frmCari.Show (vbModal)
                Else
                    Text(15) = oRs("GROUP")
                End If
                oRs.Close
            Else
                Text(Index) = Text(31)
                sSQL = "Select * from MEKANIK where MECH ='" & Text(Index) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If Not oRs.BOF Then
                    Text(15) = oRs("GROUP")
                End If
                oRs.Close
            End If
            
'            If Trim(Text(Index)) <> "" Then
'                sSQL = "Select * from MEKANIK where MECH ='" & Text(Index) & "'"
'                oRs.Open sSQL, oConn, adOpenStatic
'                If oRs.BOF Then
'                    Cancel = True
'                    gMainKeyCari = Text(Index)
'                    gSearchOption = "MEKANIK"
'                    frmCari.Show (vbModal)
'                Else
'                    Text(15) = oRs("GROUP")
'                End If
'                oRs.Close
'            End If
        
        Case 17
            sSQL = "Select * from MODEL where MDL ='" & Text(17) & "'"
            oRs.Open sSQL, oConn, adOpenStatic
            If oRs.BOF Then
                Cancel = True
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.Show (vbModal)
                gGroupModel = ""
            Else
                Text(17) = oRs("MDL")
                Text(19) = oRs("Desc")
                Text(24) = oRs("MODEL")
                gGroupModel = oRs("FS_GRP_MDL")
            End If
            oRs.Close
            gMDL = Text(17)
            Call HitungNilaiKaliRate
        Case 31
            If Trim(Text(Index)) <> "" Then
                sSQL = "Select * from MEKANIK where MECH ='" & Text(Index) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MEKANIK"
                    frmCari.Show (vbModal)
                Else
                    Text(12) = oRs("GROUP")
                End If
                oRs.Close
            End If

        Case 32
            If Trim(Text(Index)) <> "" Or Trim(Text(33)) <> "" Then
                sSQL = "Select * from JOB where JOB='" & Text(Index) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCari.Show (vbModal)
                End If
                oRs.Close
            End If

'        Case 32
'            If AdaJP Then
'                DGrid.SetFocus
'                AdaJP = False
'            End If
        Case 33
            If Trim(Text(Index)) <> "" Then
'                sSQL = "Select * from T_JOBLIST where FS_ITEM='" & Text(33) & "' and FS_MDL='" & gMDL & "'"
                sSQL = "select * from T_JOBLIST where FS_ITEM='" & Text(33) & "' and FS_GRP_MDL='" & gGroupModel & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If Not oRs.BOF Then
                    Text(34) = oRs("FN_RATE")
                    Text(32) = oRs("FS_JNS_KERJA")
                    AdaJP = True
                Else
                    AdaJP = False
'                    Text(34) = ""
                End If
                oRs.Close
            End If
        Case 34
            If Trim(Text(Index)) = "" Then
                Text(Index) = 0
            End If
    End Select
End Sub

Private Sub HitungTotal()
    Text(41) = Koma2Titik(Text(39)) + Koma2Titik(Text(40)) + Koma2Titik(Text(45))
End Sub
Private Sub HitungNilaiKaliRate()
    Select Case Trim(Text(24))
        Case "COM"
            RateCPI = gRateCOM
        Case "PAS"
            RateCPI = gRatePAS
        Case "IRC"
            RateCPI = gRateIRC
        Case Else
            RateCPI = gRateCOM
    End Select
End Sub

