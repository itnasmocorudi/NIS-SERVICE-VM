VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmLihatOrder 
   Caption         =   "Melihat Order Pekerjaan"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2400
      TabIndex        =   87
      Top             =   60
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Input Order Pekerjaan"
      Height          =   4215
      Left            =   405
      TabIndex        =   64
      Top             =   3300
      Width           =   11415
      Begin VB.CommandButton TblKeluar 
         Caption         =   "&Kembali ke Menu"
         Height          =   390
         Left            =   9465
         TabIndex        =   36
         Top             =   3765
         Width           =   1890
      End
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         Height          =   3555
         Left            =   5595
         TabIndex        =   84
         Top             =   180
         Width           =   5760
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   37
            Left            =   1350
            TabIndex        =   34
            Top             =   2715
            Width           =   2235
            _ExtentX        =   3942
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
            Index           =   38
            Left            =   1350
            TabIndex        =   35
            Top             =   3030
            Width           =   2775
            _ExtentX        =   4895
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
         Begin MSFlexGridLib.MSFlexGrid DGrid 
            Height          =   2025
            Left            =   60
            TabIndex        =   33
            Top             =   510
            Width           =   5640
            _ExtentX        =   9948
            _ExtentY        =   3572
            _Version        =   393216
            FixedCols       =   0
            ScrollBars      =   0
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pesan Khusus"
            Height          =   255
            Left            =   -345
            TabIndex        =   86
            Top             =   2775
            Width           =   1575
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saran"
            Height          =   255
            Left            =   255
            TabIndex        =   85
            Top             =   3075
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   3540
         Left            =   45
         TabIndex        =   65
         Top             =   180
         Width           =   5535
         Begin VB.Frame Frame4 
            Caption         =   "Estimasi Labor, Parts && Material"
            Height          =   1215
            Left            =   135
            TabIndex        =   66
            Top             =   2280
            Width           =   5265
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   39
               Left            =   1215
               TabIndex        =   67
               Top             =   240
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
               TabIndex        =   31
               Top             =   555
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
               Left            =   3525
               TabIndex        =   68
               Top             =   840
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
               TabIndex        =   32
               Top             =   870
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
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   255
               Left            =   2430
               TabIndex        =   72
               Top             =   885
               Width           =   945
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Part"
               Height          =   255
               Left            =   120
               TabIndex        =   71
               Top             =   600
               Width           =   945
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Labour"
               Height          =   255
               Left            =   120
               TabIndex        =   70
               Top             =   285
               Width           =   945
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Material"
               Height          =   255
               Left            =   120
               TabIndex        =   69
               Top             =   915
               Width           =   945
            End
         End
         Begin VB.ComboBox cboStatus 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            ItemData        =   "frmLihatOrder.frx":0000
            Left            =   2265
            List            =   "frmLihatOrder.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   165
            Width           =   735
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   26
            Left            =   1260
            TabIndex        =   73
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Index           =   28
            Left            =   1260
            TabIndex        =   74
            Top             =   480
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
            Index           =   29
            Left            =   2490
            TabIndex        =   75
            Top             =   480
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
            Index           =   42
            Left            =   1260
            TabIndex        =   76
            Top             =   780
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
            Left            =   2340
            TabIndex        =   77
            Top             =   795
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
            Left            =   2820
            TabIndex        =   78
            Top             =   780
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
            Left            =   4380
            TabIndex        =   23
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Left            =   4380
            TabIndex        =   24
            Top             =   480
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
            Left            =   4980
            TabIndex        =   92
            Top             =   495
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   240
            Index           =   20
            Left            =   1260
            TabIndex        =   26
            Top             =   1095
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   423
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
            Height          =   240
            Index           =   27
            Left            =   1260
            TabIndex        =   27
            Top             =   1320
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   423
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
            Height          =   240
            Index           =   46
            Left            =   1260
            TabIndex        =   28
            Top             =   1545
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   423
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
            Height          =   240
            Index           =   47
            Left            =   1260
            TabIndex        =   29
            Top             =   1770
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   423
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
            Height          =   240
            Index           =   48
            Left            =   1260
            TabIndex        =   30
            Top             =   1995
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   423
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
            Index           =   49
            Left            =   4380
            TabIndex        =   25
            Top             =   780
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
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "T SA."
            Height          =   255
            Left            =   3300
            TabIndex        =   93
            Top             =   825
            Width           =   945
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Keluhan && Permintaan Customer"
            Height          =   630
            Left            =   165
            TabIndex        =   88
            Top             =   1140
            Width           =   945
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inv. No."
            Height          =   255
            Left            =   180
            TabIndex        =   83
            Top             =   825
            Width           =   945
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. WO"
            Height          =   255
            Left            =   180
            TabIndex        =   82
            Top             =   225
            Width           =   945
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            Height          =   255
            Left            =   180
            TabIndex        =   81
            Top             =   525
            Width           =   945
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "K. Meter"
            Height          =   255
            Left            =   3285
            TabIndex        =   80
            Top             =   225
            Width           =   945
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "MCH"
            Height          =   255
            Left            =   3285
            TabIndex        =   79
            Top             =   525
            Width           =   945
         End
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2970
      Left            =   405
      TabIndex        =   41
      Top             =   330
      Width           =   11400
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   945
         TabIndex        =   1
         Top             =   210
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
         Left            =   945
         TabIndex        =   2
         Top             =   510
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
         Left            =   945
         TabIndex        =   3
         Top             =   810
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
         Left            =   945
         TabIndex        =   4
         Top             =   1110
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
         Index           =   5
         Left            =   945
         TabIndex        =   5
         Top             =   1410
         Width           =   1815
         _ExtentX        =   3201
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
         Index           =   6
         Left            =   945
         TabIndex        =   6
         Top             =   1710
         Width           =   1815
         _ExtentX        =   3201
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
         Index           =   7
         Left            =   945
         TabIndex        =   7
         Top             =   2010
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
         Index           =   8
         Left            =   945
         TabIndex        =   8
         Top             =   2310
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
         Height          =   270
         Index           =   9
         Left            =   930
         TabIndex        =   9
         Top             =   2610
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   476
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   55
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   5025
         TabIndex        =   10
         Top             =   810
         Width           =   780
         _ExtentX        =   1376
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
         Index           =   11
         Left            =   6450
         TabIndex        =   11
         Top             =   810
         Width           =   780
         _ExtentX        =   1376
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
         Index           =   12
         Left            =   5025
         TabIndex        =   12
         Top             =   1110
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   13
         Left            =   5025
         TabIndex        =   13
         Top             =   1410
         Width           =   1230
         _ExtentX        =   2170
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
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   5025
         TabIndex        =   14
         Top             =   2010
         Width           =   2235
         _ExtentX        =   3942
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
         Index           =   15
         Left            =   5025
         TabIndex        =   15
         Top             =   2310
         Width           =   2235
         _ExtentX        =   3942
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
         Index           =   16
         Left            =   8550
         TabIndex        =   37
         Top             =   195
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
         Left            =   8550
         TabIndex        =   16
         Top             =   495
         Width           =   1725
         _ExtentX        =   3043
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
         Left            =   8550
         TabIndex        =   17
         Top             =   1095
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
         Index           =   19
         Left            =   8550
         TabIndex        =   39
         Top             =   795
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
         Index           =   21
         Left            =   8550
         TabIndex        =   19
         Top             =   1695
         Width           =   1710
         _ExtentX        =   3016
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
         Index           =   22
         Left            =   8550
         TabIndex        =   20
         Top             =   1995
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
         Index           =   23
         Left            =   8550
         TabIndex        =   21
         Top             =   2310
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
         Index           =   24
         Left            =   10440
         TabIndex        =   38
         Top             =   495
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
         Index           =   25
         Left            =   8550
         TabIndex        =   18
         Top             =   1395
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   10320
         TabIndex        =   89
         Top             =   510
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
         Index           =   53
         Left            =   2910
         TabIndex        =   94
         Top             =   1710
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   2790
         TabIndex        =   95
         Top             =   1725
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label47 
         Caption         =   "(HH:MM)"
         Height          =   240
         Left            =   9300
         TabIndex        =   96
         Top             =   2340
         Width           =   660
      End
      Begin VB.Label Label30 
         Caption         =   "(DD-MM-YYYY)"
         Height          =   240
         Left            =   2190
         TabIndex        =   91
         Top             =   2340
         Width           =   1140
      End
      Begin VB.Label Label20 
         Caption         =   "(DD-MM-YYYY)"
         Height          =   240
         Left            =   2190
         TabIndex        =   90
         Top             =   2070
         Width           =   1140
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P. Jam Selesai"
         Height          =   255
         Left            =   7245
         TabIndex        =   63
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Key No."
         Height          =   255
         Left            =   7455
         TabIndex        =   62
         Top             =   2055
         Width           =   945
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E/G No."
         Height          =   255
         Left            =   7455
         TabIndex        =   61
         Top             =   1740
         Width           =   945
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
         Height          =   255
         Left            =   7455
         TabIndex        =   60
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         Height          =   255
         Left            =   7455
         TabIndex        =   59
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Left            =   7455
         TabIndex        =   58
         Top             =   540
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         Height          =   255
         Left            =   7455
         TabIndex        =   57
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Job Disc."
         Height          =   255
         Left            =   4005
         TabIndex        =   56
         Top             =   2355
         Width           =   945
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pemakai"
         Height          =   255
         Left            =   3975
         TabIndex        =   55
         Top             =   2055
         Width           =   945
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DDTE."
         Height          =   255
         Left            =   4395
         TabIndex        =   54
         Top             =   1455
         Width           =   555
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   255
         Left            =   4485
         TabIndex        =   53
         Top             =   1155
         Width           =   420
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disc."
         Height          =   255
         Left            =   6000
         TabIndex        =   52
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         Height          =   255
         Left            =   4545
         TabIndex        =   51
         Top             =   855
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         Height          =   255
         Left            =   -165
         TabIndex        =   50
         Top             =   2655
         Width           =   945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. RC."
         Height          =   255
         Left            =   -150
         TabIndex        =   49
         Top             =   2355
         Width           =   945
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Rep."
         Height          =   255
         Left            =   -150
         TabIndex        =   48
         Top             =   2055
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
         Height          =   255
         Left            =   -150
         TabIndex        =   47
         Top             =   1755
         Width           =   945
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
         Height          =   255
         Left            =   -150
         TabIndex        =   46
         Top             =   1455
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telp. No."
         Height          =   255
         Left            =   -150
         TabIndex        =   45
         Top             =   1155
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   255
         Left            =   -150
         TabIndex        =   44
         Top             =   855
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   255
         Left            =   -150
         TabIndex        =   43
         Top             =   555
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   -150
         TabIndex        =   42
         Top             =   255
         Width           =   945
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Top             =   45
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
      FormatString    =   "Pembatalan Order Pekerjaan"
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
      Left            =   285
      TabIndex        =   40
      Top             =   90
      Width           =   945
   End
End
Attribute VB_Name = "frmLihatOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Baris As Integer
Dim RateCPI As Double
Const MaxBrs = 7

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

Private Sub HapusForm()
    Dim i, j As Integer
    
    For i = 1 To 49
        Select Case i
            Case 32, 33, 34, 35, 36
            Case 10
                Text(10) = "0,1"
            Case 49
                Text(i) = gSA
            Case 7, 8, 13, 28
                Text(i) = ">>->>->>>>"
            Case 23, 29
                Text(i) = ">>:>>"
            Case Else
                Text(i) = ""
        End Select
    Next
    Text(53) = ""
'    cboStatus.Text = ""
    
'    sSQL = "Select count(*) as Jml_Data from CUSTDAT"
'    oRs1.Open sSQL, oConn, adOpenStatic
'    If oRs1.BOF Then
'        Text(36) = 0
'    Else
'        If IsNull(oRs1!Jml_Data) Then
'            Text(36) = 0
'        Else
'            Text(36) = oRs1!Jml_Data
'        End If
'    End If
'    oRs1.Close
    Baris = 1
    For i = 1 To 7
        DGrid.TextMatrix(i, 1) = ""
        DGrid.TextMatrix(i, 2) = ""
        DGrid.TextMatrix(i, 3) = ""
        DGrid.TextMatrix(i, 4) = ""
    Next
End Sub

Private Sub DGrid_DblClick()
    If Baris > 1 And DGrid.TextMatrix(DGrid.Row, 1) <> "" Then
        A = DGrid.Row
        'AMBIL
        With DGrid
            Text(33) = .TextMatrix(A, 1)
            Text(32) = .TextMatrix(A, 2)
            Text(34) = .TextMatrix(A, 3)
            Text(35) = .TextMatrix(A, 4)
        End With
        'HAPUS
        For i = A To Baris
            With DGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
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
                .TextMatrix(Baris, 3) = Format(Koma2Titik(Text(34)), "#,##0.00")
                .TextMatrix(Baris, 4) = Text(35)
            End With
            Baris = Baris + 1
            Text(33) = ""
            Text(32) = ""
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

Private Function TotalJasa() As Double
    Dim Temp As Double
    
    Temp = 0
    With DGrid
        For i = 1 To 7
            If .TextMatrix(i, 1) = "" Then
                Exit For
            Else
                Temp = Temp + Koma2Titik(.TextMatrix(i, 4))
            End If
        Next
    End With
    TotalJasa = Temp
End Function

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call DGrid_DblClick
    End If
End Sub

Private Sub Form_Load()
'    Call CenterForm(Me, True)
    Baris = 1
    DGrid.Rows = 11
    DGrid.Cols = 5
    DGrid.ColWidth(0) = 0
    DGrid.ColWidth(1) = 2600
    DGrid.ColWidth(2) = 800
    DGrid.ColWidth(3) = 1050
    DGrid.ColWidth(4) = 1095
    DGrid.ColAlignment(1) = flexAlignLeftCenter
    DGrid.ColAlignment(2) = flexAlignLeftCenter
    DGrid.ColAlignment(3) = flexAlignRightCenter
    DGrid.ColAlignment(4) = flexAlignLeftCenter
    DGrid.TextMatrix(0, 1) = "Keterangan"
    DGrid.TextMatrix(0, 2) = "JP"
    DGrid.TextMatrix(0, 3) = "Rate"
    DGrid.TextMatrix(0, 4) = "Kondisi"
    Call HapusForm
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub SaveDataCust()
    sSQL = "Select * from CUSTDAT where NO_WO ='" & Text(0) & "'"
    oRs.Open sSQL, oConn, adOpenStatic
    
    Call ClearsSql
    gsSql1(0) = "CUSTOMER"
    gsSql2(0) = Text(1)
    gsSql1(1) = "CUST_ADRES"
    gsSql2(1) = Text(2)
    gsSql1(2) = "KOTA"
    gsSql2(2) = Text(3)
    gsSql1(3) = "TELPON"
    gsSql2(3) = Text(4)
    gsSql1(4) = "NO_POLISI"
    gsSql2(4) = Text(16)
    gsSql1(5) = "NO_ENGINE"
    gsSql2(5) = Text(21)
'    gsSql1(6) = "VDS"
'    gsSql2(6) = Left(Text(25), 6)
    gsSql1(6) = "SPK"
    gsSql2(6) = Text(23)
    gsSql1(7) = "NO_CHASSIS"
    gsSql2(7) = Text(25)
    gsSql1(8) = "MDL"
    gsSql2(8) = Text(17)
    gsSql1(9) = "MODEL"
    gsSql2(9) = Text(24)
    gsSql1(10) = "COLOR"
    gsSql2(10) = Text(18)
    gsSql1(11) = "FS_KD_DEALER"
    gsSql2(11) = Text(6)
    gsSql1(12) = "DDATE"
    gsSql2(12) = Text(13)
    gsSql1(13) = "REMARK"
    gsSql2(13) = Text(9)
    gsSql1(14) = "NPWP"
    gsSql2(14) = Text(5)
    gsSql1(15) = "PPN"
    gsSql2(15) = Val(Text(10))
    gsSql1(16) = "DISC"
    gsSql2(16) = Val(Text(11))
    gsSql1(17) = "KEYNO"
    gsSql2(17) = Text(22)
    gsSql1(18) = "TGL_REP"
    gsSql2(18) = TglYYYYMMDD(Text(7))
    gsSql1(19) = "CODE"
    gsSql2(19) = Text(12)
    gsSql1(20) = "[USER]"
    gsSql2(20) = Text(14)
    gsSql1(21) = "JOBDSC"
    gsSql2(21) = Text(15)
    gsSql1(22) = "RCDATE"
    gsSql2(22) = TglYYYYMMDD(Text(8))
    gsSql1(23) = "FMDL"
    gsSql2(23) = Text(19)
    
    If Not oRs.BOF Then
        sSQL = GensSql("CUSTDAT", True) + " where NO_WO ='" & Text(0) & "'"
    Else
        sSQL = GensSql("CUSTDAT", False)
    End If
    oConn.Execute sSQL
    oRs.Close
End Sub

Private Sub SaveOrderPekerjaan()
    Dim Booking As String
    Dim NomorWO As String
    
    sSQL = "Select * from T_PROFIL"
    oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        NomorWO = Format(oRs("FN_NO_WO"), "00000000")
    End If
    oRs.Close
    
    sSQL = "Select * from BOOK where NO_WO ='" & Text(0) & "' and DELIV='BOK'"
    oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        Booking = "BOOKED"
    Else
        Booking = ""
    End If
    oRs.Close
    
    sSQL = "Select * from JOBDTA where NO_WO ='" & Text(0) & "' and DELIV<>'BTL'"
    oRs.Open sSQL, oConn, adOpenStatic
    Call ClearsSql
    gsSql1(0) = "CUSTOMER"
    gsSql2(0) = Text(1)
    gsSql1(1) = "CUST_ADRES"
    gsSql2(1) = Text(2)
    gsSql1(2) = "KOTA"
    gsSql2(2) = Text(3)
    gsSql1(3) = "TELPON"
    gsSql2(3) = Text(4)
    gsSql1(4) = "NO_POLISI"
    gsSql2(4) = Text(16)
    gsSql1(5) = "NO_ENGINE"
    gsSql2(5) = Text(21)
    gsSql1(6) = "NO_CHASSIS"
    gsSql2(6) = Text(25)
    gsSql1(7) = "MDL"
    gsSql2(7) = Text(17)
    gsSql1(8) = "MODEL"
    gsSql2(8) = Text(24)
    gsSql1(9) = "COLOR"
    gsSql2(9) = Text(18)
    gsSql1(10) = "FS_KD_DEALER"
    gsSql2(10) = Text(6)
    gsSql1(11) = "DDATE"
    gsSql2(11) = Text(13)
    gsSql1(12) = "REMARK1"
    gsSql2(12) = Text(37)
    gsSql1(13) = "NPWP"
    gsSql2(13) = Text(5)
    gsSql1(14) = "PPN"
    gsSql2(14) = SimpanKoma2Titik(Text(10))
    gsSql1(15) = "DISC"
    gsSql2(15) = SimpanKoma2Titik(Text(11))
    gsSql1(16) = "CODE"
    gsSql2(16) = Text(12)
    gsSql1(17) = "FMDL"
    gsSql2(17) = Text(19)
    gsSql1(18) = "SPK"
    gsSql2(18) = Text(23)
    gsSql1(19) = "COND7"
    gsSql2(19) = DGrid.TextMatrix(7, 4)
    gsSql1(20) = "REQ1"
    gsSql2(20) = Text(20)
    gsSql1(21) = "REQ2"
    gsSql2(21) = Text(27)
    gsSql1(22) = "REQ3"
    gsSql2(22) = Text(46)
    gsSql1(23) = "REQ4"
    gsSql2(23) = Text(47)
    gsSql1(24) = "REQ5"
    gsSql2(24) = Text(48)
    gsSql1(25) = "MECH"
    gsSql2(25) = Text(31)
    gsSql1(26) = "KM"
    gsSql2(26) = Text(30)
    gsSql1(27) = "SUGGEST"
    gsSql2(27) = Text(38)
    gsSql1(28) = "LABOR"
    gsSql2(28) = SimpanKoma2Titik(Text(39))
    gsSql1(29) = "PART"
    gsSql2(29) = SimpanKoma2Titik(Text(40))
    gsSql1(30) = "MTRL"
    gsSql2(30) = SimpanKoma2Titik(Text(45))
    gsSql1(31) = "TOTAL"
    gsSql2(31) = SimpanKoma2Titik(Text(41))
'-----------Konstanta
    gsSql1(32) = "PAID"
    gsSql2(32) = "C"
    gsSql1(33) = "DELIV"
    gsSql2(33) = "PRC"
'-----------Konstanta
    gsSql1(34) = "STATUS"
    gsSql2(34) = cboStatus.Text
    gsSql1(35) = "JOB"
    gsSql2(35) = Text(32)
    gsSql1(36) = "NOFAK"
    gsSql2(36) = Booking
    
    gsSql1(37) = "REP1"
    gsSql2(37) = DGrid.TextMatrix(1, 1)
    gsSql1(38) = "REP2"
    gsSql2(38) = DGrid.TextMatrix(2, 1)
    gsSql1(39) = "REP3"
    gsSql2(39) = DGrid.TextMatrix(3, 1)
    gsSql1(40) = "REP4"
    gsSql2(40) = DGrid.TextMatrix(4, 1)
    gsSql1(41) = "REP5"
    gsSql2(41) = DGrid.TextMatrix(5, 1)
    gsSql1(42) = "REP6"
    gsSql2(42) = DGrid.TextMatrix(6, 1)
    gsSql1(43) = "REP7"
    gsSql2(43) = DGrid.TextMatrix(7, 1)
    gsSql1(44) = "RTE1"
    gsSql2(44) = SimpanKoma2Titik(DGrid.TextMatrix(1, 3))
    gsSql1(45) = "RTE2"
    gsSql2(45) = SimpanKoma2Titik(DGrid.TextMatrix(2, 3))
    gsSql1(46) = "RTE3"
    gsSql2(46) = SimpanKoma2Titik(DGrid.TextMatrix(3, 3))
    gsSql1(47) = "RTE4"
    gsSql2(47) = SimpanKoma2Titik(DGrid.TextMatrix(4, 3))
    gsSql1(48) = "RTE5"
    gsSql2(48) = SimpanKoma2Titik(DGrid.TextMatrix(5, 3))
    gsSql1(49) = "RTE6"
    gsSql2(49) = SimpanKoma2Titik(DGrid.TextMatrix(6, 3))
    gsSql1(50) = "RTE7"
    gsSql2(50) = SimpanKoma2Titik(DGrid.TextMatrix(7, 3))
    
    gsSql1(51) = "KLPKERJA1"
    gsSql2(51) = DGrid.TextMatrix(1, 2)
    gsSql1(52) = "KLPKERJA2"
    gsSql2(52) = DGrid.TextMatrix(2, 2)
    gsSql1(53) = "KLPKERJA3"
    gsSql2(53) = DGrid.TextMatrix(3, 2)
    gsSql1(54) = "KLPKERJA4"
    gsSql2(54) = DGrid.TextMatrix(4, 2)
    gsSql1(55) = "KLPKERJA5"
    gsSql2(55) = DGrid.TextMatrix(5, 2)
    gsSql1(56) = "KLPKERJA6"
    gsSql2(56) = DGrid.TextMatrix(6, 2)
    gsSql1(57) = "KLPKERJA7"
    gsSql2(57) = DGrid.TextMatrix(7, 2)
    gsSql1(58) = "COND1"
    gsSql2(58) = DGrid.TextMatrix(1, 4)
    gsSql1(59) = "COND2"
    gsSql2(59) = DGrid.TextMatrix(2, 4)
    gsSql1(60) = "COND3"
    gsSql2(60) = DGrid.TextMatrix(3, 4)
    gsSql1(61) = "COND4"
    gsSql2(61) = DGrid.TextMatrix(4, 4)
    gsSql1(62) = "COND5"
    gsSql2(62) = DGrid.TextMatrix(5, 4)
    gsSql1(63) = "COND6"
    gsSql2(63) = DGrid.TextMatrix(6, 4)
    gsSql1(64) = "FRONT"
    gsSql2(64) = Text(49)
    
    If Not oRs.BOF Then
        sSQL = GensSql("JOBDTA", True) + " where NO_WO ='" & Text(0) & "' and DELIV<>'BTL'"
    Else
        Text(26) = NomorWO
        gsSql1(64) = "NO_WO"
        gsSql2(64) = Text(26)
        gsSql1(65) = "DATE"
        gsSql2(65) = TglYYYYMMDD(Text(28))
        gsSql1(66) = "TIME"
        gsSql2(66) = Text(29)
        sSQL = GensSql("JOBDTA", False)
    End If
    oConn.Execute sSQL
    
    If oRs.BOF Then
        sSQL = "Select * from T_PROFIL"
        oRs1.Open sSQL, oConn, adOpenStatic
        If Not oRs1.BOF Then
            sSQL = "update T_PROFIL set FN_NO_WO = " & oRs1("FN_NO_WO") + 1
            oConn.Execute sSQL
        End If
        oRs1.Close
        MsgBox "No WO : " & Text(26)
    End If
    oRs.Close
End Sub

Private Sub TblSimpan_Click()
    Dim Boleh As Boolean
    Dim Pesan As String
    Boleh = True
    
    If CekTgl(Text(7)) = False Then
        Pesan = Pesan + "Tgl Rep. "
        Boleh = False
    End If
    If CekTgl(Text(8)) = False Then
        Pesan = Pesan + "Tgl RC. "
        Boleh = False
    End If
    If Trim(cboStatus) = "" Then
        Pesan = Pesan + "Status "
        Boleh = False
    End If
    If Trim(Text(0)) = "" Then
        Pesan = Pesan + "No. Polisi "
        Boleh = False
    End If
    
    If Boleh = True Then
        Call SaveDataCust
        Text(26) = "NOWO"
        Text(28) = gTanggal
        Text(29) = Format(Time, "HH:NN")
        Call SaveOrderPekerjaan
        sSQL = "update BOOK set DELIV='OKE' where NO_WO ='" & Text(16) & "' and DELIV='BOK'"
        oConn.Execute sSQL
        Call HapusForm
        Text(0) = ""
    Else
        MsgBox "Masukkan " + Pesan + "!"
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

Private Sub PanggilDataCust(NoWO As String)
    sSQL = "Select * from JOBDTA where NO_WO ='" & NoWO & "'"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        Text(1) = oRs1("CUSTOMER")
        Text(2) = oRs1("CUST_ADRES")
        Text(3) = oRs1("KOTA")
        Text(4) = oRs1("TELPON")
        Text(16) = oRs1("NO_POLISI")
        Text(21) = oRs1("NO_ENGINE")
        Text(25) = oRs1("NO_CHASSIS")
        Text(17) = oRs1("MDL")
        Text(24) = oRs1("MODEL")
        Text(18) = oRs1("COLOR")
        Text(6) = oRs1("FS_KD_DEALER")
        Text(53) = oRs1("FS_NM_DEALER")
        Text(13) = oRs1("DDATE")
        Text(5) = oRs1("NPWP")
        Text(10) = oRs1("PPN")
        Text(11) = oRs1("DISC")
        Text(12) = oRs1("CODE")
        Text(19) = oRs1("FMDL")
        Text(23) = oRs1("SPK")
        Text(16) = oRs1("NO_POLISI")
    End If
    oRs1.Close
    sSQL = "Select * from CUSTDAT where NO_POLISI ='" & Text(16) & "'"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        Text(22) = oRs1("KEYNO")
        Text(7) = TglDDMMYYYY(oRs1("TGL_REP"))
        Text(9) = oRs1("REMARK")
        Text(14) = oRs1("USER")
        Text(15) = oRs1("JOBDSC")
        Text(8) = TglDDMMYYYY(oRs1("RCDATE"))
    End If
    oRs1.Close
End Sub

Private Sub PanggilDataOrder(NoWO As String)
    Dim i, j As Integer
    
    sSQL = "Select * from JOBDTA where NO_WO ='" & NoWO & "'"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        Text(1) = oRs1("CUSTOMER")
        Text(2) = oRs1("CUST_ADRES")
        Text(3) = oRs1("KOTA")
        Text(4) = oRs1("TELPON")
        Text(16) = oRs1("NO_POLISI")
        Text(21) = oRs1("NO_ENGINE")
        Text(25) = oRs1("NO_CHASSIS")
        Text(17) = oRs1("MDL")
        Text(24) = oRs1("MODEL")
        Text(18) = oRs1("COLOR")
        Text(6) = oRs1("FS_KD_DEALER")
        Text(13) = oRs1("DDATE")
        Text(37) = oRs1("REMARK1")
        Text(5) = oRs1("NPWP")
        Text(10) = oRs1("PPN")
        Text(11) = oRs1("DISC")
        Text(12) = oRs1("CODE")
        Text(19) = oRs1("FMDL")
        Text(23) = oRs1("SPK")
        Text(42) = oRs1("NOINV")
        Text(20) = oRs1("REQ1")
        Text(27) = oRs1("REQ2")
        Text(46) = oRs1("REQ3")
        Text(47) = oRs1("REQ4")
        Text(48) = oRs1("REQ5")
        Text(31) = oRs1("MECH")
        Text(30) = oRs1("KM")
        Text(28) = TglDDMMYYYY(oRs1("DATE"))
        Text(29) = Format(oRs1("TIME"), "HH:NN")
        Text(38) = oRs1("SUGGEST")
        Text(39) = oRs1("LABOR")
        Text(40) = oRs1("PART")
        Text(45) = oRs1("MTRL")
        Text(41) = oRs1("TOTAL")
        Text(26) = oRs1("NO_WO")
        Text(44) = oRs1("PAID")
        Text(43) = oRs1("DELIV")
        cboStatus.Text = oRs1("STATUS")
'        Text(32) = oRs1("JOB")
        Text(26) = oRs1("NO_WO")
        Text(49) = oRs1("FRONT")
        
        Baris = 1
        For i = 1 To 7
            DGrid.TextMatrix(i, 1) = oRs1("REP" & Trim(Str(i)))
            DGrid.TextMatrix(i, 2) = oRs1("KLPKERJA" & Trim(Str(i)))
            DGrid.TextMatrix(i, 3) = Format(oRs1("RTE" & Trim(Str(i))), "#,##0.00")
            DGrid.TextMatrix(i, 4) = oRs1("COND" & Trim(Str(i)))
        Next
        
        For i = 1 To 7
            If DGrid.TextMatrix(i, 1) <> "" Then
                Baris = Baris + 1
            End If
        Next
        
        If oRs1("NOINV") <> "" Then
            MsgBox "Telah dibuat Invoice dengan No : " & Trim(oRs1("NOINV")) & ", Tgl : " & Trim(TglDDMMYYYY(oRs1("INVDATE"))) & ", Sejumlah : " & Format(oRs1("Total"), "#,##0.00") & " !"
'            TblSimpan.Enabled = False
        Else
            Select Case Trim(oRs1("RPR_COND"))
                Case ""
                    MsgBox "Kendaraan dalam proses tunggu perbaikan !"
                Case "P"
                    MsgBox "Kendaraan dalam proses perbaikan !"
                Case "B"
                    MsgBox "Kendaraan dalam pending job perbaikan !"
                Case Else
                    MsgBox "Kendaraan telah selesai perbaikan !"
            End Select
        End If
    End If
    oRs1.Close
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "LIHATORDER"
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
        Case 10, 11, 34, 40, 45
            KeyAscii = CekNumerik(KeyAscii)
    End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    
    Dim NoWO As String
    
    Select Case Index
        Case 0
            If Text(0) <> "" Then
'                Text(16) = Text(0)
                NoWO = Text(0)
'                TblSimpan.Enabled = True
                sSQL = "Select * from JOBDTA where NO_WO ='" & NoWO & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Call HapusForm
                    MsgBox "No. WO tersebut tidak ada !"
                Else
                    Call PanggilDataCust(NoWO)
                    Call PanggilDataOrder(NoWO)
                End If
                oRs.Close
            
                gMDL = Text(17)
                Call HitungNilaiKaliRate
                Call HitungTotal
            Else
                Call HapusForm
            End If
        Case 7, 8
            If CekTgl(Text(Index)) = False Then
                Text(Index) = gTanggal
            End If
        Case 10, 11
            If Koma2Titik(Text(Index)) > 1 Then
                Text(Index) = 0
            End If
        Case 17
            If Trim(Text(Index)) <> "" Then
                sSQL = "Select * from MODEL where MDL ='" & Text(17) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCari.Show (vbModal)
                Else
                    Text(17) = oRs("MDL")
                    Text(19) = oRs("Desc")
                    Text(24) = oRs("MODEL")
                End If
                oRs.Close
                gMDL = Text(17)
                Call HitungNilaiKaliRate
            End If
        Case 31
            If Trim(Text(Index)) <> "" Then
                sSQL = "Select * from MEKANIK where MECH ='" & Text(Index) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MEKANIK"
                    frmCari.Show (vbModal)
                End If
                oRs.Close
            End If
        Case 23
            If CekJam(Text(Index)) = False Then
                Text(Index) = Format(Time, "HH:NN")
            End If
        Case 33
            If Trim(Text(Index)) <> "" Then
                sSQL = "Select * from T_JOBLIST where FS_ITEM='" & Text(33) & "' and FS_MDL='" & gMDL & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If Not oRs.BOF Then
                    Text(34) = oRs("FN_RATE")
'                Else
'                    Text(34) = ""
                End If
                oRs.Close
            End If
        Case 34
            If Trim(Text(Index)) = "" Then
                Text(Index) = 0
            End If
        Case 40, 45
            Call HitungTotal
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
            RateCPI = 0
    End Select
End Sub

