VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmCatatKredit 
   Caption         =   "Pencatatan Kredit"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton TblSimpan 
      Caption         =   "&Simpan"
      Height          =   390
      Left            =   8040
      TabIndex        =   63
      Top             =   7020
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9930
      TabIndex        =   64
      Top             =   7020
      Width           =   1890
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2670
      TabIndex        =   30
      Top             =   120
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   4980
      Left            =   405
      TabIndex        =   27
      Top             =   1980
      Width           =   11460
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   38
         Left            =   1575
         TabIndex        =   66
         Top             =   4080
         Width           =   5655
         _ExtentX        =   9975
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
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
         Height          =   345
         Left            =   1575
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   3705
         Width           =   9825
      End
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         ItemData        =   "frmCatatKredit.frx":0000
         Left            =   5280
         List            =   "frmCatatKredit.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Height          =   3150
         Left            =   7110
         TabIndex        =   37
         Top             =   525
         Width           =   4305
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   39
            Left            =   930
            TabIndex        =   38
            Top             =   270
            Width           =   1500
            _ExtentX        =   2646
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
            Left            =   930
            TabIndex        =   5
            Top             =   585
            Width           =   1500
            _ExtentX        =   2646
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
            Left            =   2355
            TabIndex        =   39
            Top             =   2700
            Width           =   1620
            _ExtentX        =   2858
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
            Left            =   930
            TabIndex        =   6
            Top             =   900
            Width           =   1500
            _ExtentX        =   2646
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
            Index           =   4
            Left            =   930
            TabIndex        =   7
            Top             =   1215
            Width           =   1500
            _ExtentX        =   2646
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
            Index           =   5
            Left            =   930
            TabIndex        =   8
            Top             =   1530
            Width           =   1500
            _ExtentX        =   2646
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
            Index           =   6
            Left            =   2340
            TabIndex        =   57
            Top             =   1935
            Width           =   1620
            _ExtentX        =   2858
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
            Index           =   7
            Left            =   930
            TabIndex        =   9
            Top             =   2295
            Width           =   1500
            _ExtentX        =   2646
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
         Begin VB.Label Label13 
            Caption         =   "+"
            Height          =   210
            Left            =   4020
            TabIndex        =   61
            Top             =   2535
            Width           =   180
         End
         Begin VB.Label Label12 
            Caption         =   "+"
            Height          =   210
            Left            =   3990
            TabIndex        =   60
            Top             =   1755
            Width           =   135
         End
         Begin VB.Line Line2 
            X1              =   930
            X2              =   3990
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line1 
            X1              =   930
            X2              =   3975
            Y1              =   1860
            Y2              =   1860
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   255
            Left            =   -165
            TabIndex        =   59
            Top             =   2340
            Width           =   945
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
            Height          =   255
            Left            =   1230
            TabIndex        =   58
            Top             =   1980
            Width           =   945
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SubBlet"
            Height          =   255
            Left            =   -165
            TabIndex        =   56
            Top             =   1575
            Width           =   945
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bahan"
            Height          =   255
            Left            =   -165
            TabIndex        =   55
            Top             =   1260
            Width           =   945
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Labour"
            Height          =   255
            Left            =   -165
            TabIndex        =   43
            Top             =   315
            Width           =   945
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Part"
            Height          =   255
            Left            =   -165
            TabIndex        =   42
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   255
            Left            =   1215
            TabIndex        =   41
            Top             =   2730
            Width           =   945
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Oli"
            Height          =   255
            Left            =   -165
            TabIndex        =   40
            Top             =   945
            Width           =   945
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3150
         Left            =   45
         TabIndex        =   29
         Top             =   525
         Width           =   7050
         Begin MSFlexGridLib.MSFlexGrid DGrid 
            Height          =   2025
            Left            =   75
            TabIndex        =   4
            Top             =   930
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   3572
            _Version        =   393216
            FixedCols       =   0
            ScrollBars      =   0
         End
      End
      Begin VB.Frame Frame3 
         Height          =   630
         Left            =   75
         TabIndex        =   28
         Top             =   4290
         Width           =   11325
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   37
            Left            =   2010
            TabIndex        =   10
            Top             =   225
            Width           =   3390
            _ExtentX        =   5980
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
            Left            =   10275
            TabIndex        =   11
            Top             =   225
            Width           =   570
            _ExtentX        =   1005
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
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRC"
            Height          =   255
            Left            =   9180
            TabIndex        =   62
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Slip/Keterangan"
            Height          =   255
            Left            =   75
            TabIndex        =   54
            Top             =   285
            Width           =   1845
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   31
         Left            =   6930
         TabIndex        =   2
         Top             =   240
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
         Left            =   7530
         TabIndex        =   44
         Top             =   255
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
         Index           =   26
         Left            =   3360
         TabIndex        =   47
         Top             =   255
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
         Left            =   1080
         TabIndex        =   48
         Top             =   255
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
         Index           =   42
         Left            =   10140
         TabIndex        =   51
         Top             =   240
         Width           =   1050
         _ExtentX        =   1852
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
         Mask            =   "CCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   13
         Left            =   8520
         TabIndex        =   3
         Top             =   240
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
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         Height          =   255
         Left            =   480
         TabIndex        =   67
         Top             =   4125
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "Terbilang (Rupiah)"
         Height          =   270
         Left            =   135
         TabIndex        =   68
         Top             =   3765
         Width           =   1395
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. No."
         Height          =   255
         Left            =   9060
         TabIndex        =   53
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T SA."
         Height          =   255
         Left            =   7440
         TabIndex        =   52
         Top             =   285
         Width           =   945
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Left            =   0
         TabIndex        =   50
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. WO"
         Height          =   255
         Left            =   2280
         TabIndex        =   49
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MCH"
         Height          =   255
         Left            =   5835
         TabIndex        =   46
         Top             =   285
         Width           =   945
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Left            =   4200
         TabIndex        =   45
         Top             =   315
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1560
      Left            =   405
      TabIndex        =   20
      Top             =   420
      Width           =   11460
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   825
         TabIndex        =   12
         Top             =   225
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
         Left            =   825
         TabIndex        =   13
         Top             =   540
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
         Left            =   825
         TabIndex        =   14
         Top             =   855
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
         Index           =   16
         Left            =   6165
         TabIndex        =   15
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
         Left            =   6165
         TabIndex        =   16
         Top             =   510
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
         Index           =   25
         Left            =   6165
         TabIndex        =   18
         Top             =   825
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
         Left            =   10755
         TabIndex        =   17
         Top             =   510
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
         Left            =   7890
         TabIndex        =   31
         Top             =   525
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
         Left            =   7980
         TabIndex        =   32
         Top             =   510
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
         Index           =   14
         Left            =   825
         TabIndex        =   33
         Top             =   1170
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
         Index           =   21
         Left            =   6165
         TabIndex        =   35
         Top             =   1140
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E/G No."
         Height          =   255
         Left            =   5070
         TabIndex        =   36
         Top             =   1185
         Width           =   945
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
         Height          =   255
         Left            =   -255
         TabIndex        =   34
         Top             =   1215
         Width           =   945
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
         Height          =   255
         Left            =   5070
         TabIndex        =   26
         Top             =   870
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Left            =   5070
         TabIndex        =   25
         Top             =   540
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         Height          =   255
         Left            =   5070
         TabIndex        =   24
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   255
         Left            =   -270
         TabIndex        =   23
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   255
         Left            =   -270
         TabIndex        =   22
         Top             =   585
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   -270
         TabIndex        =   21
         Top             =   270
         Width           =   945
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   105
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
      FormatString    =   "Pencatatan Kredit"
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
      Caption         =   "No. Invoice"
      Height          =   255
      Left            =   555
      TabIndex        =   19
      Top             =   150
      Width           =   945
   End
End
Attribute VB_Name = "frmCatatKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Baris As Integer
Dim RateCPI As Double
Dim NoInvoice As String
Const MaxBrs = 7

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
    For i = 1 To 45
        Select Case i
            Case 13
                Text(i) = gSA
            Case 9, 10, 11, 12, 15, 18, 20, 22, 23, 27, 29, 30, 32, 33, 34, 35, 36, 43, 44
            Case 28
                Text(i) = ">>->>->>>>"
            Case Else
                Text(i) = ""
        End Select
    Next
'    cboStatus.Text = ""
    Text1.Text = ""
    
    Baris = 1
    For i = 1 To 7
        DGrid.TextMatrix(i, 1) = ""
        DGrid.TextMatrix(i, 2) = ""
        DGrid.TextMatrix(i, 3) = ""
        DGrid.TextMatrix(i, 4) = ""
        DGrid.TextMatrix(i, 5) = ""
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
            Text(35) = .TextMatrix(A, 5)
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
                .TextMatrix(Baris, 4) = Format(Koma2Titik(Text(34)) * RateCPI, "#,##0.00")
'                .TextMatrix(Baris, 3) = Koma2Titik(Text(34))
'                .TextMatrix(Baris, 4) = Koma2Titik(Text(34)) * RateCPI
                .TextMatrix(Baris, 5) = Text(35)
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

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call DGrid_DblClick
    End If
End Sub

Private Sub Form_Load()
    Baris = 1
    DGrid.Rows = 11
    DGrid.Cols = 6
    DGrid.ColWidth(0) = 0
    DGrid.ColWidth(1) = 2670
    DGrid.ColWidth(2) = 800
    DGrid.ColWidth(3) = 1050
    DGrid.ColWidth(4) = 1700
    DGrid.ColWidth(5) = 600
    DGrid.ColAlignment(1) = flexAlignLeftCenter
    DGrid.ColAlignment(2) = flexAlignLeftCenter
    DGrid.ColAlignment(3) = flexAlignRightCenter
    DGrid.ColAlignment(4) = flexAlignRightCenter
    DGrid.ColAlignment(5) = flexAlignLeftCenter
    DGrid.TextMatrix(0, 1) = "Keterangan"
    DGrid.TextMatrix(0, 2) = "JP"
    DGrid.TextMatrix(0, 3) = "Rate"
    DGrid.TextMatrix(0, 4) = "Jumlah"
    DGrid.TextMatrix(0, 5) = "MECH"
    Call HapusForm
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub SaveOrderPekerjaan(NomorInv As String)
    sSQL = "Select * from T_INVOICE where NOINV ='" & NomorInv & "' and FS_NO_RETUR = ''"
    oRs.Open sSQL, oConn, adOpenStatic
    
    Call ClearsSql
    gsSql1(0) = "FS_KREDIT"
    gsSql2(0) = "1"
        
    If Not oRs.BOF Then
        sSQL = GensSql("T_INVOICE", True) + " where NOINV ='" & NomorInv & "' and FS_NO_RETUR = ''"
    End If
    oConn.Execute sSQL
    oRs.Close
End Sub

Private Sub SimpanData()
    Call SaveOrderPekerjaan(Text(0))
    Call HapusForm
End Sub

Private Sub TblSimpan_Click()
    Call SimpanData
    Text(0) = ""
    Text(0).SetFocus
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

Private Sub PanggilDataOrder(NomorIv As String)
    sSQL = "Select * from T_INVOICE where NOINV ='" & NomorIv & "' and FS_NO_RETUR = ''"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        'Data Customer & Order
        Text(1) = oRs1("CUSTOMER")
        Text(2) = oRs1("CUST_ADRES")
        Text(3) = oRs1("KOTA")
        Text(4) = oRs1("MTRL")
        Text(5) = oRs1("SBLT")
        Text(7) = oRs1("MET")
        Text(8) = oRs1("DELIV")
        Text(14) = oRs1("NPWP")
        Text(13) = oRs1("FRONT")
        Text(16) = oRs1("NO_POLISI")
        Text(17) = oRs1("MDL")
        Text(19) = oRs1("FMDL")
        Text(21) = oRs1("NO_ENGINE")
        Text(24) = oRs1("MODEL")
        
        Text(25) = oRs1("NO_CHASSIS")
        Text(26) = oRs1("NO_WO")
        Text(28) = TglDDMMYYYY(oRs1("DATE"))
        
        Text(31) = oRs1("MECH")
        Text(37) = oRs1("SLIP")
        Text(38) = oRs1("FS_CATATAN")
        Text(39) = oRs1("LABOR")
        Text(40) = oRs1("PART")
        Text(41) = oRs1("TOTAL")
        Text(42) = oRs1("NOINV")
        Text(45) = oRs1("OIL")
        cboStatus.Text = oRs1("STATUS")
        Call HitungNilaiKaliRate
        
        Baris = 1
        For i = 1 To 7
            DGrid.TextMatrix(i, 1) = oRs1("REP" & Trim(Str(i)))
            DGrid.TextMatrix(i, 2) = oRs1("KLPKERJA" & Trim(Str(i)))
            DGrid.TextMatrix(i, 3) = Format(oRs1("RTE" & Trim(Str(i))), "#,##0.00")
            DGrid.TextMatrix(i, 4) = Format(oRs1("OFF" & Trim(Str(i))), "#,##0.00")
            DGrid.TextMatrix(i, 5) = oRs1("MEKANIK" & Trim(Str(i)))
        Next
        
        For i = 1 To 7
            If DGrid.TextMatrix(i, 1) <> "" Then
                Baris = Baris + 1
            End If
        Next

    End If
    oRs1.Close
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gTglInv = gTanggal
                gSearchOption = "INVCATAT"
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
        Case 34
            KeyAscii = CekNumerik(KeyAscii)
    End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Dim NoInvoice As String
    
    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                NoInvoice = Text(0)
                sSQL = "Select * from T_INVOICE where NOINV ='" & NoInvoice & "' and FS_NO_RETUR = '' and FS_KREDIT=''"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gTglInv = gTanggal
                    gSearchOption = "INVCATAT"
                    frmCari.Show (vbModal)
                Else
                    Call PanggilDataOrder(NoInvoice)
                End If
                gMDL = Text(17)
                Text(39) = TotalJasa
                Call HitungTotal
                
                oRs.Close
            Else
                Call HapusForm
            End If
        Case 31, 35
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
        Case 34
            If Trim(Text(Index)) = "" Then
                Text(Index) = 0
            End If
        Case 40, 45, 4, 5, 7
            Call HitungTotal
    End Select
End Sub

Private Sub HitungTotal()
    Text(41) = Koma2Titik(Text(39)) + Koma2Titik(Text(40)) _
             + Koma2Titik(Text(45)) + Koma2Titik(Text(4)) _
             + Koma2Titik(Text(5)) + Koma2Titik(Text(7))
    Text(6) = Koma2Titik(Text(39)) + Koma2Titik(Text(40)) _
             + Koma2Titik(Text(45)) + Koma2Titik(Text(4)) _
             + Koma2Titik(Text(5))
    Text1.Text = Angka2Huruf(Text(41), 1)
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

