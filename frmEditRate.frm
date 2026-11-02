VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmEditRate 
   Caption         =   "Perubahan A.Rate, Group & Mekanik"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2265
      TabIndex        =   38
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
      Height          =   5520
      Left            =   405
      TabIndex        =   35
      Top             =   1980
      Width           =   11460
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   38
         Left            =   1560
         TabIndex        =   82
         Top             =   4110
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
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   81
         Top             =   3735
         Width           =   9825
      End
      Begin VB.CommandButton TblSimpan 
         Caption         =   "&Simpan"
         Height          =   390
         Left            =   7620
         TabIndex        =   11
         Top             =   5055
         Width           =   1890
      End
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmEditRate.frx":0000
         Left            =   5280
         List            =   "frmEditRate.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Caption         =   "Biaya"
         Height          =   3150
         Left            =   7950
         TabIndex        =   51
         Top             =   540
         Width           =   3465
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   39
            Left            =   930
            TabIndex        =   52
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
            TabIndex        =   12
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
            Left            =   1515
            TabIndex        =   53
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
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            Left            =   1500
            TabIndex        =   71
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
            TabIndex        =   16
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
            Left            =   3150
            TabIndex        =   75
            Top             =   2535
            Width           =   180
         End
         Begin VB.Label Label12 
            Caption         =   "+"
            Height          =   210
            Left            =   3135
            TabIndex        =   74
            Top             =   1755
            Width           =   135
         End
         Begin VB.Line Line2 
            X1              =   915
            X2              =   3135
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line1 
            X1              =   930
            X2              =   3105
            Y1              =   1875
            Y2              =   1875
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   255
            Left            =   -165
            TabIndex        =   73
            Top             =   2340
            Width           =   945
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   255
            Left            =   390
            TabIndex        =   72
            Top             =   1980
            Width           =   945
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SubBlet"
            Height          =   255
            Left            =   -165
            TabIndex        =   70
            Top             =   1575
            Width           =   945
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Material"
            Height          =   255
            Left            =   -165
            TabIndex        =   69
            Top             =   1260
            Width           =   945
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Labour"
            Height          =   255
            Left            =   -165
            TabIndex        =   57
            Top             =   315
            Width           =   945
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Part"
            Height          =   255
            Left            =   -165
            TabIndex        =   56
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total"
            Height          =   255
            Left            =   375
            TabIndex        =   55
            Top             =   2730
            Width           =   945
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Oli"
            Height          =   255
            Left            =   -165
            TabIndex        =   54
            Top             =   945
            Width           =   945
         End
      End
      Begin VB.CommandButton TblKeluar 
         Caption         =   "&Kembali ke Menu"
         Height          =   390
         Left            =   9510
         TabIndex        =   19
         Top             =   5055
         Width           =   1890
      End
      Begin VB.Frame Frame5 
         Height          =   3150
         Left            =   45
         TabIndex        =   37
         Top             =   540
         Width           =   7890
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   32
            Left            =   3570
            TabIndex        =   5
            Top             =   525
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
            Left            =   360
            TabIndex        =   4
            Top             =   510
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
            Left            =   4275
            TabIndex        =   6
            Top             =   525
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
            Left            =   5895
            TabIndex        =   8
            Top             =   525
            Width           =   660
            _ExtentX        =   1164
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
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   270
            Left            =   4155
            TabIndex        =   39
            Top             =   525
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
            Left            =   3450
            TabIndex        =   40
            Top             =   525
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
            Left            =   180
            TabIndex        =   10
            Top             =   915
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   3572
            _Version        =   393216
            FixedCols       =   0
            HighLight       =   2
            ScrollBars      =   0
            SelectionMode   =   1
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   270
            Left            =   6570
            TabIndex        =   76
            Top             =   540
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
            Index           =   9
            Left            =   5085
            TabIndex        =   7
            Top             =   525
            Width           =   765
            _ExtentX        =   1349
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
            Index           =   10
            Left            =   6690
            TabIndex        =   9
            Top             =   525
            Width           =   660
            _ExtentX        =   1164
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
            Height          =   285
            Left            =   7365
            TabIndex        =   79
            Top             =   540
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   2670
            TabIndex        =   85
            Top             =   150
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Group"
            Height          =   255
            Left            =   6660
            TabIndex        =   80
            Top             =   300
            Width           =   570
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "A. Rate"
            Height          =   255
            Left            =   5190
            TabIndex        =   78
            Top             =   315
            Width           =   555
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mech"
            Height          =   255
            Left            =   5865
            TabIndex        =   44
            Top             =   315
            Width           =   570
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Rate"
            Height          =   255
            Left            =   4395
            TabIndex        =   43
            Top             =   315
            Width           =   405
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
            Height          =   255
            Left            =   1080
            TabIndex        =   42
            Top             =   285
            Width           =   915
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "JP"
            Height          =   255
            Left            =   3615
            TabIndex        =   41
            Top             =   300
            Width           =   270
         End
      End
      Begin VB.Frame Frame3 
         Height          =   630
         Left            =   75
         TabIndex        =   36
         Top             =   4365
         Width           =   11325
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   37
            Left            =   2070
            TabIndex        =   17
            Top             =   225
            Width           =   3540
            _ExtentX        =   6244
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
            TabIndex        =   18
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
            TabIndex        =   77
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Slip/Keterangan"
            Height          =   255
            Left            =   195
            TabIndex        =   68
            Top             =   285
            Width           =   1755
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
         TabIndex        =   58
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
         TabIndex        =   61
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
         TabIndex        =   62
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
         TabIndex        =   65
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
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         Height          =   255
         Left            =   465
         TabIndex        =   83
         Top             =   4155
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "Terbilang (Rupiah)"
         Height          =   270
         Left            =   120
         TabIndex        =   84
         Top             =   3795
         Width           =   1395
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. No."
         Height          =   255
         Left            =   9060
         TabIndex        =   67
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T SA."
         Height          =   255
         Left            =   7440
         TabIndex        =   66
         Top             =   285
         Width           =   945
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Left            =   0
         TabIndex        =   64
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. WO"
         Height          =   255
         Left            =   2280
         TabIndex        =   63
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MCH"
         Height          =   255
         Left            =   5835
         TabIndex        =   60
         Top             =   285
         Width           =   945
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Left            =   4200
         TabIndex        =   59
         Top             =   315
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   405
      TabIndex        =   28
      Top             =   420
      Width           =   11460
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   20
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
         Left            =   840
         TabIndex        =   21
         Top             =   540
         Width           =   4335
         _ExtentX        =   7646
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
         Left            =   840
         TabIndex        =   22
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
         Left            =   6150
         TabIndex        =   23
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
         Left            =   6150
         TabIndex        =   24
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
         Left            =   6150
         TabIndex        =   26
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
         Left            =   10740
         TabIndex        =   25
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
         Left            =   7860
         TabIndex        =   45
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
         Left            =   7965
         TabIndex        =   46
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
         Left            =   840
         TabIndex        =   47
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
         Left            =   6150
         TabIndex        =   49
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
         Left            =   5055
         TabIndex        =   50
         Top             =   1185
         Width           =   945
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
         Height          =   255
         Left            =   -240
         TabIndex        =   48
         Top             =   1215
         Width           =   945
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
         Height          =   255
         Left            =   5055
         TabIndex        =   34
         Top             =   870
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Left            =   5055
         TabIndex        =   33
         Top             =   540
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         Height          =   255
         Left            =   5055
         TabIndex        =   32
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   255
         Left            =   -255
         TabIndex        =   31
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   255
         Left            =   -255
         TabIndex        =   30
         Top             =   585
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   -255
         TabIndex        =   29
         Top             =   270
         Width           =   945
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1245
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
      FormatString    =   "Perubahan A.Rate, Group & Mekanik"
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
      Left            =   135
      TabIndex        =   27
      Top             =   150
      Width           =   945
   End
End
Attribute VB_Name = "frmEditRate"
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
    With dGrid
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
            Case 12, 15, 18, 20, 22, 23, 27, 29, 30, 36, 43, 44
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
        dGrid.TextMatrix(i, 1) = ""
        dGrid.TextMatrix(i, 2) = ""
        dGrid.TextMatrix(i, 3) = ""
        dGrid.TextMatrix(i, 4) = ""
        dGrid.TextMatrix(i, 5) = ""
        dGrid.TextMatrix(i, 6) = ""
        dGrid.TextMatrix(i, 7) = ""
    Next
End Sub

Private Sub cboStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        SendKeys ("{TAB}")
    End Select
End Sub

Private Sub DGrid_DblClick()
    If Baris > 1 And dGrid.TextMatrix(dGrid.Row, 1) <> "" Then
        A = dGrid.Row
        'AMBIL
        With dGrid
            Text(33) = .TextMatrix(A, 1)
            Text(32) = .TextMatrix(A, 2)
            Text(34) = .TextMatrix(A, 3)
            Text(9) = .TextMatrix(A, 4)
            Text(35) = .TextMatrix(A, 5)
            Text(10) = .TextMatrix(A, 6)
            Text(11) = .TextMatrix(A, 7)
        End With
        'HAPUS
        For i = A To Baris
            With dGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
                .TextMatrix(i, 5) = .TextMatrix(i + 1, 5)
                .TextMatrix(i, 6) = .TextMatrix(i + 1, 6)
                .TextMatrix(i, 7) = .TextMatrix(i + 1, 7)
            End With
        Next
        Baris = Baris - 1
        Text(33).SetFocus
'        Text(39) = TotalJasa
'        Call HitungTotal
    End If
End Sub

Private Sub dGrid_GotFocus()
    If Trim(Text(33)) <> "" Then
        If MaxBrs >= Baris Then
            With dGrid
                .TextMatrix(Baris, 1) = Text(33)
                .TextMatrix(Baris, 2) = Text(32)
                .TextMatrix(Baris, 3) = Format(Koma2Titik(Text(34)), "#,##0.00")
'                .TextMatrix(Baris, 4) = Format(Koma2Titik(Text(9)), "#,##0.00")
'                .TextMatrix(Baris, 3) = Koma2Titik(Text(34))
'                .TextMatrix(Baris, 4) = Koma2Titik(Text(34)) * RateCPI
                .TextMatrix(Baris, 4) = Format(Koma2Titik(Text(9)), "#,##0.00")
                .TextMatrix(Baris, 5) = Text(35)
                .TextMatrix(Baris, 6) = Text(10)
                .TextMatrix(Baris, 7) = Text(11)
            End With
            Baris = Baris + 1
            Text(9) = ""
            Text(33) = ""
            Text(32) = ""
            Text(34) = ""
            Text(35) = ""
            Text(10) = ""
            Text(11) = ""
            Text(33).SetFocus
'            Text(39) = TotalJasa
'            Call HitungTotal
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
    If KeyCode = vbKeyReturn Then
        Call DGrid_DblClick
    End If
End Sub

Private Sub Form_Load()
    Baris = 1
    dGrid.Rows = 11
    dGrid.Cols = 8
    dGrid.ColWidth(0) = 0
    dGrid.ColWidth(1) = 2950
    dGrid.ColWidth(2) = 800
    dGrid.ColWidth(3) = 1050
    dGrid.ColWidth(4) = 1050
    dGrid.ColWidth(5) = 800
    dGrid.ColWidth(6) = 800
    dGrid.ColWidth(7) = 0
    dGrid.ColAlignment(1) = flexAlignLeftCenter
    dGrid.ColAlignment(2) = flexAlignLeftCenter
    dGrid.ColAlignment(3) = flexAlignRightCenter
    dGrid.ColAlignment(4) = flexAlignRightCenter
    dGrid.ColAlignment(5) = flexAlignLeftCenter
    dGrid.ColAlignment(6) = flexAlignLeftCenter
    
    dGrid.TextMatrix(0, 1) = "Keterangan"
    dGrid.TextMatrix(0, 2) = "JP"
    dGrid.TextMatrix(0, 3) = "Rate"
    dGrid.TextMatrix(0, 4) = "A. Rate"
    dGrid.TextMatrix(0, 5) = "Mech"
    dGrid.TextMatrix(0, 6) = "Group"
    Call HapusForm
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub SaveOrderPekerjaan(NomorWO As String, NomorPolisi As String)
    Dim RCOrder As Boolean
    
    sSQL = "Select * from JOBDTA where NO_WO ='" & NomorWO & "' and DELIV <> 'BTL'"
    oRs.Open sSQL, oConn, adOpenStatic
    
    Call ClearsSql
    gsSql1(0) = "MECH"
    gsSql2(0) = Text(31)
    gsSql1(1) = "REP1"
    gsSql2(1) = dGrid.TextMatrix(1, 1)
    gsSql1(2) = "REP2"
    gsSql2(2) = dGrid.TextMatrix(2, 1)
    gsSql1(3) = "REP3"
    gsSql2(3) = dGrid.TextMatrix(3, 1)
    gsSql1(4) = "REP4"
    gsSql2(4) = dGrid.TextMatrix(4, 1)
    gsSql1(5) = "REP5"
    gsSql2(5) = dGrid.TextMatrix(5, 1)
    gsSql1(6) = "REP6"
    gsSql2(6) = dGrid.TextMatrix(6, 1)
    gsSql1(7) = "REP7"
    gsSql2(7) = dGrid.TextMatrix(7, 1)
    gsSql1(8) = "RTE1"
    gsSql2(8) = SimpanKoma2Titik(dGrid.TextMatrix(1, 3))
    gsSql1(9) = "RTE2"
    gsSql2(9) = SimpanKoma2Titik(dGrid.TextMatrix(2, 3))
    gsSql1(10) = "RTE3"
    gsSql2(10) = SimpanKoma2Titik(dGrid.TextMatrix(3, 3))
    gsSql1(11) = "RTE4"
    gsSql2(11) = SimpanKoma2Titik(dGrid.TextMatrix(4, 3))
    gsSql1(12) = "RTE5"
    gsSql2(12) = SimpanKoma2Titik(dGrid.TextMatrix(5, 3))
    gsSql1(13) = "RTE6"
    gsSql2(13) = SimpanKoma2Titik(dGrid.TextMatrix(6, 3))
    gsSql1(14) = "RTE7"
    gsSql2(14) = SimpanKoma2Titik(dGrid.TextMatrix(7, 3))
    gsSql1(15) = "KLPKERJA1"
    gsSql2(15) = dGrid.TextMatrix(1, 2)
    gsSql1(16) = "KLPKERJA2"
    gsSql2(16) = dGrid.TextMatrix(2, 2)
    gsSql1(17) = "KLPKERJA3"
    gsSql2(17) = dGrid.TextMatrix(3, 2)
    gsSql1(18) = "KLPKERJA4"
    gsSql2(18) = dGrid.TextMatrix(4, 2)
    gsSql1(19) = "KLPKERJA5"
    gsSql2(19) = dGrid.TextMatrix(5, 2)
    gsSql1(20) = "KLPKERJA6"
    gsSql2(20) = dGrid.TextMatrix(6, 2)
    gsSql1(21) = "KLPKERJA7"
    gsSql2(21) = dGrid.TextMatrix(7, 2)
    gsSql1(22) = "MEKANIK1"
    gsSql2(22) = dGrid.TextMatrix(1, 5)
    gsSql1(23) = "MEKANIK2"
    gsSql2(23) = dGrid.TextMatrix(2, 5)
    gsSql1(24) = "MEKANIK3"
    gsSql2(24) = dGrid.TextMatrix(3, 5)
    gsSql1(25) = "MEKANIK4"
    gsSql2(25) = dGrid.TextMatrix(4, 5)
    gsSql1(26) = "MEKANIK5"
    gsSql2(26) = dGrid.TextMatrix(5, 5)
    gsSql1(27) = "MEKANIK6"
    gsSql2(27) = dGrid.TextMatrix(6, 5)
    gsSql1(28) = "MEKANIK7"
    gsSql2(28) = dGrid.TextMatrix(7, 5)
        
    gsSql1(29) = "ACTRATE1"
    gsSql2(29) = SimpanKoma2Titik(dGrid.TextMatrix(1, 4))
    gsSql1(30) = "ACTRATE2"
    gsSql2(30) = SimpanKoma2Titik(dGrid.TextMatrix(2, 4))
    gsSql1(31) = "ACTRATE3"
    gsSql2(31) = SimpanKoma2Titik(dGrid.TextMatrix(3, 4))
    gsSql1(32) = "ACTRATE4"
    gsSql2(32) = SimpanKoma2Titik(dGrid.TextMatrix(4, 4))
    gsSql1(33) = "ACTRATE5"
    gsSql2(33) = SimpanKoma2Titik(dGrid.TextMatrix(5, 4))
    gsSql1(34) = "ACTRATE6"
    gsSql2(34) = SimpanKoma2Titik(dGrid.TextMatrix(6, 4))
    gsSql1(35) = "ACTRATE7"
    gsSql2(35) = SimpanKoma2Titik(dGrid.TextMatrix(7, 4))
        
    gsSql1(36) = "GRP1"
    gsSql2(36) = dGrid.TextMatrix(1, 6)
    gsSql1(37) = "GRP2"
    gsSql2(37) = dGrid.TextMatrix(2, 6)
    gsSql1(38) = "GRP3"
    gsSql2(38) = dGrid.TextMatrix(3, 6)
    gsSql1(39) = "GRP4"
    gsSql2(39) = dGrid.TextMatrix(4, 6)
    gsSql1(40) = "GRP5"
    gsSql2(40) = dGrid.TextMatrix(5, 6)
    gsSql1(41) = "GRP6"
    gsSql2(41) = dGrid.TextMatrix(6, 6)
    gsSql1(42) = "GRP7"
    gsSql2(42) = dGrid.TextMatrix(7, 6)
    gsSql1(43) = "STATUS"
    gsSql2(43) = cboStatus.Text
    gsSql1(44) = "FRONT"
    gsSql2(44) = Text(13)
    
    gsSql1(45) = "OFF1"
    gsSql2(45) = SimpanKoma2Titik(dGrid.TextMatrix(1, 7))
    gsSql1(46) = "OFF2"
    gsSql2(46) = SimpanKoma2Titik(dGrid.TextMatrix(2, 7))
    gsSql1(47) = "OFF3"
    gsSql2(47) = SimpanKoma2Titik(dGrid.TextMatrix(3, 7))
    gsSql1(48) = "OFF4"
    gsSql2(48) = SimpanKoma2Titik(dGrid.TextMatrix(4, 7))
    gsSql1(49) = "OFF5"
    gsSql2(49) = SimpanKoma2Titik(dGrid.TextMatrix(5, 7))
    gsSql1(50) = "OFF6"
    gsSql2(50) = SimpanKoma2Titik(dGrid.TextMatrix(6, 7))
    gsSql1(51) = "OFF7"
    gsSql2(51) = SimpanKoma2Titik(dGrid.TextMatrix(7, 7))
    If Not oRs.BOF Then
        sSQL = GensSql("JOBDTA", True) + " where NO_WO ='" & NomorWO & "'"
    End If
    oConn.Execute sSQL
        
    oRs.Close
    
    'Edit Invoice
    sSQL = "Select * from T_INVOICE where NO_WO ='" & NomorWO & "' and DELIV <> 'BTL'"
    oRs.Open sSQL, oConn, adOpenStatic
    
    Call ClearsSql
    gsSql1(0) = "MECH"
    gsSql2(0) = Text(31)
    gsSql1(1) = "REP1"
    gsSql2(1) = dGrid.TextMatrix(1, 1)
    gsSql1(2) = "REP2"
    gsSql2(2) = dGrid.TextMatrix(2, 1)
    gsSql1(3) = "REP3"
    gsSql2(3) = dGrid.TextMatrix(3, 1)
    gsSql1(4) = "REP4"
    gsSql2(4) = dGrid.TextMatrix(4, 1)
    gsSql1(5) = "REP5"
    gsSql2(5) = dGrid.TextMatrix(5, 1)
    gsSql1(6) = "REP6"
    gsSql2(6) = dGrid.TextMatrix(6, 1)
    gsSql1(7) = "REP7"
    gsSql2(7) = dGrid.TextMatrix(7, 1)
    gsSql1(8) = "RTE1"
    gsSql2(8) = SimpanKoma2Titik(dGrid.TextMatrix(1, 3))
    gsSql1(9) = "RTE2"
    gsSql2(9) = SimpanKoma2Titik(dGrid.TextMatrix(2, 3))
    gsSql1(10) = "RTE3"
    gsSql2(10) = SimpanKoma2Titik(dGrid.TextMatrix(3, 3))
    gsSql1(11) = "RTE4"
    gsSql2(11) = SimpanKoma2Titik(dGrid.TextMatrix(4, 3))
    gsSql1(12) = "RTE5"
    gsSql2(12) = SimpanKoma2Titik(dGrid.TextMatrix(5, 3))
    gsSql1(13) = "RTE6"
    gsSql2(13) = SimpanKoma2Titik(dGrid.TextMatrix(6, 3))
    gsSql1(14) = "RTE7"
    gsSql2(14) = SimpanKoma2Titik(dGrid.TextMatrix(7, 3))
    gsSql1(15) = "KLPKERJA1"
    gsSql2(15) = dGrid.TextMatrix(1, 2)
    gsSql1(16) = "KLPKERJA2"
    gsSql2(16) = dGrid.TextMatrix(2, 2)
    gsSql1(17) = "KLPKERJA3"
    gsSql2(17) = dGrid.TextMatrix(3, 2)
    gsSql1(18) = "KLPKERJA4"
    gsSql2(18) = dGrid.TextMatrix(4, 2)
    gsSql1(19) = "KLPKERJA5"
    gsSql2(19) = dGrid.TextMatrix(5, 2)
    gsSql1(20) = "KLPKERJA6"
    gsSql2(20) = dGrid.TextMatrix(6, 2)
    gsSql1(21) = "KLPKERJA7"
    gsSql2(21) = dGrid.TextMatrix(7, 2)
    gsSql1(22) = "MEKANIK1"
    gsSql2(22) = dGrid.TextMatrix(1, 5)
    gsSql1(23) = "MEKANIK2"
    gsSql2(23) = dGrid.TextMatrix(2, 5)
    gsSql1(24) = "MEKANIK3"
    gsSql2(24) = dGrid.TextMatrix(3, 5)
    gsSql1(25) = "MEKANIK4"
    gsSql2(25) = dGrid.TextMatrix(4, 5)
    gsSql1(26) = "MEKANIK5"
    gsSql2(26) = dGrid.TextMatrix(5, 5)
    gsSql1(27) = "MEKANIK6"
    gsSql2(27) = dGrid.TextMatrix(6, 5)
    gsSql1(28) = "MEKANIK7"
    gsSql2(28) = dGrid.TextMatrix(7, 5)
        
    gsSql1(29) = "ACTRATE1"
    gsSql2(29) = SimpanKoma2Titik(dGrid.TextMatrix(1, 4))
    gsSql1(30) = "ACTRATE2"
    gsSql2(30) = SimpanKoma2Titik(dGrid.TextMatrix(2, 4))
    gsSql1(31) = "ACTRATE3"
    gsSql2(31) = SimpanKoma2Titik(dGrid.TextMatrix(3, 4))
    gsSql1(32) = "ACTRATE4"
    gsSql2(32) = SimpanKoma2Titik(dGrid.TextMatrix(4, 4))
    gsSql1(33) = "ACTRATE5"
    gsSql2(33) = SimpanKoma2Titik(dGrid.TextMatrix(5, 4))
    gsSql1(34) = "ACTRATE6"
    gsSql2(34) = SimpanKoma2Titik(dGrid.TextMatrix(6, 4))
    gsSql1(35) = "ACTRATE7"
    gsSql2(35) = SimpanKoma2Titik(dGrid.TextMatrix(7, 4))
        
    gsSql1(36) = "GRP1"
    gsSql2(36) = dGrid.TextMatrix(1, 6)
    gsSql1(37) = "GRP2"
    gsSql2(37) = dGrid.TextMatrix(2, 6)
    gsSql1(38) = "GRP3"
    gsSql2(38) = dGrid.TextMatrix(3, 6)
    gsSql1(39) = "GRP4"
    gsSql2(39) = dGrid.TextMatrix(4, 6)
    gsSql1(40) = "GRP5"
    gsSql2(40) = dGrid.TextMatrix(5, 6)
    gsSql1(41) = "GRP6"
    gsSql2(41) = dGrid.TextMatrix(6, 6)
    gsSql1(42) = "GRP7"
    gsSql2(42) = dGrid.TextMatrix(7, 6)
    gsSql1(43) = "STATUS"
    gsSql2(43) = cboStatus.Text
    gsSql1(44) = "FRONT"
    gsSql2(44) = Text(13)
    gsSql1(45) = "OFF1"
    gsSql2(45) = SimpanKoma2Titik(dGrid.TextMatrix(1, 7))
    gsSql1(46) = "OFF2"
    gsSql2(46) = SimpanKoma2Titik(dGrid.TextMatrix(2, 7))
    gsSql1(47) = "OFF3"
    gsSql2(47) = SimpanKoma2Titik(dGrid.TextMatrix(3, 7))
    gsSql1(48) = "OFF4"
    gsSql2(48) = SimpanKoma2Titik(dGrid.TextMatrix(4, 7))
    gsSql1(49) = "OFF5"
    gsSql2(49) = SimpanKoma2Titik(dGrid.TextMatrix(5, 7))
    gsSql1(50) = "OFF6"
    gsSql2(50) = SimpanKoma2Titik(dGrid.TextMatrix(6, 7))
    gsSql1(51) = "OFF7"
    gsSql2(51) = SimpanKoma2Titik(dGrid.TextMatrix(7, 7))
    
    If Not oRs.BOF Then
        sSQL = GensSql("T_INVOICE", True) + " where NO_WO ='" & NomorWO & "'"
    End If
    oConn.Execute sSQL
        
    oRs.Close
    
    Call HapusForm
    Text(0) = ""
End Sub

Private Sub SimpanData()
    Dim Boleh As Boolean
    Dim Pesan As String
    
    Boleh = True
    
    If Trim(cboStatus.Text) = "" Then
        Pesan = Pesan + " Status"
        Boleh = False
    End If
    If Trim(Text(31)) = "" Then
        Pesan = Pesan + " MECH"
        Boleh = False
    End If
    
    If Boleh = True Then
        Call SaveOrderPekerjaan(Text(0), Text(16))
    Else
        MsgBox "Masukkan" + Pesan + "!"
    End If
End Sub

Private Sub TblSimpan_Click()
    Call SimpanData
    Text(0).SetFocus
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
        Case 9, 10, 35
            Call ItemRegular
    End Select

End Sub

Private Sub PanggilDataOrder(NomorWO As String)
    Dim Total As Double
    
    sSQL = "select * from JOBDTA where NO_WO ='" & NomorWO & "' and DELIV <> 'BTL'"
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
        Text(39) = oRs1("LABOR")
        Text(40) = oRs1("PART")
        Text(41) = oRs1("TOTAL")
        Text(42) = oRs1("NOINV")
        Text(45) = oRs1("OIL")
        cboStatus.Text = Trim(oRs1("STATUS"))
        Total = oRs1("TOTAL")
        Baris = 1
        
        For i = 1 To 7
            dGrid.TextMatrix(i, 1) = oRs1("REP" & Trim(Str(i)))
            dGrid.TextMatrix(i, 2) = oRs1("KLPKERJA" & Trim(Str(i)))
            dGrid.TextMatrix(i, 3) = Format(oRs1("RTE" & Trim(Str(i))), "#,##0.00")
            dGrid.TextMatrix(i, 4) = Format(oRs1("ACTRATE" & Trim(Str(i))), "#,##0.00")
            dGrid.TextMatrix(i, 5) = oRs1("MEKANIK" & Trim(Str(i)))
            dGrid.TextMatrix(i, 6) = oRs1("GRP" & Trim(Str(i)))
            dGrid.TextMatrix(i, 7) = oRs1("OFF" & Trim(Str(i)))
        Next
        
        
        
'        For i = 1 To 7
'            DGrid.TextMatrix(i, 1) = oRs1("REP" & Trim(Str(i)))
'            DGrid.TextMatrix(i, 2) = oRs1("KLPKERJA" & Trim(Str(i)))
'            DGrid.TextMatrix(i, 3) = Format(oRs1("RTE" & Trim(Str(i))), "#,##0.00")
''            DGrid.TextMatrix(i, 4) = Format(oRs1("OFF" & Trim(Str(i))), "#,##0.00")
'            DGrid.TextMatrix(i, 4) = Format(oRs1("ACTRATE" & Trim(Str(i))), "#,##0.00")
'            DGrid.TextMatrix(i, 5) = oRs1("MEKANIK" & Trim(Str(i)))
'        Next
        
        For i = 1 To 7
            If dGrid.TextMatrix(i, 1) <> "" Then
                Baris = Baris + 1
            End If
        Next

    End If
    oRs1.Close
        
    sSQL = "select * from MODEL where MDL ='" & Text(17) & "'"
    oRs1.Open sSQL, oConn, adOpenStatic
    If oRs1.BOF Then
        gGroupModel = ""
    Else
        gGroupModel = oRs1("FS_GRP_MDL")
    End If
    oRs1.Close
            
    sSQL = "select * from T_INVOICE where NO_WO ='" & NomorWO & "' and NO_POLISI ='" & Trim(Text(16)) & "' and (DELIV='PRC' or DELIV='OKE') and FS_NO_RETUR = ''"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        Text(38) = oRs1("FS_CATATAN")
    End If
    oRs1.Close
    Text1.Text = Angka2Huruf(Total, 1)

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "EDITRATE"
                frmCari.Show (vbModal)
            Case 10
                gMainKeyCari = Text(Index)
                gSearchOption = "GROUP_MEKANIK"
                frmCari.Show (vbModal)
            Case 17
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.Show (vbModal)
            Case 31, 35
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
        Case 9, 34, 40, 4, 5, 45, 7
            KeyAscii = CekNumerik(KeyAscii)
    End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Dim NoWO As String
    
    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                TblSimpan.Enabled = True
                NoWO = Text(0)
                sSQL = "select * from JOBDTA where NO_WO='" & NoWO & "' and DELIV <> 'BTL'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "EDITRATE"
                    frmCari.Show (vbModal)
                Else
                    Call PanggilDataOrder(NoWO)
                End If
                gMDL = Text(17)
                
                oRs.Close
            Else
                Call HapusForm
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
        
        Case 35
            If Trim(Text(Index)) <> "" Then
                sSQL = "Select * from MEKANIK where MECH ='" & Text(Index) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MEKANIK"
                    frmCari.Show (vbModal)
                Else
                    Text(10) = oRs("GROUP")
                End If
                oRs.Close
            Else
                Text(Index) = Text(31)
                sSQL = "Select * from MEKANIK where MECH ='" & Text(Index) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If Not oRs.BOF Then
                    Text(10) = oRs("GROUP")
                End If
                oRs.Close
            End If
    End Select
End Sub

Private Sub Text1_GotFocus()
    Dim DiTab As Boolean
    DiTab = False
    If Text1.BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text1.Text = gMainKeyCari
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

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub
